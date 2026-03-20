"""
EML to PST Converter
A GUI application to convert EML/EMLX files to Outlook PST format.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import glob
import hashlib
from email import policy
from email.parser import BytesParser
from datetime import datetime
import threading
import subprocess
import sys
import time
import re
import gc
import logging
import tempfile
import uuid
import atexit
from pathlib import Path
from email.utils import parsedate_to_datetime

from i18n import (
    LANG_EN,
    LANG_HU,
    LANGUAGE_NAMES,
    detect_default_lang,
    lang_from_display,
    t,
)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Outlook constants
OL_MAIL_ITEM = 0
OL_DISCARD = 1
OL_STORE_UNICODE = 2  # Unicode PST format (Outlook 2003+)

# Folder names
INBOX_FOLDER_NAME = "Inbox"

# Check for pywin32
OUTLOOK_AVAILABLE = False
WIN32COM = None
PYTHONCOM = None
PYWINTYPES = None
try:
    import win32com.client
    import pythoncom
    WIN32COM = win32com.client
    PYTHONCOM = pythoncom
    OUTLOOK_AVAILABLE = True
    try:
        import pywintypes as _pywintypes
        PYWINTYPES = _pywintypes
    except ImportError:
        pass
except ImportError:
    logger.warning("pywin32 not available - PST conversion will require installation")

# Maximum files to process (memory safety)
MAX_FILES = 50000
def _env_int_mb(name: str, default: int = 100, *, lo: int = 1, hi: int = 2048) -> int:
    try:
        v = int(os.environ.get(name, str(default)).strip())
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


# Single-message read cap (mitigates huge / malicious .eml memory use)
MAX_EML_FILE_BYTES = _env_int_mb("EML2PST_MAX_FILE_MB", 100) * 1024 * 1024


def _env_float(name: str, default: float, *, lo: float = 0.0, hi: float = 600.0) -> float:
    """Parse optional float env; clamp to [lo, hi] to avoid abuse or typos."""
    try:
        v = float(os.environ.get(name, "").strip())
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


def _safe_staging_subdir(raw: str | None, default: str = "EML2PST_Staging") -> str:
    """
    Single folder name only — env must not inject path traversal or drive letters.
    """
    s = (raw or default).strip()
    if len(s) > 80:
        s = s[:80]
    if not s or s in (".", ".."):
        return default
    for bad in (os.sep, ".."):
        if bad in s:
            return default
    if os.altsep and os.altsep in s:
        return default
    if ":" in s:  # Windows streams / drive-relative tricks
        return default
    if not re.match(r"^[\w.\- ]+$", s):
        return default
    return s


# After native import, Outlook may still read the staging .eml from disk asynchronously.
# Deleting it immediately causes "file may have been moved or deleted" (localized in HU/EN).
# Staging files are kept until the batch ends; optional extra delay after COM release.
_NATIVE_POST_BATCH_DELAY = _env_float("EML2PST_POST_BATCH_DELAY", 3.0)
# Not a dot-folder: Outlook often fails to open file:/// URLs under ".something" paths.
NATIVE_STAGING_SUBDIR = _safe_staging_subdir(os.environ.get("EML2PST_STAGING_SUBDIR"))


class EmlToPstConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("EML to PST Converter")
        # Default / min size: long HU/EN labels need room; bottom bar must stay visible.
        self.root.geometry("880x720")
        self.root.minsize(780, 640)
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.folder_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.pst_option = tk.StringVar(value="new")
        self.remove_duplicates = tk.BooleanVar(value=False)
        # Off by default: use manual fallback + Windows file mtime for time when checked.
        self.strict_date_preservation = tk.BooleanVar(value=False)
        # Explorer "Date modified" / "Módosítás dátuma" on the original .eml file.
        self.use_file_mtime_for_date = tk.BooleanVar(value=True)
        self.lang_var = tk.StringVar(value=LANGUAGE_NAMES[detect_default_lang()])
        self.file_pattern = tk.StringVar(value="*.eml")
        self.eml_files = []
        self.processed_hashes = set()
        
        # Thread safety
        self._lock = threading.Lock()
        self._is_converting = False
        self._temp_files = []
        self._conversion_options = {}
        # Native OpenSharedItem staging: paths kept until batch finishes (Outlook async read).
        self._native_staging_paths = []
        self._staging_dir = None  # set per run; folder next to target PST
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        atexit.register(self._cleanup_temp_files)
        
        self.create_widgets()
        self.apply_language()

    def _current_lang(self) -> str:
        return lang_from_display(self.lang_var.get())

    def apply_language(self, _event=None):
        """Refresh all UI strings (English / Hungarian)."""
        lang = self._current_lang()
        self.root.title(t(lang, "window_title"))

        self._lf_folder.config(text=t(lang, "folder_section"))
        self._lf_pst.config(text=t(lang, "save_pst_section"))
        self._lf_list.config(text=t(lang, "eml_files_section"))
        self._btn_add.config(text=t(lang, "add_files"))
        self._lbl_pattern.config(text=t(lang, "file_pattern_label"))
        self._lbl_pattern_hint.config(text=t(lang, "file_pattern_hint"))
        self._rb_new.config(text=t(lang, "create_new_pst"))
        self._rb_existing.config(text=t(lang, "save_existing_pst"))
        self._cb_dup.config(text=t(lang, "remove_duplicates"))
        self._cb_strict.config(text=t(lang, "strict_date_preservation"))
        explorer = t(lang, "explorer_date_modified")
        self._cb_mtime.config(text=t(lang, "use_file_mtime", explorer=explorer))
        self.file_tree.heading("name", text=t(lang, "col_name"))
        self.file_tree.heading("path", text=t(lang, "col_path"))
        self.file_tree.heading("size", text=t(lang, "col_size"))
        self.file_tree.heading("date", text=t(lang, "col_date"))
        self.update_file_count()
        self._lbl_dest.config(text=t(lang, "destination_label"))
        self._btn_browse_dest.config(text=t(lang, "browse_destination"))
        self._btn_exit.config(text=t(lang, "btn_exit"))
        self._btn_convert.config(text=t(lang, "btn_convert"))
        self._lbl_lang.config(text=t(lang, "language_label"))

        try:
            self.context_menu.entryconfig(0, label=t(lang, "context_remove"))
            self.context_menu.entryconfig(1, label=t(lang, "context_clear_all"))
        except tk.TclError:
            pass

        with self._lock:
            if not self._is_converting:
                self.status_label.config(text=t(lang, "status_ready"))

    def create_widgets(self):
        # Main frame: grid so the file list absorbs shrink/grow; dest + buttons stay visible.
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        # Row 3 = file list label frame — only this row expands vertically.
        main_frame.grid_rowconfigure(3, weight=1)

        # Language (Nyelv)
        lang_row = ttk.Frame(main_frame)
        lang_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._lbl_lang = ttk.Label(lang_row, text="")
        self._lbl_lang.pack(side=tk.LEFT)
        lang_combo = ttk.Combobox(
            lang_row,
            textvariable=self.lang_var,
            values=(LANGUAGE_NAMES[LANG_EN], LANGUAGE_NAMES[LANG_HU]),
            state="readonly",
            width=12,
        )
        lang_combo.pack(side=tk.LEFT, padx=(6, 0))
        lang_combo.bind("<<ComboboxSelected>>", self.apply_language)
        
        # === Add Folder Section ===
        self._lf_folder = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_folder.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        folder_frame = self._lf_folder
        
        # Folder path entry and browse
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.folder_entry = ttk.Entry(path_frame, textvariable=self.folder_path, width=70)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self._btn_add = ttk.Button(path_frame, text="", command=self.browse_folder)
        self._btn_add.pack(side=tk.RIGHT)
        
        # Wildcard pattern frame
        pattern_frame = ttk.Frame(folder_frame)
        pattern_frame.pack(fill=tk.X, pady=(5, 0))
        
        self._lbl_pattern = ttk.Label(pattern_frame, text="")
        self._lbl_pattern.pack(side=tk.LEFT)
        
        pattern_combo = ttk.Combobox(pattern_frame, textvariable=self.file_pattern, width=15)
        pattern_combo['values'] = ('*.eml', '*.emlx', '*.eml;*.emlx')
        pattern_combo.pack(side=tk.LEFT, padx=(5, 10))
        
        self._lbl_pattern_hint = ttk.Label(pattern_frame, text="")
        self._lbl_pattern_hint.pack(side=tk.LEFT)
        
        # === Save in PST Section ===
        self._lf_pst = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_pst.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        pst_frame = self._lf_pst
        
        options_frame = ttk.Frame(pst_frame)
        options_frame.pack(fill=tk.X)
        
        self._rb_new = ttk.Radiobutton(
            options_frame, text="", variable=self.pst_option, value="new"
        )
        self._rb_new.pack(side=tk.LEFT, padx=(0, 20))
        self._rb_existing = ttk.Radiobutton(
            options_frame, text="", variable=self.pst_option, value="existing"
        )
        self._rb_existing.pack(side=tk.LEFT, padx=(0, 20))
        self._cb_dup = ttk.Checkbutton(
            options_frame, text="", variable=self.remove_duplicates
        )
        self._cb_dup.pack(side=tk.LEFT)
        self._cb_strict = ttk.Checkbutton(
            options_frame,
            text="",
            variable=self.strict_date_preservation,
        )
        self._cb_strict.pack(side=tk.LEFT, padx=(20, 0))

        pst_options_row2 = ttk.Frame(pst_frame)
        pst_options_row2.pack(fill=tk.X, pady=(6, 0))
        self._cb_mtime = ttk.Checkbutton(
            pst_options_row2,
            text="",
            variable=self.use_file_mtime_for_date,
        )
        self._cb_mtime.pack(side=tk.LEFT)
        
        # === File List Section ===
        self._lf_list = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_list.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        list_frame = self._lf_list
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Create Treeview with scrollbars
        tree_container = ttk.Frame(list_frame)
        tree_container.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        y_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        y_scroll.grid(row=0, column=1, sticky="ns")
        
        x_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        # Treeview (height= rows; grid gives remaining space so list scales with window)
        self.file_tree = ttk.Treeview(tree_container, columns=("name", "path", "size", "date"),
                                       show="headings", height=8,
                                       yscrollcommand=y_scroll.set,
                                       xscrollcommand=x_scroll.set)
        
        self.file_tree.heading("name", text="")
        self.file_tree.heading("path", text="")
        self.file_tree.heading("size", text="")
        self.file_tree.heading("date", text="")
        
        self.file_tree.column("name", width=200, minwidth=150)
        self.file_tree.column("path", width=300, minwidth=200)
        self.file_tree.column("size", width=80, minwidth=60)
        self.file_tree.column("date", width=120, minwidth=100)
        
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        
        y_scroll.config(command=self.file_tree.yview)
        x_scroll.config(command=self.file_tree.xview)
        
        # File count label
        self.file_count_label = ttk.Label(list_frame, text="")
        self.file_count_label.grid(row=1, column=0, sticky="w", pady=(5, 0))
        
        # Context menu for treeview
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="", command=self.remove_selected)
        self.context_menu.add_command(label="", command=self.clear_all)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # === Destination Section ===
        dest_frame = ttk.Frame(main_frame)
        dest_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        
        self._lbl_dest = ttk.Label(dest_frame, text="")
        self._lbl_dest.pack(side=tk.LEFT)
        
        self.dest_entry = ttk.Entry(dest_frame, textvariable=self.destination_path, width=60)
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        self._btn_browse_dest = ttk.Button(
            dest_frame, text="", command=self.browse_destination
        )
        self._btn_browse_dest.pack(side=tk.RIGHT)
        
        # === Bottom Buttons ===
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(button_frame, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(button_frame, text="")
        self.status_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # Convert and Exit buttons
        self._btn_exit = ttk.Button(
            button_frame, text="", command=self._on_close, width=10
        )
        self._btn_exit.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.convert_btn = ttk.Button(
            button_frame, text="", command=self.start_conversion, width=10
        )
        self.convert_btn.pack(side=tk.RIGHT)
        self._btn_convert = self.convert_btn
        
    def browse_folder(self):
        """Browse for folder containing EML/EMLX files"""
        folder = filedialog.askdirectory(
            title=t(self._current_lang(), "dialog_select_folder")
        )
        if folder:
            self.folder_path.set(folder)
            lang = self._current_lang()
            # Snapshot pattern on UI thread — Tk StringVar must not be read from workers.
            pattern_snapshot = self.file_pattern.get()
            thread = threading.Thread(
                target=self._scan_folder_thread,
                args=(folder, lang, pattern_snapshot),
            )
            thread.daemon = True
            thread.start()
    
    def _scan_folder_thread(self, folder, lang, pattern_snapshot):
        """Background thread for folder scanning"""
        self.root.after(
            0,
            lambda: self.status_label.config(text=t(lang, "status_scanning")),
        )
        
        try:
            files = self._scan_folder_impl(folder, lang, pattern_snapshot)
            self.root.after(0, lambda: self._add_files_to_list(files, lang))
        except Exception as e:
            logger.error("Error scanning folder: %s", e)
            self.root.after(
                0,
                lambda err=str(e), lg=lang: messagebox.showerror(
                    t(lg, "title_error"),
                    t(lg, "err_scan", err=err),
                ),
            )
            
    def _scan_folder_impl(self, folder, lang, pattern_input):
        """Scan folder for EML/EMLX files using wildcard pattern"""
        # Validate pattern - only allow safe patterns
        if not self._is_valid_pattern(pattern_input):
            raise ValueError(t(lang, "err_invalid_pattern"))
        
        patterns = pattern_input.split(';')
        found_files = []
        
        for pattern in patterns:
            pattern = pattern.strip()
            if not pattern:
                continue
            # Search recursively
            search_path = os.path.join(folder, '**', pattern)
            files = glob.glob(search_path, recursive=True)
            found_files.extend(files)
            
            # Safety limit
            if len(found_files) > MAX_FILES:
                logger.warning("File limit reached (%d)", MAX_FILES)
                break
        
        # Remove duplicates while preserving order
        seen = set()
        unique_files = []
        for f in found_files:
            normalized = os.path.normpath(f)
            if normalized not in seen:
                seen.add(normalized)
                unique_files.append(normalized)
                
        return unique_files[:MAX_FILES]
    
    def _is_valid_pattern(self, pattern_input):
        """Validate file pattern to prevent dangerous patterns"""
        allowed_patterns = re.compile(r'^[\*\?a-zA-Z0-9_\-\.;]+$')
        if not allowed_patterns.match(pattern_input):
            return False
        
        # Must contain .eml or .emlx extension
        patterns = pattern_input.lower().split(';')
        for p in patterns:
            p = p.strip()
            if p and not (p.endswith('.eml') or p.endswith('.emlx')):
                return False
        return True
    
    def _add_files_to_list(self, files, lang=None):
        """Add files to the treeview (runs on main thread)"""
        if lang is None:
            lang = self._current_lang()
        for file_path in files:
            self.add_file_to_list(file_path)
            
        self.update_file_count()
        self.status_label.config(text=t(lang, "status_found", n=len(files)))
            
    def add_file_to_list(self, file_path):
        """Add a file to the treeview list"""
        with self._lock:
            if file_path in self.eml_files:
                return
            self.eml_files.append(file_path)
        
        filename = os.path.basename(file_path)
        folder = os.path.dirname(file_path)
        
        try:
            size = os.path.getsize(file_path)
            size_str = self.format_size(size)
            mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            date_str = mod_time.strftime("%Y-%m-%d %H:%M")
        except OSError as e:
            logger.warning("Could not get file info for %s: %s", file_path, e)
            size_str = "N/A"
            date_str = "N/A"
            
        self.file_tree.insert("", tk.END, values=(filename, folder, size_str, date_str))
        
    def format_size(self, size):
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"
        
    def update_file_count(self):
        """Update the file count label"""
        count = len(self.file_tree.get_children())
        lang = self._current_lang()
        self.file_count_label.config(text=t(lang, "total_files", n=count))
        
    def show_context_menu(self, event):
        """Show context menu on right-click"""
        self.context_menu.post(event.x_root, event.y_root)
        
    def remove_selected(self):
        """Remove selected items from the list"""
        selected = self.file_tree.selection()
        for item in selected:
            values = self.file_tree.item(item)['values']
            if values:
                file_path = os.path.join(str(values[1]), str(values[0]))
                with self._lock:
                    if file_path in self.eml_files:
                        self.eml_files.remove(file_path)
            self.file_tree.delete(item)
        self.update_file_count()
        
    def clear_all(self):
        """Clear all items from the list"""
        self.file_tree.delete(*self.file_tree.get_children())
        with self._lock:
            self.eml_files.clear()
        self.update_file_count()
        
    def browse_destination(self):
        """Browse for destination PST file"""
        lang = self._current_lang()
        pst_type = (t(lang, "dialog_filetype_pst"), "*.pst")
        all_type = (t(lang, "dialog_all_files"), "*.*")
        if self.pst_option.get() == "new":
            file_path = filedialog.asksaveasfilename(
                title=t(lang, "dialog_save_pst"),
                defaultextension=".pst",
                filetypes=[pst_type, all_type],
            )
        else:
            file_path = filedialog.askopenfilename(
                title=t(lang, "dialog_open_pst"),
                filetypes=[pst_type, all_type],
            )
        
        if file_path:
            self.destination_path.set(file_path)
            
    def get_email_hash(self, file_path):
        """Generate hash for duplicate detection using SHA256"""
        try:
            sha256 = hashlib.sha256()
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(8192), b''):
                    sha256.update(chunk)
            return sha256.hexdigest()
        except OSError as e:
            logger.warning("Could not hash file %s: %s", file_path, e)
            return None
            
    def parse_eml(self, file_path):
        """
        Parse an EML file and return email data.
        Uses the stdlib parser so all MIME headers (Date, Received chain, etc.)
        stay available for transport-header and date preservation in Outlook.
        """
        try:
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            
            return {
                'subject': msg.get('Subject', '(No Subject)'),
                'from': msg.get('From', ''),
                'to': msg.get('To', ''),
                'cc': msg.get('Cc', ''),
                'date': msg.get('Date', ''),
                'body': self.get_email_body(msg),
                'attachments': self.get_attachments(msg),
                'message': msg
            }
        except Exception as e:
            logger.error("Error parsing EML file %s: %s", file_path, e)
            return None
            
    def get_email_body(self, msg):
        """Extract email body from message"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    try:
                        body = part.get_content()
                        break
                    except (KeyError, LookupError, UnicodeDecodeError) as e:
                        logger.debug("Could not decode text/plain part: %s", e)
                elif content_type == "text/html" and not body:
                    try:
                        body = part.get_content()
                    except (KeyError, LookupError, UnicodeDecodeError) as e:
                        logger.debug("Could not decode text/html part: %s", e)
        else:
            try:
                body = msg.get_content()
            except (KeyError, LookupError, UnicodeDecodeError):
                payload = msg.get_payload(decode=True)
                body = payload.decode('utf-8', errors='replace') if payload else ""
        return body
        
    def get_attachments(self, msg):
        """Extract attachments from message"""
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        # Sanitize filename to prevent path traversal
                        safe_filename = self._sanitize_filename(filename)
                        attachments.append({
                            'filename': safe_filename,
                            'data': part.get_payload(decode=True),
                            'content_type': part.get_content_type()
                        })
        return attachments
    
    def _sanitize_filename(self, filename):
        """Sanitize filename to prevent path traversal attacks"""
        # Remove any path components
        filename = os.path.basename(filename)
        # Remove potentially dangerous characters
        filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        # Ensure non-empty
        if not filename:
            filename = "attachment"
        # Limit length
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
        return filename
        
    def _on_close(self):
        """Handle window close with proper cleanup"""
        lang = self._current_lang()
        with self._lock:
            if self._is_converting:
                if not messagebox.askyesno(
                    t(lang, "confirm_exit_title"),
                    t(lang, "confirm_exit_msg"),
                ):
                    return
        self._cleanup_temp_files()
        self.root.destroy()
    
    def _cleanup_temp_files(self):
        """Remove any leftover temp files created during attachment handling"""
        self._cleanup_native_staging_files()
        with self._lock:
            paths = list(self._temp_files)
            self._temp_files.clear()
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError as e:
                logger.debug("Could not remove temp file %s: %s", path, e)

    def _ensure_native_staging_dir(self, pst_path: str) -> str:
        """Folder beside PST for staging (non-hidden name; Outlook-friendly)."""
        base = os.path.dirname(os.path.abspath(pst_path))
        if not base:
            base = os.getcwd()
        staging = os.path.join(base, NATIVE_STAGING_SUBDIR)
        os.makedirs(staging, exist_ok=True)
        return staging

    def _native_temp_staging_dirs(self, source_file_path: str) -> list:
        """
        Directories to try for normalized .eml staging, in order:
        1) Beside the source .eml (same access as the file user added)
        2) Beside the target PST
        3) System temp
        """
        dirs = []
        src_root = os.path.dirname(os.path.abspath(source_file_path))
        if src_root and os.path.isdir(src_root):
            d = os.path.join(src_root, NATIVE_STAGING_SUBDIR)
            try:
                os.makedirs(d, exist_ok=True)
                dirs.append(d)
            except OSError as e:
                logger.debug("Could not create source staging dir %s: %s", d, e)
        if self._staging_dir and os.path.isdir(self._staging_dir):
            if self._staging_dir not in dirs:
                dirs.append(self._staging_dir)
        td = tempfile.gettempdir()
        if td not in dirs:
            dirs.append(td)
        return dirs

    def _purge_orphan_staging_eml(self, staging_dir: str) -> None:
        """Remove leftover emlx_as_eml_*.eml from interrupted prior runs."""
        try:
            for name in os.listdir(staging_dir):
                if name.startswith("emlx_as_eml_") and name.endswith(".eml"):
                    p = os.path.join(staging_dir, name)
                    try:
                        os.remove(p)
                    except OSError:
                        pass
        except OSError:
            pass

    def _cleanup_native_staging_files(self) -> None:
        """Delete deferred native-import staging files; drop from _temp_files."""
        with self._lock:
            paths = list(self._native_staging_paths)
            self._native_staging_paths.clear()
        for p in paths:
            try:
                if os.path.isfile(p):
                    os.remove(p)
            except OSError as e:
                logger.debug("Could not remove native staging file %s: %s", p, e)
            with self._lock:
                if p in self._temp_files:
                    self._temp_files.remove(p)
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        lang = self._current_lang()
        with self._lock:
            if self._is_converting:
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_in_progress")
                )
                return
            if not self.eml_files:
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_no_files")
                )
                return
            # Snapshot UI state on the main thread; worker threads must not read Tk variables.
            self._conversion_options = {
                "destination_path": os.path.abspath(self.destination_path.get().strip()),
                "remove_duplicates": bool(self.remove_duplicates.get()),
                "strict_date_preservation": bool(self.strict_date_preservation.get()),
                "use_file_mtime_for_date": bool(self.use_file_mtime_for_date.get()),
                "pst_option": self.pst_option.get(),
                "lang": lang,
            }
            self._is_converting = True
            
        if not self._conversion_options["destination_path"]:
            with self._lock:
                self._is_converting = False
            messagebox.showwarning(
                t(lang, "title_warning"), t(lang, "warn_no_destination")
            )
            return
        
        dest_path = self._conversion_options["destination_path"]
        dest_dir = os.path.dirname(os.path.abspath(dest_path))
        if not os.path.isdir(dest_dir):
            with self._lock:
                self._is_converting = False
            messagebox.showerror(
                t(lang, "title_error"),
                t(lang, "err_dest_dir", path=dest_dir),
            )
            return
            
        self.convert_btn.config(state='disabled')
        
        thread = threading.Thread(target=self._convert_files_thread, args=(self._conversion_options.copy(),))
        thread.daemon = True
        thread.start()
    
    def _convert_files_thread(self, conversion_options):
        """Thread wrapper for conversion with COM initialization"""
        com_initialized = False
        try:
            if PYTHONCOM:
                PYTHONCOM.CoInitialize()
                com_initialized = True
            self.convert_files(conversion_options)
        finally:
            if com_initialized:
                PYTHONCOM.CoUninitialize()
            with self._lock:
                self._is_converting = False
            self.root.after(0, lambda: self.convert_btn.config(state='normal'))
        
    def convert_files(self, conversion_options):
        """Convert EML files to PST format"""
        with self._lock:
            total = len(self.eml_files)
            files_to_process = self.eml_files.copy()
            
        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
        with self._lock:
            self.processed_hashes.clear()
        
        # Check if pywin32 is available
        if not OUTLOOK_AVAILABLE:
            self.root.after(0, self.prompt_install_pywin32)
            return
        
        # Try to use Outlook COM for PST creation
        try:
            self.convert_with_outlook(files_to_process, conversion_options)
        except Exception as e:
            error_msg = str(e)
            logger.error("Conversion error: %s", error_msg)
            lg = conversion_options.get("lang", LANG_EN)

            def _show_conv_err(err=error_msg, lang=lg):
                messagebox.showerror(
                    t(lang, "title_conversion_error"),
                    t(lang, "msg_conversion_error", err=err),
                )

            self.root.after(0, _show_conv_err)
    
    def prompt_install_pywin32(self):
        """Prompt user to install pywin32"""
        lang = self._current_lang()
        result = messagebox.askyesno(
            t(lang, "title_missing_dep"),
            t(lang, "msg_missing_dep"),
        )
        
        if result:
            self.status_label.config(text=t(lang, "status_installing_pywin32"))
            self.root.update()
            
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "pywin32"],
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )
                messagebox.showinfo(
                    t(lang, "title_install_ok"),
                    t(lang, "msg_install_ok"),
                )
            except subprocess.CalledProcessError as e:
                logger.error("Failed to install pywin32: %s", e)
                messagebox.showerror(
                    t(lang, "title_install_fail"),
                    t(lang, "msg_install_fail"),
                )
            
            self.status_label.config(text=t(lang, "status_ready"))
            
    def convert_with_outlook(self, files_to_process, conversion_options):
        """Convert using Outlook COM interface (requires Microsoft Outlook)"""
        lang = conversion_options.get("lang", LANG_EN)
        pst_path = conversion_options["destination_path"]
        # Staging beside PST before COM (avoids stale dir; Outlook may block %TEMP%).
        self._staging_dir = self._ensure_native_staging_dir(pst_path)
        self._purge_orphan_staging_eml(self._staging_dir)
        # Orphans beside source folders (same subdir name as beside PST).
        seen_staging = {self._staging_dir}
        for fp in files_to_process:
            root = os.path.dirname(os.path.abspath(fp))
            sd = os.path.join(root, NATIVE_STAGING_SUBDIR)
            if sd in seen_staging:
                continue
            seen_staging.add(sd)
            if os.path.isdir(sd):
                self._purge_orphan_staging_eml(sd)
        with self._lock:
            self._native_staging_paths.clear()

        self._update_status(t(lang, "status_connecting"))
        
        outlook = None
        namespace = None
        try:
            try:
                outlook = WIN32COM.Dispatch("Outlook.Application")
            except Exception as e:
                raise RuntimeError(f"Could not connect to Outlook: {e}") from e
            
            try:
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                raise RuntimeError(f"Could not access MAPI namespace: {e}") from e
            
            self._cleanup_stale_store_for_path(namespace, pst_path)
            
            self._update_status(t(lang, "status_creating_pst"))
            self._setup_pst_store(outlook, namespace, pst_path, conversion_options)
            
            time.sleep(1)
                
            pst_store = self._find_pst_store(namespace, pst_path)
            if not pst_store:
                raise RuntimeError(f"Could not access PST file after creation. Path: {pst_path}")
                
            root_folder = pst_store.GetRootFolder()
            target_folder = self._get_or_create_inbox(root_folder)
                
            total = len(files_to_process)
            converted, skipped, errors, error_messages, skipped_messages = self._process_email_files(
                outlook, namespace, target_folder, files_to_process, total, conversion_options
            )
            
            note = t(lang, "note_pst_saved", path=pst_path)
            if error_messages:
                note += t(lang, "note_errors") + "\n".join(error_messages[:5])
                if len(error_messages) > 5:
                    note += t(
                        lang,
                        "note_more_errors",
                        n=len(error_messages) - 5,
                    )
            if skipped_messages:
                note += t(lang, "note_skipped") + "\n".join(skipped_messages[:5])
                if len(skipped_messages) > 5:
                    note += t(
                        lang,
                        "note_more_skipped",
                        n=len(skipped_messages) - 5,
                    )
                
            self.root.after(
                0,
                lambda c=converted, s=skipped, e=errors, n=note, lg=lang: self.show_completion(
                    c, s, e, n, lg
                ),
            )
        finally:
            # Release COM objects to prevent Outlook from hanging
            try:
                del namespace
            except Exception:
                pass
            try:
                del outlook
            except Exception:
                pass
            gc.collect()
            time.sleep(max(0.0, _NATIVE_POST_BATCH_DELAY))
            self._cleanup_native_staging_files()
    
    def _update_status(self, text):
        """Thread-safe status update"""
        self.root.after(0, lambda: self.status_label.config(text=text))
    
    def _cleanup_stale_store_for_path(self, namespace, pst_path):
        """Remove stale store reference only for the target PST path."""
        try:
            target = pst_path.lower()
            stale_stores = []
            for store in namespace.Stores:
                try:
                    file_path = store.FilePath
                    if file_path and file_path.lower() == target and not os.path.exists(file_path):
                        stale_stores.append((store, file_path))
                except (AttributeError, OSError):
                    continue
            
            # Remove stale stores
            for store, file_path in stale_stores:
                try:
                    root = store.GetRootFolder()
                    namespace.RemoveStore(root)
                    logger.info("Removed stale store reference: %s", file_path)
                except Exception as e:
                    logger.warning("Could not remove stale store %s: %s", file_path, e)
                    
        except Exception as e:
            logger.debug("Target stale store cleanup failed: %s", e)
    
    def _setup_pst_store(self, outlook, namespace, pst_path, conversion_options):
        """Create or open PST store"""
        try:
            # First, remove any existing store reference with this path
            self._remove_existing_store(namespace, pst_path)
            
            if conversion_options.get("pst_option") == "new":
                # Remove existing file if present
                if os.path.exists(pst_path):
                    try:
                        os.remove(pst_path)
                    except OSError as e:
                        logger.warning("Could not remove existing PST: %s", e)
                
                # Create PST using AddStoreEx (creates Unicode PST)
                self._create_new_pst(outlook, namespace, pst_path)
            else:
                # Open existing PST
                if not os.path.exists(pst_path):
                    raise FileNotFoundError(f"PST file not found: {pst_path}")
                namespace.AddStore(pst_path)
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"Could not create/open PST file: {e}") from e
    
    def _remove_existing_store(self, namespace, pst_path):
        """Remove any existing store reference with the given path"""
        pst_path_lower = pst_path.lower()
        stores_to_remove = []
        
        # Find stores matching this path
        for store in namespace.Stores:
            try:
                if store.FilePath and store.FilePath.lower() == pst_path_lower:
                    stores_to_remove.append(store)
            except (AttributeError, OSError):
                continue
        
        # Remove found stores
        for store in stores_to_remove:
            try:
                root = store.GetRootFolder()
                namespace.RemoveStore(root)
                logger.info("Removed existing store reference: %s", pst_path)
                time.sleep(0.5)
            except Exception as e:
                logger.warning("Could not remove store reference: %s", e)
    
    def _create_new_pst(self, outlook, namespace, pst_path):
        """Create a new PST file using the most reliable method"""
        # Method 1: Try AddStoreEx (preferred - creates Unicode PST)
        try:
            namespace.AddStoreEx(pst_path, OL_STORE_UNICODE)
            logger.info("Created PST using AddStoreEx: %s", pst_path)
            return
        except AttributeError:
            logger.debug("AddStoreEx not available (older Outlook version)")
        except Exception as e1:
            logger.debug("AddStoreEx failed: %s", e1)
        
        # Method 2: Try AddStore
        try:
            namespace.AddStore(pst_path)
            logger.info("Created PST using AddStore: %s", pst_path)
            return
        except Exception as e2:
            logger.debug("AddStore failed: %s", e2)
        
        # Method 3: Initialize Outlook with a temp item, then try AddStore
        try:
            temp_mail = outlook.CreateItem(OL_MAIL_ITEM)
            temp_mail.Subject = "Temp"
            temp_mail.Save()
            entry_id = temp_mail.EntryID
            temp_mail.Delete()
            # Permanently remove from Deleted Items
            try:
                deleted = namespace.GetDefaultFolder(3)  # olFolderDeletedItems
                for item in deleted.Items:
                    if item.EntryID == entry_id:
                        item.Delete()
                        break
            except Exception as e:
                logger.debug("Could not purge temp item from Deleted Items: %s", e)
            
            namespace.AddStore(pst_path)
            logger.info("Created PST using AddStore after init: %s", pst_path)
            return
        except Exception as e3:
            logger.debug("Method 3 failed: %s", e3)
        
        raise RuntimeError(
            "Could not create PST file. Please try:\n"
            "1. Close Outlook completely and restart the converter\n"
            "2. Choose a different filename\n"
            "3. Run Outlook as Administrator"
        )
    
    def _find_pst_store(self, namespace, pst_path, retries=5):
        """Find the PST store in Outlook with retry logic"""
        pst_path_lower = pst_path.lower()
        
        for attempt in range(retries):
            for store in namespace.Stores:
                try:
                    if store.FilePath and store.FilePath.lower() == pst_path_lower:
                        return store
                except (AttributeError, OSError):
                    continue
            
            # Wait and retry if not found
            if attempt < retries - 1:
                time.sleep(1)
                logger.debug("PST store not found, retry %d/%d", attempt + 2, retries)
        
        return None
    
    def _get_or_create_inbox(self, root_folder):
        """Find or create Inbox folder in PST"""
        for folder in root_folder.Folders:
            if folder.Name.lower() == INBOX_FOLDER_NAME.lower():
                return folder
        return root_folder.Folders.Add(INBOX_FOLDER_NAME)
    
    def _process_email_files(self, outlook, namespace, target_folder, files_to_process, total, conversion_options):
        """Process all email files"""
        converted = 0
        skipped = 0
        errors = 0
        error_messages = []
        skipped_messages = []
        
        lang = conversion_options.get("lang", LANG_EN)
        for i, file_path in enumerate(files_to_process):
            current_file = os.path.basename(file_path)
            self.root.after(
                0,
                lambda f=current_file, idx=i, tot=total, lg=lang: self.status_label.config(
                    text=t(lg, "status_converting", name=f, cur=idx + 1, total=tot)
                ),
            )
            
            try:
                status, detail = self._process_single_email(
                    namespace, target_folder, file_path, outlook, conversion_options
                )
                if status == "converted":
                    converted += 1
                elif status == "skipped":
                    skipped += 1
                    if detail:
                        skipped_messages.append(f"{current_file}: {detail}")
                else:
                    errors += 1
                    error_messages.append(f"{current_file}: {detail or 'Unknown error'}")
            except Exception as e:
                errors += 1
                error_messages.append(f"{current_file}: {e}")
                logger.error("Error processing %s: %s", file_path, e)
                
            # Update progress - capture i properly
            self.root.after(0, lambda v=i+1: self.progress.config(value=v))
        
        # Verify items were added
        try:
            item_count = target_folder.Items.Count
            logger.info("Target folder now contains %d items", item_count)
        except (AttributeError, OSError) as e:
            logger.debug("Could not get item count: %s", e)
        
        return converted, skipped, errors, error_messages, skipped_messages
    
    def _process_single_email(self, namespace, target_folder, file_path, outlook, conversion_options):
        """Process a single email file"""
        try:
            sz = os.path.getsize(file_path)
            if sz > MAX_EML_FILE_BYTES:
                mb = MAX_EML_FILE_BYTES // (1024 * 1024)
                return (
                    "error",
                    f"File too large ({sz // (1024 * 1024)} MB); max {mb} MB (set EML2PST_MAX_FILE_MB)",
                )
        except OSError as e:
            return "error", f"Cannot read file: {e}"

        # Check for duplicates
        if conversion_options["remove_duplicates"]:
            file_hash = self.get_email_hash(file_path)
            if file_hash:
                with self._lock:
                    if file_hash in self.processed_hashes:
                        return "skipped", "Duplicate content"
                    self.processed_hashes.add(file_hash)
        
        # Method 1: Native Outlook import (best MIME preservation: full RFC822 in PST)
        native_result = self._import_with_outlook_native(
            namespace, target_folder, file_path, conversion_options
        )
        if native_result == "converted":
            return "converted", ""
        if conversion_options["strict_date_preservation"]:
            logger.warning(
                "Native import failed for %s; trying strict metadata fallback. Error: %s",
                file_path,
                native_result,
            )
            strict_result = self._process_single_email_manual_fallback(
                file_path,
                target_folder,
                outlook,
                require_date_preservation=True,
                conversion_options=conversion_options,
            )
            if strict_result == "converted":
                return "converted", ""
            return "skipped", (
                "Native import required to preserve original arrival date. "
                f"Reason: {native_result}; strict fallback failed: {strict_result}"
            )

        logger.warning(
            "Native import failed for %s; strict date preservation is OFF, using manual fallback. Error: %s",
            file_path,
            native_result,
        )
        fallback_result = self._process_single_email_manual_fallback(
            file_path,
            target_folder,
            outlook,
            require_date_preservation=False,
            conversion_options=conversion_options,
        )
        if fallback_result == "converted":
            return "converted", ""
        return "error", fallback_result

    def _process_single_email_manual_fallback(
        self, file_path, target_folder, outlook, require_date_preservation, conversion_options
    ):
        """
        Manual fallback import.
        Writes PR_TRANSPORT_MESSAGE_HEADERS from the parsed message (MIME headers).
        If strict (require_date_preservation): delivery time from Date/Received only;
        if that fails, the message is skipped.
        If not strict and "file mtime" is on: Windows modified time on the original
        .eml is applied first for delivery/submit time; otherwise Date/Received, then mtime.
        """
        try:
            email_data = self.parse_eml(file_path)
            if not email_data:
                return "Could not parse email"

            mail = outlook.CreateItem(OL_MAIL_ITEM)
            mail.Subject = str(email_data['subject'] or "(No Subject)")
            if email_data.get('to'):
                mail.To = str(email_data['to'])
            if email_data.get('cc'):
                mail.CC = str(email_data['cc'])

            body = email_data['body']
            if isinstance(body, str):
                if '<html' in body.lower() or '<body' in body.lower():
                    mail.HTMLBody = body
                else:
                    mail.Body = body
            else:
                mail.Body = str(body) if body else ""

            try:
                if email_data.get('from'):
                    mail.SentOnBehalfOfName = str(email_data['from'])
            except (AttributeError, TypeError) as e:
                logger.debug("Could not set sender: %s", e)

            use_mtime = conversion_options.get("use_file_mtime_for_date", True)
            prefer_mtime_first = not require_date_preservation and use_mtime
            date_ok = self._apply_original_metadata(
                mail,
                email_data,
                source_file_path=file_path,
                use_file_mtime=use_mtime,
                prefer_file_mtime_first=prefer_mtime_first,
            )
            if require_date_preservation and not date_ok:
                return "Could not preserve original arrival date metadata"

            self._add_attachments(mail, email_data.get('attachments', []))

            mail.Save()
            moved_mail = mail.Move(target_folder)
            # Move into PST often resets Received/list times to "now"; re-apply on the stored item.
            self._apply_original_metadata(
                moved_mail,
                email_data,
                source_file_path=file_path,
                use_file_mtime=use_mtime,
                prefer_file_mtime_first=prefer_mtime_first,
            )
            moved_mail.Save()
            return "converted"
        except Exception as e:
            logger.debug("Manual fallback failed for %s: %s", file_path, e)
            return str(e)

    def _open_shared_item_candidates(self, file_path):
        """
        Outlook OpenSharedItem often rejects plain paths and returns
        'Invalid path or URL'. Try short path and file:/// URI as well.
        """
        candidates = []
        try:
            abs_path = os.path.normpath(os.path.abspath(file_path))
        except OSError:
            abs_path = file_path

        if os.path.isfile(abs_path):
            # 1) Short path (helps with spaces / Unicode in path)
            if OUTLOOK_AVAILABLE:
                try:
                    import win32api
                    short = win32api.GetShortPathName(abs_path)
                    if short and os.path.isfile(short):
                        candidates.append(short)
                except Exception as e:
                    logger.debug("GetShortPathName skipped: %s", e)
            # 2) Absolute path string
            if abs_path not in candidates:
                candidates.append(abs_path)
            # 3) file:/// URI (Outlook COM often expects this)
            try:
                uri = Path(abs_path).resolve().as_uri()
                if uri not in candidates:
                    candidates.append(uri)
            except Exception as e:
                logger.debug("as_uri failed: %s", e)

        # De-dupe preserving order
        seen = set()
        out = []
        for c in candidates:
            if c and c not in seen:
                seen.add(c)
                out.append(c)
        return out

    def _open_shared_item(self, namespace, file_path):
        """Try OpenSharedItem with path variants; return mail item or raise last error."""
        last_err = None
        for candidate in self._open_shared_item_candidates(file_path):
            try:
                return namespace.OpenSharedItem(candidate)
            except Exception as e:
                last_err = e
                logger.debug("OpenSharedItem failed for %r: %s", candidate, e)
        if last_err:
            raise last_err
        raise OSError(f"Not a file or unreadable: {file_path}")

    def _outlook_com_datetime(self, dt_naive: datetime):
        """
        Outlook often ignores Python datetime for MAPI / MailItem time fields;
        pywintypes.Time is required for list columns (Received, etc.).
        """
        if dt_naive.tzinfo is not None:
            dt_naive = dt_naive.astimezone().replace(tzinfo=None)
        if PYWINTYPES:
            try:
                return PYWINTYPES.Time(dt_naive)
            except Exception:
                pass
        return dt_naive

    def _stamp_outlook_item_with_explorer_mtime(self, mail_item, source_eml_path: str) -> bool:
        """
        Set delivery/submit/creation-style MAPI times and MailItem.ReceivedTime /
        SentOn from the source file's Windows last-write time (Explorer Date modified).
        """
        if not source_eml_path or not os.path.isfile(source_eml_path):
            return False
        try:
            ts = os.path.getmtime(source_eml_path)
            dt_local = datetime.fromtimestamp(ts)
        except OSError as e:
            logger.warning("Could not read mtime for %s: %s", source_eml_path, e)
            return False

        com_val = self._outlook_com_datetime(dt_local)
        ok = False
        # MAPI PT_SYSTIME tags the message list / sorting often use
        mapi_time_urls = (
            "http://schemas.microsoft.com/mapi/proptag/0x0E060040",  # PR_MESSAGE_DELIVERY_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x00390040",  # PR_CLIENT_SUBMIT_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x30070040",  # PR_CREATION_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x30080040",  # PR_LAST_MODIFICATION_TIME
        )
        try:
            pa = mail_item.PropertyAccessor
            for url in mapi_time_urls:
                try:
                    pa.SetProperty(url, com_val)
                    ok = True
                except Exception as e:
                    logger.debug("SetProperty %s: %s", url, e)
        except Exception as e:
            logger.warning("PropertyAccessor unavailable for mtime stamp: %s", e)

        try:
            mail_item.ReceivedTime = com_val
            ok = True
        except Exception as e:
            logger.debug("MailItem.ReceivedTime: %s", e)
        try:
            mail_item.SentOn = com_val
            ok = True
        except Exception as e:
            logger.debug("MailItem.SentOn: %s", e)

        if not ok:
            logger.warning(
                "Explorer mtime stamp had no effect for %s (Outlook may block writes)",
                source_eml_path,
            )
        return ok

    def _apply_original_metadata(
        self,
        mail,
        email_data,
        source_file_path=None,
        use_file_mtime=True,
        prefer_file_mtime_first=False,
    ):
        """
        Preserve MIME headers and message time in Outlook (fallback path).

        - PR_TRANSPORT_MESSAGE_HEADERS: full header block from the .eml.
        - PR_MESSAGE_DELIVERY_TIME / PR_CLIENT_SUBMIT_TIME:
          If prefer_file_mtime_first and use_file_mtime: Windows ``mtime`` on the
          original file first (Explorer "Date modified"); else Date/Received headers,
          then mtime as last resort.

        Returns True if a message date was applied (header-based or mtime).
        """
        date_applied = False
        try:
            pa = mail.PropertyAccessor
        except Exception as e:
            logger.debug("PropertyAccessor unavailable: %s", e)
            return False

        try:
            headers = self._build_transport_headers(email_data)
            if headers:
                try:
                    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E", headers)
                except Exception:
                    # Unicode transport headers (some Outlook builds)
                    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F", headers)
        except Exception as e:
            logger.debug("Could not set transport headers: %s", e)

        def _apply_mtime_from_source():
            nonlocal date_applied
            if not use_file_mtime or not source_file_path:
                return
            try:
                if os.path.isfile(source_file_path):
                    ts = os.path.getmtime(source_file_path)
                    dt_local = datetime.fromtimestamp(ts)
                    com_val = self._outlook_com_datetime(dt_local)
                    for url in (
                        "http://schemas.microsoft.com/mapi/proptag/0x0E060040",
                        "http://schemas.microsoft.com/mapi/proptag/0x00390040",
                        "http://schemas.microsoft.com/mapi/proptag/0x30070040",
                        "http://schemas.microsoft.com/mapi/proptag/0x30080040",
                    ):
                        try:
                            pa.SetProperty(url, com_val)
                        except Exception:
                            pass
                    try:
                        mail.ReceivedTime = com_val
                    except Exception:
                        pass
                    try:
                        mail.SentOn = com_val
                    except Exception:
                        pass
                    date_applied = True
                    logger.info(
                        "Applied delivery time from file modification time: %s",
                        source_file_path,
                    )
            except Exception as e:
                logger.debug("Could not apply file mtime as date: %s", e)

        if prefer_file_mtime_first:
            _apply_mtime_from_source()

        if not date_applied:
            try:
                date_raw = self._first_parseable_date_header(email_data)
                if date_raw:
                    dt = parsedate_to_datetime(date_raw)
                    if dt is not None:
                        if dt.tzinfo is None:
                            dt_local = dt
                        else:
                            dt_local = dt.astimezone().replace(tzinfo=None)
                        com_val = self._outlook_com_datetime(dt_local)
                        for url in (
                            "http://schemas.microsoft.com/mapi/proptag/0x0E060040",
                            "http://schemas.microsoft.com/mapi/proptag/0x00390040",
                        ):
                            try:
                                pa.SetProperty(url, com_val)
                            except Exception:
                                pass
                        try:
                            mail.ReceivedTime = com_val
                        except Exception:
                            pass
                        try:
                            mail.SentOn = com_val
                        except Exception:
                            pass
                        date_applied = True
            except Exception as e:
                logger.debug("Could not set original date properties: %s", e)

        if not date_applied:
            _apply_mtime_from_source()

        return date_applied

    def _first_parseable_date_header(self, email_data):
        """
        First usable instant for Outlook delivery/submit time.

        Order: ``Date:``, then **each** ``Received:`` line (common in real mail;
        msg.get('Received') only returns one), then Resent-Date / Delivery-Date.
        """
        msg = email_data.get("message")

        def _parse_stamp(raw: str):
            if not raw:
                return None
            raw = raw.strip()
            try:
                if parsedate_to_datetime(raw) is not None:
                    return raw
            except (TypeError, ValueError):
                pass
            return None

        # Date:
        raw = None
        if msg is not None:
            raw = msg.get("Date")
        if not raw:
            raw = email_data.get("date")
        found = _parse_stamp(raw) if raw else None
        if found:
            return found

        # Every Received: (hop chain); date is usually after the last ';'
        if msg is not None:
            received_vals = msg.get_all("Received") or []
            if received_vals:
                for rec in received_vals:
                    candidate = rec
                    if ";" in rec:
                        candidate = rec.rsplit(";", 1)[-1].strip()
                    found = _parse_stamp(candidate)
                    if found:
                        return found

        for key in ("Resent-Date", "Delivery-Date"):
            raw = msg.get(key) if msg is not None else None
            if raw:
                found = _parse_stamp(raw)
                if found:
                    return found

        return email_data.get("date")

    def _build_transport_headers(self, email_data):
        """
        Build RFC822 header block for PR_TRANSPORT_MESSAGE_HEADERS (MIME headers).

        Prefer ``raw_items()`` so order and **repeated** headers (e.g. multiple
        ``Received:``) match the source. Otherwise join ``get_all()`` per header
        name so Received chains are not dropped.
        """
        msg = email_data.get("message")
        if msg is None:
            return ""
        try:
            if hasattr(msg, "raw_items"):
                lines = [f"{k}: {v}" for (k, v) in msg.raw_items()]
            else:
                lines = []
                for name in msg.keys():
                    vals = msg.get_all(name)
                    if not vals:
                        continue
                    for value in vals:
                        lines.append(f"{name}: {value}")
            if lines:
                return "\r\n".join(lines) + "\r\n"
        except Exception as e:
            logger.debug("Could not build transport headers: %s", e)
        return ""

    def _import_with_outlook_native(self, namespace, target_folder, file_path, conversion_options):
        """
        Import via Outlook OpenSharedItem (best path for intact MIME: Date, Received
        chain, Content-Type, etc. stay with the message as Outlook stored them).

        If use_file_mtime_for_date is True, overwrites list-view times with the source
        file's Explorer 'Date modified' so PST dates match the folder you picked files from.
        If False, leaves Outlook's dates from the .eml content only.

        For .emlx or malformed files, normalizes line endings to a staging .eml
        without rewriting headers so MIME remains preserved.
        """
        use_explorer_mtime = conversion_options.get("use_file_mtime_for_date", True)

        # 1) Direct native import first (.eml and some .emlx files)
        try:
            mail_item = self._open_shared_item(namespace, file_path)
            moved_item = mail_item.Move(target_folder)
            moved_item.Save()
            # First Save often locks "Received" to import time; stamp then Save again.
            if use_explorer_mtime:
                self._stamp_outlook_item_with_explorer_mtime(moved_item, file_path)
                moved_item.Save()
            del mail_item
            del moved_item
            gc.collect()
            return "converted"
        except Exception as e:
            logger.debug("OpenSharedItem direct failed for %s: %s", file_path, e)

        # 2) Convert source to normalized RFC822 .eml and retry native import.
        # This helps Outlook parse some malformed .eml exports while preserving headers/date.
        ext = os.path.splitext(file_path)[1].lower()
        try:
            if ext == ".emlx":
                rfc822_bytes = self._emlx_to_rfc822_bytes(file_path)
            else:
                with open(file_path, "rb") as f:
                    rfc822_bytes = f.read()
            if not rfc822_bytes:
                return "Could not decode source message content"

            rfc822_bytes = self._normalize_rfc822_line_endings(rfc822_bytes)

            last_err = None
            for staging_dir in self._native_temp_staging_dirs(file_path):
                temp_eml_path = None
                try:
                    fd, temp_eml_path = tempfile.mkstemp(
                        suffix=".eml",
                        prefix=f"emlx_as_eml_{uuid.uuid4().hex[:8]}_",
                        dir=staging_dir,
                    )
                    with self._lock:
                        self._temp_files.append(temp_eml_path)
                    try:
                        os.write(fd, rfc822_bytes)
                        os.fsync(fd)
                    finally:
                        os.close(fd)

                    if not os.path.isfile(temp_eml_path):
                        raise OSError(f"staging file missing after write: {temp_eml_path}")
                    time.sleep(0.08)

                    mail_item = self._open_shared_item(namespace, temp_eml_path)
                    moved_item = mail_item.Move(target_folder)
                    moved_item.Save()
                    if use_explorer_mtime:
                        self._stamp_outlook_item_with_explorer_mtime(moved_item, file_path)
                        moved_item.Save()
                    del mail_item
                    del moved_item
                    gc.collect()
                    with self._lock:
                        self._native_staging_paths.append(temp_eml_path)
                    return "converted"
                except Exception as e:
                    last_err = e
                    logger.debug(
                        "RFC822 native import failed (staging_dir=%s): %s",
                        staging_dir,
                        e,
                    )
                    if temp_eml_path and os.path.exists(temp_eml_path):
                        try:
                            os.remove(temp_eml_path)
                        except OSError:
                            pass
                        with self._lock:
                            if temp_eml_path in self._temp_files:
                                self._temp_files.remove(temp_eml_path)

            return f"Native import failed after RFC822 normalization: {last_err}"
        except Exception as e:
            return f"Native import failed after RFC822 normalization: {e}"

    def _normalize_rfc822_line_endings(self, data):
        """Normalize RFC822 data to CRLF endings for better Outlook compatibility."""
        if not data:
            return data
        normalized = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
        return b"\r\n".join(normalized.split(b"\n"))

    def _emlx_to_rfc822_bytes(self, file_path):
        """
        Convert Apple .emlx to RFC822 .eml bytes.
        .emlx is usually: <byte_count_line>\\n<rfc822_message><plist...>
        """
        try:
            with open(file_path, "rb") as f:
                data = f.read()
        except OSError as e:
            logger.debug("Could not read EMLX %s: %s", file_path, e)
            return None

        newline_idx = data.find(b"\n")
        if newline_idx <= 0:
            return data

        first_line = data[:newline_idx].strip()
        if not first_line.isdigit():
            return data

        try:
            expected_len = int(first_line)
        except ValueError:
            return data

        start = newline_idx + 1
        end = start + expected_len
        if end <= len(data):
            return data[start:end]
        return data[start:]
    
    def _add_attachments(self, mail, attachments):
        """Add attachments to mail item"""
        for att in attachments:
            if not att.get('data'):
                continue
            
            temp_path = None
            try:
                filename = att.get('filename', 'attachment')
                safe_filename = self._sanitize_filename(filename)
                _, ext = os.path.splitext(safe_filename)
                
                fd, temp_path = tempfile.mkstemp(suffix=ext, prefix=f"eml_att_{uuid.uuid4().hex[:8]}_")
                with self._lock:
                    self._temp_files.append(temp_path)
                try:
                    os.write(fd, att['data'])
                finally:
                    os.close(fd)
                
                mail.Attachments.Add(temp_path)
            except (OSError, IOError) as e:
                logger.warning("Could not add attachment %s: %s", att.get('filename'), e)
            finally:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                        with self._lock:
                            if temp_path in self._temp_files:
                                self._temp_files.remove(temp_path)
                    except OSError as e:
                        logger.debug("Could not remove temp file %s: %s", temp_path, e)
        
    def show_completion(self, converted, skipped, errors, note="", lang=None):
        """Show completion message"""
        if lang is None:
            lang = self._current_lang()
        self.status_label.config(text=t(lang, "status_complete"))
        self.progress.config(value=self.progress['maximum'])
        
        msg = t(
            lang,
            "msg_complete",
            converted=converted,
            skipped=skipped,
            errors=errors,
        )
        msg += note
        
        messagebox.showinfo(t(lang, "title_complete"), msg)


def main():
    root = tk.Tk()
    
    # Set style
    style = ttk.Style()
    style.theme_use('clam')
    
    EmlToPstConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()

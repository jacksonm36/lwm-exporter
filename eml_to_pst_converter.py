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
import logging
import tempfile
import uuid
import atexit

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
try:
    import win32com.client
    import pythoncom
    WIN32COM = win32com.client
    PYTHONCOM = pythoncom
    OUTLOOK_AVAILABLE = True
except ImportError:
    logger.warning("pywin32 not available - PST conversion will require installation")

# Maximum files to process (memory safety)
MAX_FILES = 50000


class EmlToPstConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("EML to PST Converter")
        self.root.geometry("750x550")
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.folder_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.pst_option = tk.StringVar(value="new")
        self.remove_duplicates = tk.BooleanVar(value=False)
        self.file_pattern = tk.StringVar(value="*.eml")
        self.eml_files = []
        self.processed_hashes = set()
        
        # Thread safety
        self._lock = threading.Lock()
        self._is_converting = False
        self._temp_files = []
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        atexit.register(self._cleanup_temp_files)
        
        self.create_widgets()
        
    def create_widgets(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # === Add Folder Section ===
        folder_frame = ttk.LabelFrame(main_frame, text="Add Folder Having *.eml / *.emlx Files", padding="10")
        folder_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Folder path entry and browse
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.folder_entry = ttk.Entry(path_frame, textvariable=self.folder_path, width=70)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        add_files_btn = ttk.Button(path_frame, text="Add Files", command=self.browse_folder)
        add_files_btn.pack(side=tk.RIGHT)
        
        # Wildcard pattern frame
        pattern_frame = ttk.Frame(folder_frame)
        pattern_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(pattern_frame, text="File Pattern:").pack(side=tk.LEFT)
        
        pattern_combo = ttk.Combobox(pattern_frame, textvariable=self.file_pattern, width=15)
        pattern_combo['values'] = ('*.eml', '*.emlx', '*.eml;*.emlx')
        pattern_combo.pack(side=tk.LEFT, padx=(5, 10))
        
        ttk.Label(pattern_frame, text="(Wildcards: * for any characters, ? for single character)").pack(side=tk.LEFT)
        
        # === Save in PST Section ===
        pst_frame = ttk.LabelFrame(main_frame, text="Save in PST", padding="10")
        pst_frame.pack(fill=tk.X, pady=(0, 10))
        
        options_frame = ttk.Frame(pst_frame)
        options_frame.pack(fill=tk.X)
        
        ttk.Radiobutton(options_frame, text="Create New PST File", 
                       variable=self.pst_option, value="new").pack(side=tk.LEFT, padx=(0, 20))
        ttk.Radiobutton(options_frame, text="Saving in Existing PST File", 
                       variable=self.pst_option, value="existing").pack(side=tk.LEFT, padx=(0, 20))
        ttk.Checkbutton(options_frame, text="Remove Duplicate Content", 
                       variable=self.remove_duplicates).pack(side=tk.LEFT)
        
        # === File List Section ===
        list_frame = ttk.LabelFrame(main_frame, text="EML/EMLX Files", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Create Treeview with scrollbars
        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbars
        y_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        x_scroll.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview
        self.file_tree = ttk.Treeview(tree_container, columns=("name", "path", "size", "date"),
                                       show="headings", height=10,
                                       yscrollcommand=y_scroll.set,
                                       xscrollcommand=x_scroll.set)
        
        self.file_tree.heading("name", text="EML/EMLX Name")
        self.file_tree.heading("path", text="Path")
        self.file_tree.heading("size", text="Size")
        self.file_tree.heading("date", text="Date Modified")
        
        self.file_tree.column("name", width=200, minwidth=150)
        self.file_tree.column("path", width=300, minwidth=200)
        self.file_tree.column("size", width=80, minwidth=60)
        self.file_tree.column("date", width=120, minwidth=100)
        
        self.file_tree.pack(fill=tk.BOTH, expand=True)
        
        y_scroll.config(command=self.file_tree.yview)
        x_scroll.config(command=self.file_tree.xview)
        
        # File count label
        self.file_count_label = ttk.Label(list_frame, text="Total Files: 0")
        self.file_count_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Context menu for treeview
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Remove Selected", command=self.remove_selected)
        self.context_menu.add_command(label="Clear All", command=self.clear_all)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # === Destination Section ===
        dest_frame = ttk.Frame(main_frame)
        dest_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(dest_frame, text="Destination:").pack(side=tk.LEFT)
        
        self.dest_entry = ttk.Entry(dest_frame, textvariable=self.destination_path, width=60)
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        browse_dest_btn = ttk.Button(dest_frame, text="Browse Destination", command=self.browse_destination)
        browse_dest_btn.pack(side=tk.RIGHT)
        
        # === Bottom Buttons ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(button_frame, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(button_frame, text="Ready")
        self.status_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # Convert and Exit buttons
        exit_btn = ttk.Button(button_frame, text="Exit", command=self._on_close, width=10)
        exit_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.convert_btn = ttk.Button(button_frame, text="Convert", command=self.start_conversion, width=10)
        self.convert_btn.pack(side=tk.RIGHT)
        
    def browse_folder(self):
        """Browse for folder containing EML/EMLX files"""
        folder = filedialog.askdirectory(title="Select Folder Containing EML/EMLX Files")
        if folder:
            self.folder_path.set(folder)
            # Run folder scan in background thread to prevent UI freeze
            thread = threading.Thread(target=self._scan_folder_thread, args=(folder,))
            thread.daemon = True
            thread.start()
    
    def _scan_folder_thread(self, folder):
        """Background thread for folder scanning"""
        self.root.after(0, lambda: self.status_label.config(text="Scanning folder..."))
        
        try:
            files = self._scan_folder_impl(folder)
            self.root.after(0, lambda: self._add_files_to_list(files))
        except Exception as e:
            logger.error(f"Error scanning folder: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Error scanning folder: {e}"))
            
    def _scan_folder_impl(self, folder):
        """Scan folder for EML/EMLX files using wildcard pattern"""
        pattern_input = self.file_pattern.get()
        
        # Validate pattern - only allow safe patterns
        if not self._is_valid_pattern(pattern_input):
            raise ValueError("Invalid file pattern. Only *.eml, *.emlx patterns are allowed.")
        
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
                logger.warning(f"File limit reached ({MAX_FILES})")
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
    
    def _add_files_to_list(self, files):
        """Add files to the treeview (runs on main thread)"""
        for file_path in files:
            self.add_file_to_list(file_path)
            
        self.update_file_count()
        self.status_label.config(text=f"Found {len(files)} files")
            
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
            logger.warning(f"Could not get file info for {file_path}: {e}")
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
        self.file_count_label.config(text=f"Total Files: {count}")
        
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
        if self.pst_option.get() == "new":
            file_path = filedialog.asksaveasfilename(
                title="Save PST File",
                defaultextension=".pst",
                filetypes=[("Outlook PST Files", "*.pst"), ("All Files", "*.*")]
            )
        else:
            file_path = filedialog.askopenfilename(
                title="Select Existing PST File",
                filetypes=[("Outlook PST Files", "*.pst"), ("All Files", "*.*")]
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
            logger.warning(f"Could not hash file {file_path}: {e}")
            return None
            
    def parse_eml(self, file_path):
        """Parse an EML file and return email data"""
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
            logger.error(f"Error parsing EML file {file_path}: {e}")
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
                    except (KeyError, LookupError) as e:
                        logger.debug(f"Could not decode text/plain part: {e}")
                elif content_type == "text/html" and not body:
                    try:
                        body = part.get_content()
                    except (KeyError, LookupError) as e:
                        logger.debug(f"Could not decode text/html part: {e}")
        else:
            try:
                body = msg.get_content()
            except (KeyError, LookupError):
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
        with self._lock:
            if self._is_converting:
                if not messagebox.askyesno("Confirm", "Conversion in progress. Exit anyway?"):
                    return
        self._cleanup_temp_files()
        self.root.destroy()
    
    def _cleanup_temp_files(self):
        """Remove any leftover temp files created during attachment handling"""
        for path in self._temp_files:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError:
                pass
        self._temp_files.clear()
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        with self._lock:
            if self._is_converting:
                messagebox.showwarning("Warning", "Conversion already in progress!")
                return
            if not self.eml_files:
                messagebox.showwarning("Warning", "No EML/EMLX files to convert!")
                return
            self._is_converting = True
            
        if not self.destination_path.get():
            with self._lock:
                self._is_converting = False
            messagebox.showwarning("Warning", "Please select a destination path!")
            return
        
        dest_path = self.destination_path.get()
        dest_dir = os.path.dirname(os.path.abspath(dest_path))
        if not os.path.isdir(dest_dir):
            with self._lock:
                self._is_converting = False
            messagebox.showerror("Error", f"Destination directory does not exist: {dest_dir}")
            return
            
        self.convert_btn.config(state='disabled')
        
        thread = threading.Thread(target=self._convert_files_thread)
        thread.daemon = True
        thread.start()
    
    def _convert_files_thread(self):
        """Thread wrapper for conversion with COM initialization"""
        com_initialized = False
        try:
            if PYTHONCOM:
                PYTHONCOM.CoInitialize()
                com_initialized = True
            self.convert_files()
        finally:
            if com_initialized:
                PYTHONCOM.CoUninitialize()
            with self._lock:
                self._is_converting = False
            self.root.after(0, lambda: self.convert_btn.config(state='normal'))
        
    def convert_files(self):
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
            self.convert_with_outlook(files_to_process)
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Conversion error: {error_msg}")
            self.root.after(0, lambda: messagebox.showerror(
                "Conversion Error",
                f"Error during conversion:\n{error_msg}\n\n"
                "Make sure Microsoft Outlook is installed and working properly."
            ))
    
    def prompt_install_pywin32(self):
        """Prompt user to install pywin32"""
        result = messagebox.askyesno(
            "Missing Dependency",
            "The 'pywin32' library is required for PST conversion.\n\n"
            "Would you like to install it now?\n\n"
            "(This requires an internet connection)"
        )
        
        if result:
            self.status_label.config(text="Installing pywin32...")
            self.root.update()
            
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "pywin32"],
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )
                messagebox.showinfo(
                    "Installation Complete",
                    "pywin32 has been installed successfully!\n\n"
                    "Please restart the application to use PST conversion."
                )
            except subprocess.CalledProcessError as e:
                logger.error(f"Failed to install pywin32: {e}")
                messagebox.showerror(
                    "Installation Failed",
                    f"Failed to install pywin32.\n\n"
                    "Please run manually in command prompt:\npip install pywin32"
                )
            
            self.status_label.config(text="Ready")
            
    def convert_with_outlook(self, files_to_process):
        """Convert using Outlook COM interface (requires Microsoft Outlook)"""
        self._update_status("Connecting to Outlook...")
        
        outlook = None
        namespace = None
        try:
            try:
                outlook = WIN32COM.Dispatch("Outlook.Application")
            except Exception as e:
                raise Exception(f"Could not connect to Outlook: {e}")
            
            try:
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                raise Exception(f"Could not access MAPI namespace: {e}")
            
            self._cleanup_stale_stores(namespace)
            
            pst_path = os.path.abspath(self.destination_path.get())
            
            self._update_status("Creating PST file...")
            self._setup_pst_store(outlook, namespace, pst_path)
            
            time.sleep(1)
                
            pst_store = self._find_pst_store(namespace, pst_path)
            if not pst_store:
                raise Exception(f"Could not access PST file after creation. Path: {pst_path}")
                
            root_folder = pst_store.GetRootFolder()
            target_folder = self._get_or_create_inbox(root_folder)
                
            total = len(files_to_process)
            converted, skipped, errors, error_messages = self._process_email_files(
                outlook, namespace, target_folder, files_to_process, total
            )
            
            note = f"\n\nPST saved to: {pst_path}"
            if error_messages:
                note += "\n\nErrors:\n" + "\n".join(error_messages[:5])
                if len(error_messages) > 5:
                    note += f"\n... and {len(error_messages) - 5} more errors"
                
            self.root.after(0, lambda: self.show_completion(converted, skipped, errors, note))
        finally:
            # Release COM objects to prevent Outlook from hanging
            del namespace
            del outlook
            import gc
            gc.collect()
    
    def _update_status(self, text):
        """Thread-safe status update"""
        self.root.after(0, lambda: self.status_label.config(text=text))
    
    def _cleanup_stale_stores(self, namespace):
        """Remove store references to PST files that no longer exist"""
        try:
            stale_stores = []
            for store in namespace.Stores:
                try:
                    file_path = store.FilePath
                    # Check if this is a PST file that no longer exists
                    if file_path and file_path.lower().endswith('.pst'):
                        if not os.path.exists(file_path):
                            stale_stores.append((store, file_path))
                except (AttributeError, OSError):
                    continue
            
            # Remove stale stores
            for store, file_path in stale_stores:
                try:
                    root = store.GetRootFolder()
                    namespace.RemoveStore(root)
                    logger.info(f"Removed stale store reference: {file_path}")
                except Exception as e:
                    logger.warning(f"Could not remove stale store {file_path}: {e}")
                    
        except Exception as e:
            logger.debug(f"Stale store cleanup failed: {e}")
    
    def _setup_pst_store(self, outlook, namespace, pst_path):
        """Create or open PST store"""
        try:
            # First, remove any existing store reference with this path
            self._remove_existing_store(namespace, pst_path)
            
            if self.pst_option.get() == "new":
                # Remove existing file if present
                if os.path.exists(pst_path):
                    try:
                        os.remove(pst_path)
                    except OSError as e:
                        logger.warning(f"Could not remove existing PST: {e}")
                
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
            raise Exception(f"Could not create/open PST file: {e}")
    
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
                logger.info(f"Removed existing store reference: {pst_path}")
                time.sleep(0.5)
            except Exception as e:
                logger.warning(f"Could not remove store reference: {e}")
    
    def _create_new_pst(self, outlook, namespace, pst_path):
        """Create a new PST file using the most reliable method"""
        # Method 1: Try AddStoreEx (preferred - creates Unicode PST)
        try:
            namespace.AddStoreEx(pst_path, OL_STORE_UNICODE)
            logger.info(f"Created PST using AddStoreEx: {pst_path}")
            return
        except AttributeError:
            logger.debug("AddStoreEx not available (older Outlook version)")
        except Exception as e1:
            logger.debug(f"AddStoreEx failed: {e1}")
        
        # Method 2: Try AddStore
        try:
            namespace.AddStore(pst_path)
            logger.info(f"Created PST using AddStore: {pst_path}")
            return
        except Exception as e2:
            logger.debug(f"AddStore failed: {e2}")
        
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
            except Exception:
                pass
            
            namespace.AddStore(pst_path)
            logger.info(f"Created PST using AddStore after init: {pst_path}")
            return
        except Exception as e3:
            logger.debug(f"Method 3 failed: {e3}")
        
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
                logger.debug(f"PST store not found, retry {attempt + 2}/{retries}")
        
        return None
    
    def _get_or_create_inbox(self, root_folder):
        """Find or create Inbox folder in PST"""
        for folder in root_folder.Folders:
            if folder.Name.lower() == INBOX_FOLDER_NAME.lower():
                return folder
        return root_folder.Folders.Add(INBOX_FOLDER_NAME)
    
    def _process_email_files(self, outlook, namespace, target_folder, files_to_process, total):
        """Process all email files"""
        converted = 0
        skipped = 0
        errors = 0
        error_messages = []
        
        for i, file_path in enumerate(files_to_process):
            current_file = os.path.basename(file_path)
            # Capture i in closure properly
            self.root.after(0, lambda f=current_file, idx=i: self.status_label.config(
                text=f"Converting: {f} ({idx+1}/{total})"))
            
            try:
                result = self._process_single_email(namespace, target_folder, file_path, outlook)
                if result == "converted":
                    converted += 1
                elif result == "skipped":
                    skipped += 1
                else:
                    errors += 1
                    error_messages.append(f"{current_file}: {result}")
            except Exception as e:
                errors += 1
                error_messages.append(f"{current_file}: {e}")
                logger.error(f"Error processing {file_path}: {e}")
                
            # Update progress - capture i properly
            self.root.after(0, lambda v=i+1: self.progress.config(value=v))
        
        # Verify items were added
        try:
            item_count = target_folder.Items.Count
            logger.info(f"Target folder now contains {item_count} items")
        except (AttributeError, OSError) as e:
            logger.debug(f"Could not get item count: {e}")
        
        return converted, skipped, errors, error_messages
    
    def _process_single_email(self, namespace, target_folder, file_path, outlook):
        """Process a single email file"""
        # Check for duplicates
        if self.remove_duplicates.get():
            file_hash = self.get_email_hash(file_path)
            if file_hash and file_hash in self.processed_hashes:
                return "skipped"
            if file_hash:
                self.processed_hashes.add(file_hash)
        
        # Method 1: Try OpenSharedItem and move directly to target folder
        try:
            mail_item = namespace.OpenSharedItem(file_path)
            # Move directly to target folder (returns the moved item)
            moved_item = mail_item.Move(target_folder)
            # Save to ensure it's persisted
            moved_item.Save()
            return "converted"
        except Exception as e1:
            logger.debug(f"OpenSharedItem failed for {file_path}: {e1}")
        
        # Method 2: Create mail item using Outlook.CreateItem and copy to target
        try:
            email_data = self.parse_eml(file_path)
            if not email_data:
                return "Could not parse email"
            
            # Create mail item in default location
            mail = outlook.CreateItem(OL_MAIL_ITEM)
            
            mail.Subject = str(email_data['subject'] or "(No Subject)")
            
            body = email_data['body']
            if isinstance(body, str):
                if '<html' in body.lower() or '<body' in body.lower():
                    mail.HTMLBody = body
                else:
                    mail.Body = body
            else:
                mail.Body = str(body) if body else ""
            
            # Try to set sender info (may fail if property is read-only)
            try:
                if email_data.get('from'):
                    mail.SentOnBehalfOfName = str(email_data['from'])
            except (AttributeError, TypeError) as e:
                logger.debug(f"Could not set sender: {e}")
            
            # Add attachments
            self._add_attachments(mail, email_data.get('attachments', []))
            
            # Save first, then move to target folder
            mail.Save()
            moved_mail = mail.Move(target_folder)
            moved_mail.Save()
            
            return "converted"
        except Exception as e2:
            logger.debug(f"Method 2 failed for {file_path}: {e2}")
            return str(e2)
    
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
                self._temp_files.append(temp_path)
                try:
                    os.write(fd, att['data'])
                finally:
                    os.close(fd)
                
                mail.Attachments.Add(temp_path)
            except (OSError, IOError) as e:
                logger.warning(f"Could not add attachment {att.get('filename')}: {e}")
            finally:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                        if temp_path in self._temp_files:
                            self._temp_files.remove(temp_path)
                    except OSError:
                        pass
        
    def show_completion(self, converted, skipped, errors, note=""):
        """Show completion message"""
        self.status_label.config(text="Conversion Complete")
        self.progress.config(value=self.progress['maximum'])
        
        msg = f"Conversion Complete!\n\n"
        msg += f"Converted: {converted} files\n"
        msg += f"Skipped (duplicates): {skipped} files\n"
        msg += f"Errors: {errors} files"
        msg += note
        
        messagebox.showinfo("Complete", msg)


def main():
    root = tk.Tk()
    
    # Set style
    style = ttk.Style()
    style.theme_use('clam')
    
    EmlToPstConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()

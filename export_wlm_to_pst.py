import os
import sys
import traceback
import threading
import logging
import shutil
import tempfile
import csv
from pathlib import Path
from tkinter import Tk, Label, Button, Entry, StringVar, filedialog, messagebox
from tkinter import ttk

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    win32com = None
    pythoncom = None


APP_TITLE = "Windows Live Mail → Outlook PST Exporter"


def select_wlm_folder(path_var: StringVar):
    folder = filedialog.askdirectory(title="Select Windows Live Mail root folder")
    if folder:
        path_var.set(folder)


def select_pst_file(path_var: StringVar):
    file = filedialog.asksaveasfilename(
        title="Select / create target PST file",
        defaultextension=".pst",
        filetypes=[("Outlook Data File (*.pst)", "*.pst")],
    )
    if file:
        path_var.set(file)


def list_eml_files(root: Path):
    for dirpath, _dirnames, filenames in os.walk(root):
        for name in filenames:
            if name.lower().endswith(".eml"):
                yield Path(dirpath) / name


def ensure_pst_folder_for_eml(import_root, root: Path, eml_path: Path, folder_cache):
    """
    Determine / create the PST folder that corresponds to the .eml's relative path.
    This preserves the account / folder layout from the Windows Live Mail export.
    """
    try:
        rel = eml_path.relative_to(root)
    except ValueError:
        rel = eml_path.name

    # All parent parts (account + folders), without the .eml filename
    if isinstance(rel, Path):
        parts = list(rel.parts[:-1])
    else:
        parts = []

    current_key = ""
    current_folder = import_root
    for part in parts:
        current_key = current_key + "/" + part if current_key else part
        cached = folder_cache.get(current_key)
        if cached is not None:
            current_folder = cached
            continue
        try:
            next_folder = current_folder.Folders(part)
        except Exception:
            next_folder = current_folder.Folders.Add(part)
        folder_cache[current_key] = next_folder
        current_folder = next_folder

    return current_folder


def run_export_async(
    root: Path,
    pst_path: Path,
    status_var: StringVar,
    progress_bar: "ttk.Progressbar",
    root_window: Tk,
    control_flags: dict,
):
    thread = threading.Thread(
        target=run_export,
        args=(root, pst_path, status_var, progress_bar, root_window, control_flags),
        daemon=True,
    )
    thread.start()


def run_export(
    root: Path,
    pst_path: Path,
    status_var: StringVar,
    progress_bar: "ttk.Progressbar",
    root_window: Tk,
    control_flags: dict,
):
    if win32com is None or pythoncom is None:
        messagebox.showerror(
            APP_TITLE,
            "pywin32 is not installed.\n\nInstall dependencies with:\n\n"
            "  pip install -r requirements.txt",
        )
        return

    try:
        logging.info("Starting export: root=%s, pst=%s", root, pst_path)
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        session = outlook.Session

        # Attach / create the PST store
        pst_full = str(pst_path)
        if not pst_path.exists():
            # Creating a new PST
            try:
                # 1 == olStoreUnicode for modern Unicode PST (if supported)
                logging.info("Creating new PST via AddStoreEx: %s", pst_full)
                session.AddStoreEx(pst_full, 1)
            except Exception:
                # Fallback to default AddStore for older Outlook versions
                logging.warning("AddStoreEx failed, falling back to AddStore for %s", pst_full)
                session.AddStore(pst_full)

        # Find the store we just added / opened
        target_store = None
        for store in session.Stores:
            if os.path.abspath(store.FilePath).lower() == os.path.abspath(pst_full).lower():
                target_store = store
                break

        if target_store is None:
            logging.error("Failed to find attached PST store for %s", pst_full)
            messagebox.showerror(
                APP_TITLE,
                "Could not attach to the PST in Outlook.\n"
                "Verify that Outlook is installed and supports PST files.",
            )
            return

        root_folder = target_store.GetRootFolder()

        # Create / reuse a root folder inside the PST for imported mail
        import_folder_name = "Imported from Windows Live Mail"
        try:
            import_root = root_folder.Folders(import_folder_name)
        except Exception:
            import_root = root_folder.Folders.Add(import_folder_name)

        eml_files = list(list_eml_files(root))
        total = len(eml_files)
        if total == 0:
            logging.info("No .eml files found under %s", root)
            messagebox.showinfo(
                APP_TITLE,
                "No .eml files were found under the selected Windows Live Mail folder.",
            )
            return

        imported = 0
        errors = 0
        folder_cache = {}

        # Temporary directory for Outlook-friendly .eml copies
        temp_dir = Path(tempfile.mkdtemp(prefix="wlm_export_"))

        # CSV results log (one line per message)
        results_csv_path = Path.cwd() / "export_results.csv"
        results_file = open(results_csv_path, mode="w", newline="", encoding="utf-8")
        results_writer = csv.writer(results_file)
        results_writer.writerow(
            ["index", "total", "eml_path", "relative_path", "status", "error"]
        )

        # Configure progress bar on UI thread
        def init_progress():
            progress_bar["maximum"] = total
            progress_bar["value"] = 0

        root_window.after(0, init_progress)

        for idx, eml_path in enumerate(eml_files, start=1):
            # Check for stop request
            if control_flags.get("stop"):
                logging.info("Stop requested by user at item %d/%d", idx, total)
                break

            # Determine relative path once per item (used in status and logging)
            try:
                rel = eml_path.relative_to(root)
            except ValueError:
                rel = eml_path.name

            # Handle pause request
            while control_flags.get("pause") and not control_flags.get("stop"):
                root_window.after(0, status_var.set, f"Paused at {idx}/{total}: {rel}")
                pythoncom.PumpWaitingMessages()
                threading.Event().wait(0.2)

            status_text = f"Importing {idx}/{total}: {rel}"
            root_window.after(0, status_var.set, status_text)

            try:
                # Find / create the matching PST folder structure for this item
                target_folder = ensure_pst_folder_for_eml(
                    import_root, root, eml_path, folder_cache
                )

                # Work around Outlook path/encoding quirks by copying the .eml
                # to a short, ASCII-only temp path before opening.
                temp_name = f"msg_{idx}.eml"
                temp_path = temp_dir / temp_name
                shutil.copy2(str(eml_path), str(temp_path))

                # Try up to 2 times to open and move the message
                last_err = ""
                for attempt in range(2):
                    try:
                        mail_item = session.OpenSharedItem(str(temp_path))
                        copied = mail_item.Copy()
                        copied.Move(target_folder)
                        imported += 1
                        results_writer.writerow(
                            [idx, total, str(eml_path), str(rel), "imported", ""]
                        )
                        break
                    except Exception as e:
                        last_err = repr(e)
                        logging.exception(
                            "Attempt %d failed to import %s via %s", attempt + 1, eml_path, temp_path
                        )
                        if attempt == 1:
                            errors += 1
                            results_writer.writerow(
                                [
                                    idx,
                                    total,
                                    str(eml_path),
                                    str(rel),
                                    "failed",
                                    last_err,
                                ]
                            )
                else:
                    # Should not reach here because of break/attempt logic, but keep for safety
                    errors += 1
            except Exception as e:
                # If Outlook cannot open a given file, skip it
                logging.exception("Failed to import %s", eml_path)
                errors += 1
                results_writer.writerow(
                    [idx, total, str(eml_path), str(rel), "failed", repr(e)]
                )
                continue

            # Update progress bar on UI thread
            def update_progress():
                progress_bar["value"] = idx

            root_window.after(0, update_progress)

        summary = (
            f"Export finished.\n\n"
            f"Total .eml files found: {total}\n"
            f"Successfully imported: {imported}\n"
            f"Skipped / failed: {errors}\n\n"
            f"PST: {pst_full}"
        )

        logging.info(
            "Export complete: total=%d, imported=%d, errors=%d, pst=%s",
            total,
            imported,
            errors,
            pst_full,
        )

        root_window.after(0, status_var.set, "Done.")
        messagebox.showinfo(APP_TITLE, summary)

    except Exception as e:
        tb = traceback.format_exc()
        logging.error("Unexpected error during export: %s\n%s", e, tb)
        messagebox.showerror(
            APP_TITLE,
            f"An unexpected error occurred:\n\n{e}\n\nTraceback:\n{tb}",
        )
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        # Close CSV log
        try:
            if "results_file" in locals() and not results_file.closed:
                results_file.close()
        except Exception:
            logging.warning("Failed to close results CSV", exc_info=True)
        # Best-effort cleanup of temp directory
        try:
            if "temp_dir" in locals() and temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            logging.warning("Failed to clean up temp dir %s", temp_dir, exc_info=True)


def main():
    log_path = Path.cwd() / "export_log.txt"
    logging.basicConfig(
        filename=log_path,
        filemode="a",
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    logging.info("Launcher started")

    root = Tk()
    root.title(APP_TITLE)
    root.geometry("640x200")
    root.resizable(False, False)

    wlm_path_var = StringVar()
    pst_path_var = StringVar()
    status_var = StringVar()
    status_var.set("Select Windows Live Mail folder and target PST file.")

    # Control flags for pause/stop
    control_flags = {"pause": False, "stop": False}

    # Row 0: WLM folder
    Label(root, text="Windows Live Mail folder:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
    Entry(root, textvariable=wlm_path_var, width=60).grid(row=0, column=1, padx=5, pady=10, sticky="w")
    Button(root, text="Browse...", command=lambda: select_wlm_folder(wlm_path_var)).grid(
        row=0, column=2, padx=10, pady=10
    )

    # Row 1: PST file
    Label(root, text="Target PST file:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
    Entry(root, textvariable=pst_path_var, width=60).grid(row=1, column=1, padx=5, pady=10, sticky="w")
    Button(root, text="Browse...", command=lambda: select_pst_file(pst_path_var)).grid(
        row=1, column=2, padx=10, pady=10
    )

    # Row 2: Controls
    def on_export_click():
        wlm_root = wlm_path_var.get().strip()
        pst_file = pst_path_var.get().strip()

        if not wlm_root:
            messagebox.showwarning(APP_TITLE, "Please select the Windows Live Mail folder.")
            return
        if not pst_file:
            messagebox.showwarning(APP_TITLE, "Please select or enter a target PST file.")
            return

        root_path = Path(wlm_root)
        if not root_path.exists():
            messagebox.showerror(APP_TITLE, "The selected Windows Live Mail folder does not exist.")
            return

        pst_path = Path(pst_file)
        control_flags["pause"] = False
        control_flags["stop"] = False
        status_var.set("Starting export… this may take a while. Keep Outlook open.")
        run_export_async(root_path, pst_path, status_var, progress_bar, root, control_flags)

    start_button = Button(root, text="Start Export", command=on_export_click, width=15)
    start_button.grid(row=2, column=1, pady=10, sticky="e")

    # Pause / resume and stop buttons
    def on_pause_resume():
        if control_flags["pause"]:
            control_flags["pause"] = False
            status_var.set("Resuming…")
            pause_button.config(text="Pause")
        else:
            control_flags["pause"] = True
            status_var.set("Pausing after current email…")
            pause_button.config(text="Resume")

    def on_stop():
        control_flags["stop"] = True
        status_var.set("Stopping after current email…")

    pause_button = Button(root, text="Pause", command=on_pause_resume, width=10)
    pause_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")

    stop_button = Button(root, text="Stop", command=on_stop, width=10)
    stop_button.grid(row=2, column=2, padx=10, pady=10, sticky="e")

    # Row 3: Progress bar
    progress_bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", length=400)
    progress_bar.grid(row=3, column=0, columnspan=3, padx=10, pady=(5, 0), sticky="we")

    # Row 4: Status
    Label(root, textvariable=status_var, anchor="w", fg="blue", wraplength=600, justify="left").grid(
        row=4, column=0, columnspan=3, padx=10, pady=10, sticky="w"
    )

    root.mainloop()


if __name__ == "__main__":
    # Ensure we're on Windows
    if sys.platform != "win32":
        print("This exporter only runs on Windows with Outlook installed.", file=sys.stderr)
        sys.exit(1)
    main()


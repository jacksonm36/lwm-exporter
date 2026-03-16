## Windows Live Mail → Outlook PST Exporter

This is a small Windows-only GUI tool that takes a Windows Live Mail message store (folders full of `.eml` files) and imports everything into an Outlook PST file using Outlook's COM automation.

It preserves the **account/folder structure** from the Windows Live Mail export and shows a **progress bar**, the **currently imported email**, and **Pause/Resume/Stop** controls.

### Requirements

- Windows (Win7 or newer)
- Outlook (desktop) installed and configured (32-bit or 64-bit)
- For running from source:  
  - Python 3.8+ (64-bit recommended on modern Windows)

### Recommended Python downloads

- **Python 3.8 (32-bit, for Win7 / 32-bit Outlook)**  
  Download the *Windows x86 executable installer* from the official downloads page:  
  `https://www.python.org/downloads/release/python-3810/`

- **Python 3.14 (64-bit, for modern Win10/11)**  
  Download the *Windows installer (64-bit)* from the latest 3.14.x release page:  
  `https://www.python.org/downloads/windows/`

> Any 3.8+ version is fine for running the script on Win10/11.  
> For **Windows 7**, use a **32-bit Python 3.8** to avoid runtime issues with newer runtimes.

### Setup (run from source)

1. Install a suitable Python version (see above).
2. Open a terminal in this folder.
3. (Optional but recommended) Create and activate a virtual environment.
4. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

### Usage (from source)

1. Make sure Outlook is installed. You can have it open during import.
2. Run the exporter:

   ```bash
   python export_wlm_to_pst.py
   ```

3. In the GUI:
   - **Windows Live Mail folder**: browse to the root folder where your `.eml` files are stored (often the export from Windows Live Mail). Subfolders (accounts, Inbox/Sent, etc.) are mirrored into the PST.
   - **Target PST file**: choose an existing PST or type a new filename (e.g. `C:\Users\you\Documents\WLM-archive.pst`).
   - Click **Start Export**.
   - Use **Pause / Resume** to temporarily stop after the current email, or **Stop** to end the export after the current email.

4. The tool will:
   - Walk all subfolders under the selected Windows Live Mail folder.
   - Find every `.eml` file.
   - Ask Outlook to open each `.eml` and copy it into a folder tree under **Imported from Windows Live Mail** inside the selected PST that reflects your original account/folder layout.

### Building standalone executables

You can build platform-specific `.exe` files so target machines do not need Python installed.

#### 64-bit exe (modern Windows with 64-bit Outlook)

On a Win10/11 machine with 64-bit Python installed:

```bash
py -m pip install pyinstaller pywin32
py -m PyInstaller --onefile --windowed export_wlm_to_pst.py
```

The result will be in the `dist` folder as `export_wlm_to_pst.exe`.

#### Windows 7–compatible 32-bit exe (32-bit Outlook)

On a Win10/11 machine, install **Python 3.8 (32-bit)**:

1. Download the 32-bit 3.8 installer (`Windows x86 executable installer`) from:  
   `https://www.python.org/downloads/release/python-3810/`
2. Install it and ensure it is visible to the `py` launcher (`py -0` shows a `-3.8-32` entry).
3. In this project folder:

   ```bash
   py -3.8-32 -m pip install "pyinstaller<6" pywin32
   py -3.8-32 -m PyInstaller --onefile --windowed export_wlm_to_pst.py
   ```

4. Copy the resulting `dist\export_wlm_to_pst.exe` to the Windows 7 machine (with 32-bit Outlook).

### Notes and limitations

- Outlook must be present on the machine; this script uses the Outlook Object Model (`OpenSharedItem`, PST stores, etc.).
- Some `.eml` files that Outlook cannot open may be skipped; these will be counted as "Skipped / failed" in the final summary and logged to `export_log.txt`.
- Large mail stores can take a long time to import. The progress bar and status line at the bottom of the window show ongoing progress and the currently processed email.



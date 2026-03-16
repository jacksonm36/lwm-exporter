# Windows Live Mail → Outlook PST Exporter (Rust)

Windows-only GUI tool that takes a Windows Live Mail message store (folders of `.eml` files) and imports them into an Outlook PST file. Built with Rust and [Iced](https://github.com/iced-rs/iced).

## Requirements

- **Windows** (with Outlook installed)
- Outlook desktop app (32-bit or 64-bit)

## Build and run

```powershell
cd wlm_exporter_rust
cargo build --release
.\target\release\wlm_exporter_rust.exe
```

Or double-click `wlm_exporter_rust.exe` in `target\release\` after building.

## Usage

1. Run the app and choose:
   - **Windows Live Mail folder** – root folder containing your `.eml` files (and subfolders).
   - **Target PST file** – path to an existing or new `.pst` file.
2. Click **Start export**. Keep Outlook open during the run.
3. Results are written next to the PST as `export_results.csv`; logs go to `export_log.txt` in the current directory.

## Project layout

- `wlm_exporter_rust/` – Rust GUI application (Iced), CSV logging, EML scanning, and Outlook COM stubs.
- `AUDIT.md` – Security and code-quality audit notes.

## Notes

- Outlook COM integration is currently **stubbed**; the app scans and logs but does not yet write messages into the PST. COM implementation can be added using the `windows` crate.
- Maximum 100,000 `.eml` files per run; single files over 50 MiB are skipped in fallback parse.

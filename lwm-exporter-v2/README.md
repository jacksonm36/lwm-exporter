# WLM PST Exporter v2 (Tauri)

Same Windows Live Mail → Outlook PST export as the Iced version, with a **Tauri v2** web-based GUI.

## Requirements

- **Windows** with Outlook installed
- **Node.js** (for Tauri CLI and frontend tooling)
- **Rust** (same as the Iced build)

## Build the v2 .exe

From the repo root (one level up from this folder):

```powershell
# Install Tauri CLI if needed
npm install

# Build the Tauri app (produces .exe in src-tauri/target/release/)
npm run tauri build
```

The executable will be at:

`lwm-exporter-v2\src-tauri\target\release\lwm-exporter-v2.exe`

(Or under `target\release\bundle\msi\` for an installer.)

## Run in development

```powershell
npm run tauri dev
```

## Notes

- Export logic lives in the shared library `wlm_exporter_rust` (used by both the Iced app and this Tauri app).
- The GUI is a single `index.html` with vanilla JS; it uses Tauri’s dialog plugin for folder/file pickers and invokes the Rust `run_export_command` to run the export.

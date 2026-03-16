#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::path::PathBuf;
use std::sync::Once;

use simplelog::{Config as LogConfig, LevelFilter, WriteLogger};
use wlm_exporter_lib::run_export;

static LOG_INIT: Once = Once::new();

fn init_log() {
    LOG_INIT.call_once(|| {
        let _ = std::fs::OpenOptions::new()
            .create(true)
            .append(true)
            .open(
                std::env::current_dir()
                    .unwrap_or_else(|_| PathBuf::from("."))
                    .join("export_log.txt"),
            )
            .and_then(|f| {
                WriteLogger::init(LevelFilter::Info, LogConfig::default(), f).map_err(|e| {
                    eprintln!("Log init failed: {}", e);
                    std::io::Error::new(std::io::ErrorKind::Other, "log init")
                })
            });
    });
}

#[tauri::command]
fn run_export_command(wlm_path: String, pst_path: String) -> Result<String, String> {
    init_log();
    let wlm_path = wlm_path.trim();
    let pst_path = pst_path.trim();
    if wlm_path.is_empty() || pst_path.is_empty() {
        return Err("Source folder and target PST path are required.".to_string());
    }
    let wlm = PathBuf::from(wlm_path);
    let pst = PathBuf::from(pst_path);

    if !wlm.exists() {
        return Err("The selected Windows Live Mail folder does not exist.".to_string());
    }
    if !wlm.is_dir() {
        return Err("The selected source is not a folder.".to_string());
    }
    if let Some(parent) = pst.parent() {
        if !parent.exists() {
            return Err("The target PST folder does not exist. Create it or choose another location.".to_string());
        }
    }
    if pst.extension().map_or(true, |e| !e.eq_ignore_ascii_case("pst")) {
        return Err("Target file must have a .pst extension.".to_string());
    }

    run_export(&wlm, &pst).map_err(|e| format!("{:#}", e))
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    init_log();

    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .invoke_handler(tauri::generate_handler![run_export_command])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

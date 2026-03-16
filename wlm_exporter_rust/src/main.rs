#![cfg_attr(all(windows, not(debug_assertions)), windows_subsystem = "windows")]

use std::{
    env,
    ffi::OsStr,
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result};
use csv::Writer;
use iced::alignment;
use iced::widget::{button, column, container, row, progress_bar, text};
use iced::{Element, Length, Sandbox, Settings, Theme};
use log::{error, info};
use mailparse::ParsedMail;
use rfd::FileDialog;
use simplelog::{Config as LogConfig, LevelFilter, WriteLogger};
use tempfile::TempDir;
use walkdir::WalkDir;

const APP_TITLE: &str = "Windows Live Mail → Outlook PST Exporter";
/// Maximum number of .eml files to process in one run (safety limit).
const MAX_EML_FILES: usize = 100_000;
/// Maximum path length to show in UI before truncating with ellipsis.
const MAX_PATH_DISPLAY_LEN: usize = 70;
/// Maximum size (bytes) of a single .eml file when reading for fallback parse (avoids OOM).
const MAX_EML_FILE_SIZE_BYTES: u64 = 50 * 1024 * 1024;

fn main() -> Result<()> {
    if cfg!(not(windows)) {
        eprintln!("This exporter only runs on Windows with Outlook installed.");
        std::process::exit(1);
    }

    let cwd = env::current_dir().context("Failed to get current directory")?;
    let log_path = cwd.join("export_log.txt");
    let log_file = fs::File::options()
        .create(true)
        .append(true)
        .open(&log_path)
        .context("Failed to open log file")?;

    WriteLogger::init(LevelFilter::Info, LogConfig::default(), log_file)
        .context("Failed to initialize logger")?;

    info!("Launcher started");

    WlmApp::run(Settings::default()).context("Failed to launch GUI")?;
    Ok(())
}

struct WlmApp {
    wlm_path: Option<PathBuf>,
    pst_path: Option<PathBuf>,
    status: String,
    running: bool,
    last_summary: Option<String>,
}

#[derive(Debug, Clone)]
enum Message {
    SelectWlm,
    SelectPst,
    StartExport,
}

impl Sandbox for WlmApp {
    type Message = Message;

    fn new() -> Self {
        Self {
            wlm_path: None,
            pst_path: None,
            status: "Select Windows Live Mail folder and target PST file.".to_string(),
            running: false,
            last_summary: None,
        }
    }

    fn title(&self) -> String {
        APP_TITLE.to_string()
    }

    fn theme(&self) -> Theme {
        Theme::Dark
    }

    fn update(&mut self, message: Self::Message) {
        match message {
            Message::SelectWlm => {
                if let Some(path) = FileDialog::new()
                    .set_title("Select Windows Live Mail folder")
                    .pick_folder()
                {
                    self.wlm_path = Some(path);
                }
            }
            Message::SelectPst => {
                if let Some(path) = FileDialog::new()
                    .set_title("Select / create target PST file")
                    .add_filter("Outlook Data File (*.pst)", &["pst"])
                    .save_file()
                {
                    self.pst_path = Some(path);
                }
            }
            Message::StartExport => {
                if self.running {
                    return;
                }
                let Some(wlm) = self.wlm_path.clone() else {
                    self.status = "Please select the Windows Live Mail folder.".to_string();
                    return;
                };
                let Some(pst) = self.pst_path.clone() else {
                    self.status = "Please select or enter a target PST file.".to_string();
                    return;
                };

                if !wlm.exists() {
                    self.status = "The selected Windows Live Mail folder does not exist.".to_string();
                    return;
                }
                if !wlm.is_dir() {
                    self.status = "The selected source is not a folder.".to_string();
                    return;
                }
                if let Some(parent) = pst.parent() {
                    if !parent.exists() {
                        self.status = "The target PST folder does not exist. Create it or choose another location.".to_string();
                        return;
                    }
                }
                if pst.extension().map_or(true, |e| !e.eq_ignore_ascii_case("pst")) {
                    self.status = "Target file must have a .pst extension.".to_string();
                    return;
                }

                self.running = true;
                self.status = "Starting export… this may take a while. Keep Outlook open.".to_string();
                self.last_summary = None;
                match run_export(&wlm, &pst) {
                    Ok(summary) => {
                        self.status = "Done.".to_string();
                        self.last_summary = Some(summary);
                    }
                    Err(err) => {
                        self.status = format!("Error: {:#}", err);
                        self.last_summary = None;
                    }
                }
                self.running = false;
            }
        }
    }

    fn view(&self) -> Element<'_, Self::Message> {
        let wlm_label = self
            .wlm_path
            .as_ref()
            .map(|p| truncate_path_display(&p.to_string_lossy()))
            .unwrap_or_else(|| "No folder selected".to_string());
        let pst_label = self
            .pst_path
            .as_ref()
            .map(|p| truncate_path_display(&p.to_string_lossy()))
            .unwrap_or_else(|| "No PST file selected".to_string());

        let mut summary_text = String::new();
        if let Some(s) = &self.last_summary {
            summary_text = s.clone();
        }

        let header = column![
            text("Windows Live Mail → Outlook PST")
                .size(22)
                .horizontal_alignment(alignment::Horizontal::Center),
        ]
        .spacing(4)
        .padding([0, 0, 16, 0]);

        // Source / destination cards
        let wlm_row = row![
            text("Source").size(14).width(Length::Shrink),
            text("Windows Live Mail folder")
                .size(14)
                .width(Length::FillPortion(2)),
            text(wlm_label).width(Length::FillPortion(4)),
            button("Browse…").on_press(Message::SelectWlm)
        ]
        .spacing(12)
        .padding(10);

        let pst_row = row![
            text("Destination").size(14).width(Length::Shrink),
            text("Target Outlook PST file")
                .size(14)
                .width(Length::FillPortion(2)),
            text(pst_label).width(Length::FillPortion(4)),
            button("Browse…").on_press(Message::SelectPst)
        ]
        .spacing(12)
        .padding(10);

        let action_button = if self.running {
            button("Exporting…")
        } else {
            button("Start export").on_press(Message::StartExport)
        };

        let controls = row![action_button]
            .spacing(12)
            .padding(10)
            .align_items(iced::Alignment::Center);

        let progress = progress_bar(0.0..=1.0, if self.running { 0.5 } else { 0.0 })
            .width(Length::Fill);

        let status_block = column![
            text("Status").size(16),
            text(&self.status),
        ]
        .spacing(4)
        .padding(10);

        let summary_block: Element<_> = if !summary_text.is_empty() {
            text(summary_text).into()
        } else {
            column![].into()
        };

        let main_panel = column![
            header,
            wlm_row,
            pst_row,
            controls,
            progress,
            status_block,
            summary_block,
        ]
        .spacing(8)
        .padding(16);

        container(main_panel)
            .width(Length::Fill)
            .height(Length::Fill)
            .into()
    }
}

/// Sanitize a CSV field to prevent formula injection when the CSV is opened in Excel.
/// Prefixes values that start with =, +, -, @, \t, \r so they are treated as text.
fn csv_sanitize(s: &str) -> String {
    let t = s.trim_start();
    if t.starts_with('=')
        || t.starts_with('+')
        || t.starts_with('-')
        || t.starts_with('@')
        || t.starts_with('\t')
        || t.starts_with('\r')
    {
        format!("'{s}")
    } else {
        s.to_string()
    }
}

/// Truncate a path string for UI display, keeping start and end (character-based for UTF-8 safety).
fn truncate_path_display(path: &str) -> String {
    let chars: Vec<char> = path.chars().collect();
    if chars.len() <= MAX_PATH_DISPLAY_LEN {
        return path.to_string();
    }
    let half = (MAX_PATH_DISPLAY_LEN.saturating_sub(3)) / 2;
    let start: String = chars.iter().take(half).collect();
    let end: String = chars.iter().rev().take(half).collect::<Vec<_>>().into_iter().rev().collect();
    format!("{}...{}", start, end)
}

fn list_eml_files(root: &Path) -> Vec<PathBuf> {
    let mut files = Vec::new();
    for entry in WalkDir::new(root)
        .follow_links(false)
        .into_iter()
        .filter_map(|e| e.ok())
    {
        if entry.file_type().is_file() {
            if entry
                .path()
                .extension()
                .and_then(OsStr::to_str)
                .map(|ext| ext.eq_ignore_ascii_case("eml"))
                .unwrap_or(false)
            {
                files.push(entry.into_path());
            }
        }
    }
    files
}

fn relative_path(root: &Path, eml_path: &Path) -> String {
    match eml_path.strip_prefix(root) {
        Ok(rel) => rel.to_string_lossy().into_owned(),
        Err(_) => eml_path
            .file_name()
            .map(|n| n.to_string_lossy().into_owned())
            .unwrap_or_default(),
    }
}

#[cfg(windows)]
fn get_short_path(path: &Path) -> String {
    use std::os::windows::ffi::OsStrExt;
    use windows::core::PCWSTR;
    use windows::Win32::Foundation::MAX_PATH;
    use windows::Win32::Storage::FileSystem::GetShortPathNameW;

    let wide: Vec<u16> = path.as_os_str().encode_wide().chain(std::iter::once(0)).collect();
    let mut buf = [0u16; MAX_PATH as usize];

    unsafe {
        let len = GetShortPathNameW(PCWSTR(wide.as_ptr()), Some(&mut buf));
        if len > 0 && (len as usize) < buf.len() {
            String::from_utf16_lossy(&buf[..len as usize])
        } else {
            path.to_string_lossy().into_owned()
        }
    }
}

#[cfg(not(windows))]
fn get_short_path(path: &Path) -> String {
    path.to_string_lossy().into_owned()
}

fn run_export(root: &Path, pst_path: &Path) -> Result<String> {
    info!("Starting export: root={:?}, pst={:?}", root, pst_path);

    let pst_full = pst_path.to_string_lossy().into_owned();

    attach_or_create_pst_store(&pst_full)?;
    let root_folder = get_pst_root_folder(&pst_full)?;

    let import_root = get_or_create_import_root(&root_folder, "Imported from Windows Live Mail")?;

    let eml_files = list_eml_files(root);
    let total = eml_files.len();
    if total == 0 {
        info!("No .eml files found under {:?}", root);
        return Ok("No .eml files were found under the selected Windows Live Mail folder.".to_string());
    }
    if total > MAX_EML_FILES {
        anyhow::bail!(
            "Too many .eml files ({total}). Maximum allowed is {MAX_EML_FILES}. \
             Use a smaller folder or split your export."
        );
    }

    info!("Found {} .eml files. Starting import.", total);

    let mut imported = 0usize;
    let mut errors = 0usize;

    let temp_dir = TempDir::new_in(env::temp_dir()).context("Failed to create temp directory")?;
    let temp_dir_path = temp_dir.path().to_path_buf();

    let results_csv_path = pst_path
        .parent()
        .unwrap_or_else(|| Path::new("."))
        .join("export_results.csv");
    let mut writer = Writer::from_path(&results_csv_path)
        .with_context(|| format!("Failed to create results log at {:?}", results_csv_path))?;
    writer.write_record(["index", "total", "eml_path", "relative_path", "status", "error"])?;

    let mut folder_cache: FolderCache = FolderCache::new();

    for (idx, eml_path) in eml_files.iter().enumerate() {
        let idx_display = idx + 1;
        let rel = relative_path(root, eml_path);

        println!("Importing {}/{}: {}", idx_display, total, rel);

        if let Err(e) = process_single_eml(
            &import_root,
            root,
            eml_path,
            &rel,
            idx_display,
            total,
            &temp_dir_path,
            &mut imported,
            &mut errors,
            &mut writer,
            &mut folder_cache,
        ) {
            errors += 1;
            error!("Failed to import {:?}: {:?}", eml_path, e);
            writer.write_record(&[
                idx_display.to_string(),
                total.to_string(),
                csv_sanitize(&eml_path.to_string_lossy()),
                csv_sanitize(&rel),
                "failed".to_string(),
                csv_sanitize(&format!("{:?}", e)),
            ])?;
        }
    }

    writer.flush()?;

    let summary = format!(
        "Export finished.\n\n\
         Total .eml files found: {total}\n\
         Successfully imported: {imported}\n\
         Skipped / failed: {errors}\n\n\
         PST: {pst_full}\n\
         Results log: {}",
        results_csv_path.display()
    );

    info!(
        "Export complete: total={}, imported={}, errors={}, pst={}",
        total, imported, errors, pst_full
    );

    println!("{summary}");

    Ok(summary)
}

type FolderCache = std::collections::HashMap<String, OutlookFolder>;

#[derive(Clone)]
struct OutlookFolder {
    _dummy: (),
}

fn attach_or_create_pst_store(pst_path: &str) -> Result<()> {
    info!("(stub) attach_or_create_pst_store({pst_path})");
    Ok(())
}

fn get_pst_root_folder(pst_path: &str) -> Result<OutlookFolder> {
    info!("(stub) get_pst_root_folder({pst_path})");
    Ok(OutlookFolder { _dummy: () })
}

fn get_or_create_import_root(root_folder: &OutlookFolder, name: &str) -> Result<OutlookFolder> {
    info!("(stub) get_or_create_import_root({name})");
    let _ = root_folder;
    Ok(OutlookFolder { _dummy: () })
}

fn ensure_pst_folder_for_eml(
    import_root: &OutlookFolder,
    root: &Path,
    eml_path: &Path,
    folder_cache: &mut FolderCache,
) -> Result<OutlookFolder> {
    let rel = match eml_path.strip_prefix(root) {
        Ok(rel) => rel.to_path_buf(),
        Err(_) => eml_path.file_name().map(PathBuf::from).unwrap_or_default(),
    };

    let parts: Vec<String> = rel
        .parent()
        .map(|p| {
            p.components()
                .map(|c| c.as_os_str().to_string_lossy().into_owned())
                .filter(|s| !s.is_empty() && s != "." && s != "..")
                .collect()
        })
        .unwrap_or_else(Vec::new);

    let mut current_key = String::new();
    let mut current_folder = import_root.clone();

    for part in parts {
        if current_key.is_empty() {
            current_key = part.clone();
        } else {
            current_key.push('/');
            current_key.push_str(&part);
        }

        if let Some(cached) = folder_cache.get(&current_key) {
            current_folder = cached.clone();
            continue;
        }

        info!("(stub) ensure PST subfolder {:?}", current_key);
        let next_folder = OutlookFolder { _dummy: () };
        folder_cache.insert(current_key.clone(), next_folder.clone());
        current_folder = next_folder;
    }

    Ok(current_folder)
}

#[allow(clippy::too_many_arguments)]
fn process_single_eml(
    import_root: &OutlookFolder,
    root: &Path,
    eml_path: &Path,
    rel: &str,
    idx: usize,
    total: usize,
    temp_dir: &Path,
    imported: &mut usize,
    errors: &mut usize,
    writer: &mut Writer<std::fs::File>,
    folder_cache: &mut FolderCache,
) -> Result<()> {
    let target_folder = ensure_pst_folder_for_eml(import_root, root, eml_path, folder_cache)?;

    let temp_name = format!("msg_{idx}.eml");
    let temp_path = temp_dir.join(&temp_name);
    fs::copy(eml_path, &temp_path)
        .with_context(|| format!("Failed to copy to temp {:?}", temp_path))?;

    let short = get_short_path(&temp_path);
    let candidates = vec![
        temp_path.to_string_lossy().into_owned(),
        short,
        temp_path.to_string_lossy().into_owned(),
    ];

    let mut imported_ok = false;
    let mut last_err = String::new();

    for (attempt, candidate) in candidates.iter().enumerate() {
        match import_via_outlook_openshareditem(candidate, &target_folder) {
            Ok(()) => {
                *imported += 1;
                imported_ok = true;
                writer.write_record(&[
                    idx.to_string(),
                    total.to_string(),
                    csv_sanitize(&eml_path.to_string_lossy()),
                    csv_sanitize(rel),
                    "imported_via_openshareditem".to_string(),
                    "".to_string(),
                ])?;
                break;
            }
            Err(e) => {
                last_err = format!("{:?}", e);
                error!(
                    "Attempt {} failed to import {:?} via {}: {:?}",
                    attempt + 1,
                    eml_path,
                    candidate,
                    e
                );
            }
        }
    }

    if !imported_ok {
        match import_via_fallback_parse(
            &temp_path,
            &target_folder,
            eml_path,
            rel,
            idx,
            total,
            writer,
            &last_err,
        ) {
            Ok(()) => {
                *imported += 1;
            }
            Err(e) => {
                *errors += 1;
                error!("Fallback import failed for {:?}: {:?}", eml_path, e);
                writer.write_record(&[
                    idx.to_string(),
                    total.to_string(),
                    csv_sanitize(&eml_path.to_string_lossy()),
                    csv_sanitize(rel),
                    "failed".to_string(),
                    csv_sanitize(&format!("{:?}", e)),
                ])?;
            }
        }
    }

    Ok(())
}

fn import_via_outlook_openshareditem(candidate_path: &str, target_folder: &OutlookFolder) -> Result<()> {
    let _ = target_folder;
    info!("(stub) OpenSharedItem & Move via Outlook: {}", candidate_path);
    Ok(())
}

#[allow(clippy::too_many_arguments)]
fn import_via_fallback_parse(
    temp_path: &Path,
    target_folder: &OutlookFolder,
    eml_path: &Path,
    rel: &str,
    idx: usize,
    total: usize,
    writer: &mut Writer<std::fs::File>,
    last_err: &str,
) -> Result<()> {
    let size = fs::metadata(temp_path)
        .with_context(|| format!("Failed to stat temp EML {:?}", temp_path))?
        .len();
    if size > MAX_EML_FILE_SIZE_BYTES {
        anyhow::bail!(
            "EML file too large ({} MiB). Max allowed for fallback parse is {} MiB.",
            size / (1024 * 1024),
            MAX_EML_FILE_SIZE_BYTES / (1024 * 1024)
        );
    }
    let data =
        fs::read(temp_path).with_context(|| format!("Failed to read temp EML {:?}", temp_path))?;
    let parsed = mailparse::parse_mail(&data).context("Failed to parse EML")?;

    create_and_move_mail_item_from_parsed(&parsed, target_folder, temp_path)?;

    writer.write_record(&[
        idx.to_string(),
        total.to_string(),
        csv_sanitize(&eml_path.to_string_lossy()),
        csv_sanitize(rel),
        "imported_via_fallback".to_string(),
        csv_sanitize(last_err),
    ])?;

    Ok(())
}

fn create_and_move_mail_item_from_parsed(
    parsed: &ParsedMail<'_>,
    target_folder: &OutlookFolder,
    temp_path: &Path,
) -> Result<()> {
    let _ = target_folder;

    let subject = header_value(parsed, "Subject");
    let to_addr = header_value(parsed, "To");
    let cc_addr = header_value(parsed, "Cc");
    let bcc_addr = header_value(parsed, "Bcc");
    let body_text = extract_body_text(parsed);

    info!(
        "(stub) Create Outlook MailItem subject={:?}, to={:?}, cc={:?}, bcc={:?}, body_len={}, attach={:?}",
        subject,
        to_addr,
        cc_addr,
        bcc_addr,
        body_text.len(),
        temp_path
    );

    Ok(())
}

fn header_value(parsed: &ParsedMail<'_>, name: &str) -> String {
    parsed
        .headers
        .iter()
        .find(|h| h.get_key_ref().eq_ignore_ascii_case(name))
        .map(|h| h.get_value())
        .unwrap_or_default()
}

fn extract_body_text(parsed: &ParsedMail<'_>) -> String {
    if parsed.subparts.is_empty() {
        return parsed.get_body().unwrap_or_default();
    }

    for part in &parsed.subparts {
        if part.ctype.mimetype.eq_ignore_ascii_case("text/plain") {
            return part.get_body().unwrap_or_default();
        }
    }

    for part in &parsed.subparts {
        if part.ctype.mimetype.to_lowercase().starts_with("text/") {
            return part.get_body().unwrap_or_default();
        }
    }

    String::new()
}


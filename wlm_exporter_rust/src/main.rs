#![cfg_attr(all(windows, not(debug_assertions)), windows_subsystem = "windows")]

use std::{env, fs, path::PathBuf};

use anyhow::{Context, Result};
use iced::alignment;
use iced::widget::{button, column, container, row, progress_bar, text};
use iced::{Element, Length, Sandbox, Settings, Theme};
use log::info;
use rfd::FileDialog;
use simplelog::{Config as LogConfig, LevelFilter, WriteLogger};

use wlm_exporter_lib::run_export;

const APP_TITLE: &str = "Windows Live Mail → Outlook PST Exporter";
/// Maximum path length to show in UI before truncating with ellipsis.
const MAX_PATH_DISPLAY_LEN: usize = 70;

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


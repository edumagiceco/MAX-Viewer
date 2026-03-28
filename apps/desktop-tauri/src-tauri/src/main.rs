#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::{fs, path::Path};

use max_viewer_core::{
    APP_NAME, Block, Document, DocumentDiagnostics, DocumentMetadata, FormatInspector,
    FormatSupport, Paragraph, Section, TextRun,
};
use max_viewer_export::to_plain_text;
use max_viewer_hwp::HwpInspector;
use max_viewer_hwpx::HwpxInspector;
use max_viewer_layout::{LayoutSummary, summarize};
use serde::Serialize;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct DesktopAppInfo {
    app_name: String,
    app_version: String,
    shell: String,
    parser_crates: Vec<String>,
    supported_formats: Vec<FormatSupport>,
    roadmap: Vec<String>,
    sample_layout: LayoutPreview,
    sample_export: String,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct LayoutPreview {
    section_count: usize,
    paragraph_count: usize,
    table_count: usize,
    image_count: usize,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
struct LoadedDocument {
    file_name: String,
    document: Document,
    diagnostics: DocumentDiagnostics,
    layout: LayoutSummary,
    plain_text: String,
}

#[tauri::command]
fn app_info() -> DesktopAppInfo {
    let sample_document = Document {
        format: None,
        metadata: DocumentMetadata {
            title: Some("Scaffold Preview".to_string()),
            ..DocumentMetadata::default()
        },
        sections: vec![Section {
            id: 0,
            blocks: vec![Block::Paragraph(Paragraph {
                marker: None,
                runs: vec![TextRun {
                    text: "MAX Viewer workspace is wired and ready for parser work.".to_string(),
                    style: None,
                }],
                style: None,
                line_segment_count: None,
                layout_height_hint: None,
                page_break_before: false,
            })],
            page_layout: None,
            headers: Vec::new(),
            footers: Vec::new(),
            page_start_number: None,
        }],
        assets: Vec::new(),
    };

    let layout = summarize(&sample_document);

    DesktopAppInfo {
        app_name: APP_NAME.to_string(),
        app_version: env!("CARGO_PKG_VERSION").to_string(),
        shell: "Tauri 2 desktop shell".to_string(),
        parser_crates: vec![
            "max_viewer_core".to_string(),
            "max_viewer_hwpx".to_string(),
            "max_viewer_hwp".to_string(),
            "max_viewer_layout".to_string(),
            "max_viewer_export".to_string(),
        ],
        supported_formats: vec![
            HwpxInspector::scaffold_support(),
            HwpInspector::scaffold_support(),
        ],
        roadmap: vec![
            "Add shared style and numbering resolution for HWPX".to_string(),
            "Extend HWP parsing from minimal BodyText preview to DocInfo and control decoding"
                .to_string(),
            "Add drag and drop loading alongside the native open dialog".to_string(),
            "Improve inline image and paragraph style fidelity".to_string(),
        ],
        sample_layout: LayoutPreview {
            section_count: layout.section_count,
            paragraph_count: layout.paragraph_count,
            table_count: layout.table_count,
            image_count: layout.image_count,
        },
        sample_export: to_plain_text(&sample_document),
    }
}

#[tauri::command]
fn open_document(file_name: String, bytes: Vec<u8>) -> Result<LoadedDocument, String> {
    load_document(&file_name, &bytes)
}

#[tauri::command]
fn pick_and_open_document() -> Result<Option<LoadedDocument>, String> {
    let Some(path) = rfd::FileDialog::new()
        .add_filter("Hancom documents", &["hwp", "hwpx"])
        .pick_file()
    else {
        return Ok(None);
    };

    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("document")
        .to_string();
    let bytes = fs::read(&path).map_err(|error| error.to_string())?;

    load_document(&file_name, &bytes).map(Some)
}

fn load_document(file_name: &str, bytes: &[u8]) -> Result<LoadedDocument, String> {
    let extension = Path::new(&file_name)
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase());

    let fallback_title = Path::new(&file_name)
        .file_stem()
        .and_then(|stem| stem.to_str());

    let hwpx = HwpxInspector;
    let hwp = HwpInspector;

    let (document, diagnostics) = match extension.as_deref() {
        Some("hwpx") => {
            let parsed = hwpx
                .parse_bytes(&bytes, fallback_title)
                .map_err(|error| error.to_string())?;
            (parsed.document, parsed.diagnostics)
        }
        Some("hwp") => {
            let parsed = hwp
                .parse_bytes(&bytes, fallback_title)
                .map_err(|error| error.to_string())?;
            (parsed.document, parsed.diagnostics)
        }
        _ => {
            if hwpx.inspect_bytes(&bytes).is_ok() {
                let parsed = hwpx
                    .parse_bytes(&bytes, fallback_title)
                    .map_err(|error| error.to_string())?;
                (parsed.document, parsed.diagnostics)
            } else if hwp.inspect_bytes(&bytes).is_ok() {
                let parsed = hwp
                    .parse_bytes(&bytes, fallback_title)
                    .map_err(|error| error.to_string())?;
                (parsed.document, parsed.diagnostics)
            } else {
                return Err("Unsupported document type. Select a .hwpx or .hwp file.".to_string());
            }
        }
    };

    let layout = summarize(&document);
    let plain_text = to_plain_text(&document);

    Ok(LoadedDocument {
        file_name: file_name.to_string(),
        document,
        diagnostics,
        layout,
        plain_text,
    })
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            app_info,
            open_document,
            pick_and_open_document
        ])
        .run(tauri::generate_context!())
        .expect("failed to run MAX-Viewer desktop shell");
}

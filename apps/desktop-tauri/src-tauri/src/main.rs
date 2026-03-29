#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::{
    fs,
    path::Path,
    process::Command,
    time::{SystemTime, UNIX_EPOCH},
};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use max_viewer_core::{
    APP_NAME, Block, Document, DocumentDiagnostics, DocumentMetadata, FormatInspector,
    FormatSupport, Paragraph, Section, TextRun,
};
use max_viewer_export::to_plain_text;
use max_viewer_hwp::HwpInspector;
use max_viewer_hwpx::HwpxInspector;
use max_viewer_markdown::MarkdownInspector;
use max_viewer_pdf::PdfInspector;
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
    file_path: Option<String>,
    source_text: Option<String>,
    binary_data_uri: Option<String>,
    binary_mime_type: Option<String>,
    is_editable: bool,
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
            "max_viewer_markdown".to_string(),
            "max_viewer_pdf".to_string(),
            "max_viewer_layout".to_string(),
            "max_viewer_export".to_string(),
        ],
        supported_formats: vec![
            HwpxInspector::scaffold_support(),
            HwpInspector::scaffold_support(),
            MarkdownInspector::scaffold_support(),
            PdfInspector::scaffold_support(),
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
    load_document(&file_name, None, &bytes)
}

#[tauri::command]
fn open_document_path(file_path: String) -> Result<LoadedDocument, String> {
    let path = Path::new(&file_path);
    let bytes = fs::read(path).map_err(|error| error.to_string())?;
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("document")
        .to_string();

    load_document(&file_name, Some(file_path), &bytes)
}

fn extract_pdf_search_text_with_lopdf(bytes: &[u8]) -> Result<Vec<String>, String> {
    PdfInspector
        .extract_page_texts(bytes)
        .map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
fn extract_pdf_search_text_with_pdfkit(file_path: &Path) -> Result<Vec<String>, String> {
    const SCRIPT: &str = r#"
ObjC.import('Foundation');
ObjC.import('PDFKit');

const env = $.NSProcessInfo.processInfo.environment;
const path = ObjC.unwrap(env.objectForKey('MAX_VIEWER_PDF_PATH'));
if (!path) {
  throw new Error('MAX_VIEWER_PDF_PATH is missing');
}

const url = $.NSURL.fileURLWithPath(path);
const document = $.PDFDocument.alloc.initWithURL(url);
if (!document) {
  throw new Error(`Unable to open PDF: ${path}`);
}

const pages = [];
const pageCount = Number(document.pageCount);
for (let index = 0; index < pageCount; index += 1) {
  const page = document.pageAtIndex(index);
  const text = page ? String(ObjC.unwrap(page.string || '')) : '';
  pages.push(text);
}

JSON.stringify(pages);
"#;

    let output = Command::new("osascript")
        .arg("-l")
        .arg("JavaScript")
        .arg("-e")
        .arg(SCRIPT)
        .env("MAX_VIEWER_PDF_PATH", file_path)
        .output()
        .map_err(|error| error.to_string())?;

    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr).trim().to_string();
        if stderr.is_empty() {
            return Err(format!("osascript exited with status {}", output.status));
        }
        return Err(stderr);
    }

    let stdout = String::from_utf8(output.stdout).map_err(|error| error.to_string())?;
    serde_json::from_str(stdout.trim()).map_err(|error| error.to_string())
}

fn extract_pdf_search_text_from_path(file_path: &Path) -> Result<Vec<String>, String> {
    #[cfg(target_os = "macos")]
    {
        if let Ok(pages) = extract_pdf_search_text_with_pdfkit(file_path) {
            return Ok(pages);
        }
    }

    let bytes = fs::read(file_path).map_err(|error| error.to_string())?;
    extract_pdf_search_text_with_lopdf(&bytes)
}

fn extract_pdf_search_text_from_bytes(bytes: &[u8]) -> Result<Vec<String>, String> {
    #[cfg(target_os = "macos")]
    {
        let file_name = format!(
            "max-viewer-search-{}-{}.pdf",
            std::process::id(),
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_nanos()
        );
        let temp_path = std::env::temp_dir().join(file_name);
        fs::write(&temp_path, bytes).map_err(|error| error.to_string())?;
        let result = extract_pdf_search_text_from_path(&temp_path);
        let _ = fs::remove_file(&temp_path);
        return result;
    }

    #[cfg(not(target_os = "macos"))]
    {
        extract_pdf_search_text_with_lopdf(bytes)
    }
}

#[tauri::command]
fn extract_pdf_search_text(file_path: String) -> Result<Vec<String>, String> {
    extract_pdf_search_text_from_path(Path::new(&file_path))
}

#[tauri::command]
fn extract_pdf_search_text_bytes(bytes: Vec<u8>) -> Result<Vec<String>, String> {
    extract_pdf_search_text_from_bytes(&bytes)
}

#[tauri::command]
fn parse_markdown_text(
    file_name: String,
    file_path: Option<String>,
    source_text: String,
) -> Result<LoadedDocument, String> {
    let inspector = MarkdownInspector;
    let parsed = inspector
        .parse_bytes_with_base_dir(
            source_text.as_bytes(),
            Path::new(&file_name)
                .file_stem()
                .and_then(|stem| stem.to_str()),
            file_path
                .as_deref()
                .and_then(|path| Path::new(path).parent()),
        )
        .map_err(|error| error.to_string())?;

    Ok(build_loaded_document(
        file_name,
        file_path,
        parsed.document,
        parsed.diagnostics,
        Some(source_text),
        None,
        None,
        true,
    ))
}

#[tauri::command]
fn save_markdown_document(
    file_path: String,
    source_text: String,
) -> Result<LoadedDocument, String> {
    fs::write(&file_path, source_text.as_bytes()).map_err(|error| error.to_string())?;
    let file_name = Path::new(&file_path)
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("document.md")
        .to_string();

    parse_markdown_text(file_name, Some(file_path), source_text)
}

#[tauri::command]
fn save_markdown_document_as(
    suggested_file_name: String,
    source_text: String,
) -> Result<Option<LoadedDocument>, String> {
    let default_name = if suggested_file_name.ends_with(".md")
        || suggested_file_name.ends_with(".markdown")
    {
        suggested_file_name
    } else {
        format!("{suggested_file_name}.md")
    };

    let Some(path) = rfd::FileDialog::new()
        .add_filter("Markdown", &["md", "markdown"])
        .set_file_name(&default_name)
        .save_file()
    else {
        return Ok(None);
    };

    let file_path = path.to_string_lossy().into_owned();
    fs::write(&file_path, source_text.as_bytes()).map_err(|error| error.to_string())?;
    let file_name = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or("document.md")
        .to_string();

    parse_markdown_text(file_name, Some(file_path), source_text).map(Some)
}

#[tauri::command]
fn print_current_document(webview_window: tauri::WebviewWindow) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    {
        webview_window.print().map_err(|error| error.to_string())
    }

    #[cfg(not(target_os = "macos"))]
    {
        webview_window
            .eval("window.print()")
            .map_err(|error| error.to_string())
    }
}

#[tauri::command]
fn pick_and_open_document() -> Result<Option<LoadedDocument>, String> {
    let Some(path) = rfd::FileDialog::new()
        .add_filter("Supported documents", &["hwp", "hwpx", "md", "markdown", "pdf"])
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

    load_document(
        &file_name,
        Some(path.to_string_lossy().into_owned()),
        &bytes,
    )
    .map(Some)
}

fn load_document(
    file_name: &str,
    file_path: Option<String>,
    bytes: &[u8],
) -> Result<LoadedDocument, String> {
    let extension = Path::new(&file_name)
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase());

    let fallback_title = Path::new(&file_name)
        .file_stem()
        .and_then(|stem| stem.to_str());

    let hwpx = HwpxInspector;
    let hwp = HwpInspector;
    let markdown = MarkdownInspector;
    let pdf = PdfInspector;

    let (document, diagnostics, source_text, binary_data_uri, binary_mime_type, is_editable) =
        match extension.as_deref() {
        Some("hwpx") => {
            let parsed = hwpx
                .parse_bytes(&bytes, fallback_title)
                .map_err(|error| error.to_string())?;
            (parsed.document, parsed.diagnostics, None, None, None, false)
        }
        Some("hwp") => {
            let parsed = hwp
                .parse_bytes(&bytes, fallback_title)
                .map_err(|error| error.to_string())?;
            (parsed.document, parsed.diagnostics, None, None, None, false)
        }
        Some("md") | Some("markdown") => {
            let source_text =
                String::from_utf8(bytes.to_vec()).map_err(|error| error.to_string())?;
            let parsed = markdown
                .parse_bytes_with_base_dir(
                    source_text.as_bytes(),
                    fallback_title,
                    file_path
                        .as_deref()
                        .and_then(|path| Path::new(path).parent()),
                )
                .map_err(|error| error.to_string())?;
            (
                parsed.document,
                parsed.diagnostics,
                Some(source_text),
                None,
                None,
                true,
            )
        }
        Some("pdf") => {
            let parsed = pdf
                .parse_bytes(bytes, fallback_title)
                .map_err(|error| error.to_string())?;
            (
                parsed.document,
                parsed.diagnostics,
                None,
                Some(pdf_data_uri(bytes)),
                Some("application/pdf".to_string()),
                false,
            )
        }
        _ => {
            if hwpx.inspect_bytes(&bytes).is_ok() {
                let parsed = hwpx
                    .parse_bytes(&bytes, fallback_title)
                    .map_err(|error| error.to_string())?;
                (parsed.document, parsed.diagnostics, None, None, None, false)
            } else if hwp.inspect_bytes(&bytes).is_ok() {
                let parsed = hwp
                    .parse_bytes(&bytes, fallback_title)
                    .map_err(|error| error.to_string())?;
                (parsed.document, parsed.diagnostics, None, None, None, false)
            } else if pdf.inspect_bytes(bytes).is_ok() {
                let parsed = pdf
                    .parse_bytes(bytes, fallback_title)
                    .map_err(|error| error.to_string())?;
                (
                    parsed.document,
                    parsed.diagnostics,
                    None,
                    Some(pdf_data_uri(bytes)),
                    Some("application/pdf".to_string()),
                    false,
                )
            } else {
                return Err(
                    "Unsupported document type. Select a .hwpx, .hwp, .md, or .pdf file."
                        .to_string(),
                );
            }
        }
    };

    Ok(build_loaded_document(
        file_name.to_string(),
        file_path,
        document,
        diagnostics,
        source_text,
        binary_data_uri,
        binary_mime_type,
        is_editable,
    ))
}

fn build_loaded_document(
    file_name: String,
    file_path: Option<String>,
    document: Document,
    diagnostics: DocumentDiagnostics,
    source_text: Option<String>,
    binary_data_uri: Option<String>,
    binary_mime_type: Option<String>,
    is_editable: bool,
) -> LoadedDocument {
    let layout = summarize(&document);
    let plain_text = to_plain_text(&document);

    LoadedDocument {
        file_name,
        file_path,
        source_text,
        binary_data_uri,
        binary_mime_type,
        is_editable,
        document,
        diagnostics,
        layout,
        plain_text,
    }
}

fn main() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            app_info,
            open_document,
            open_document_path,
            extract_pdf_search_text,
            extract_pdf_search_text_bytes,
            parse_markdown_text,
            save_markdown_document,
            save_markdown_document_as,
            print_current_document,
            pick_and_open_document
        ])
        .run(tauri::generate_context!())
        .expect("failed to run MAX-Viewer desktop shell");
}

fn pdf_data_uri(bytes: &[u8]) -> String {
    format!(
        "data:application/pdf;base64,{}",
        BASE64_STANDARD.encode(bytes)
    )
}

use wasm_bindgen::prelude::*;

/// Parse an HWP, HWPX, Markdown, or PDF document from raw bytes and return a JSON document model.
#[wasm_bindgen]
pub fn parse_document(file_name: &str, bytes: &[u8]) -> Result<String, JsValue> {
    let ext = file_name.rsplit('.').next().unwrap_or("").to_lowercase();

    let document = match ext.as_str() {
        "hwpx" => {
            let inspector = max_viewer_hwpx::HwpxInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "hwp" => {
            let inspector = max_viewer_hwp::HwpInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "md" | "markdown" => {
            let inspector = max_viewer_markdown::MarkdownInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "pdf" => {
            let inspector = max_viewer_pdf::PdfInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        _ => return Err(JsValue::from_str(&format!("unsupported format: .{ext}"))),
    };

    serde_json::to_string(&document).map_err(|e| JsValue::from_str(&e.to_string()))
}

/// Extract plain text from an HWP, HWPX, Markdown, or PDF document.
#[wasm_bindgen]
pub fn extract_text(file_name: &str, bytes: &[u8]) -> Result<String, JsValue> {
    let ext = file_name.rsplit('.').next().unwrap_or("").to_lowercase();

    let document = match ext.as_str() {
        "hwpx" => {
            let inspector = max_viewer_hwpx::HwpxInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "hwp" => {
            let inspector = max_viewer_hwp::HwpInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "md" | "markdown" => {
            let inspector = max_viewer_markdown::MarkdownInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        "pdf" => {
            let inspector = max_viewer_pdf::PdfInspector;
            inspector
                .parse_bytes(bytes, Some(file_name))
                .map(|r| r.document)
                .map_err(|e| JsValue::from_str(&e.to_string()))?
        }
        _ => return Err(JsValue::from_str(&format!("unsupported format: .{ext}"))),
    };

    Ok(max_viewer_export::to_plain_text(&document))
}

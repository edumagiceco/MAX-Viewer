use lopdf::{Document as LoPdfDocument, Object};
use max_viewer_core::{
    Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata, FormatInspector,
    FormatSupport, ParseError,
};

#[derive(Debug, Default)]
pub struct PdfInspector;

#[derive(Debug, Clone)]
pub struct PdfParseResult {
    pub document: Document,
    pub diagnostics: DocumentDiagnostics,
}

impl PdfInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Pdf,
            status: "active".to_string(),
            implemented: vec![
                "PDF metadata and page count detection".to_string(),
                "desktop PDF page rendering via the viewer shell".to_string(),
            ],
            planned: vec![
                "plain text extraction from PDF content streams".to_string(),
                "outline/bookmark extraction".to_string(),
            ],
        }
    }

    pub fn parse_bytes(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
    ) -> Result<PdfParseResult, ParseError> {
        let diagnostics = self.inspect_bytes(bytes)?;
        let pdf = LoPdfDocument::load_mem(bytes)
            .map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let page_count = pdf.get_pages().len() as u32;

        Ok(PdfParseResult {
            document: Document {
                format: Some(DocumentFormat::Pdf),
                metadata: DocumentMetadata {
                    title: extract_info_string(&pdf, b"Title")
                        .or_else(|| fallback_title.map(ToOwned::to_owned)),
                    author: extract_info_string(&pdf, b"Author"),
                    page_count: Some(page_count),
                    language: Some("pdf".to_string()),
                    ..DocumentMetadata::default()
                },
                sections: Vec::new(),
                assets: Vec::new(),
            },
            diagnostics,
        })
    }

    pub fn extract_page_texts(&self, bytes: &[u8]) -> Result<Vec<String>, ParseError> {
        let pdf = LoPdfDocument::load_mem(bytes)
            .map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let mut page_numbers: Vec<u32> = pdf.get_pages().into_keys().collect();
        page_numbers.sort_unstable();

        let mut pages = Vec::with_capacity(page_numbers.len());
        for page_number in page_numbers {
            let text = pdf.extract_text(&[page_number]).unwrap_or_default();
            pages.push(normalize_extracted_text(&text));
        }

        Ok(pages)
    }
}

impl FormatInspector for PdfInspector {
    fn format(&self) -> DocumentFormat {
        DocumentFormat::Pdf
    }

    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError> {
        let pdf = LoPdfDocument::load_mem(bytes)
            .map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let version = format!("PDF {}", pdf.version.trim());
        let page_count = pdf.get_pages().len();

        Ok(DocumentDiagnostics {
            format: DocumentFormat::Pdf,
            entry_count: page_count,
            section_count: page_count,
            asset_count: 0,
            is_encrypted: pdf.is_encrypted(),
            version_hint: Some(version),
            notes: vec![format!("Detected {page_count} PDF pages.")],
        })
    }
}

fn extract_info_string(pdf: &LoPdfDocument, key: &[u8]) -> Option<String> {
    let info_ref = pdf.trailer.get(b"Info").ok()?.as_reference().ok()?;
    let info_dict = pdf.get_dictionary(info_ref).ok()?;
    let value = info_dict.get(key).ok()?;
    match value {
        Object::String(bytes, _) => Some(String::from_utf8_lossy(bytes).trim().to_string())
            .filter(|value| !value.is_empty()),
        Object::Name(bytes) => Some(String::from_utf8_lossy(bytes).trim().to_string())
            .filter(|value| !value.is_empty()),
        _ => None,
    }
}

fn normalize_extracted_text(text: &str) -> String {
    text.replace('\0', "").replace('\r', "\n")
}

#[cfg(test)]
mod tests {
    use super::*;

    fn minimal_pdf() -> Vec<u8> {
        let mut pdf = b"%PDF-1.4\n".to_vec();
        let objects = [
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
            "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] >>\nendobj\n",
        ];

        let mut offsets = vec![0usize];
        for object in objects {
            offsets.push(pdf.len());
            pdf.extend_from_slice(object.as_bytes());
        }

        let xref_start = pdf.len();
        pdf.extend_from_slice(format!("xref\n0 {}\n", offsets.len()).as_bytes());
        pdf.extend_from_slice(b"0000000000 65535 f \n");
        for offset in offsets.iter().skip(1) {
            pdf.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
        }
        pdf.extend_from_slice(
            format!(
                "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{}\n%%EOF\n",
                offsets.len(),
                xref_start,
            )
            .as_bytes(),
        );
        pdf
    }

    #[test]
    fn parses_minimal_pdf_document() {
        let inspector = PdfInspector;
        let parsed = inspector
            .parse_bytes(&minimal_pdf(), Some("fixture.pdf"))
            .expect("minimal pdf should parse");

        assert_eq!(parsed.document.format, Some(DocumentFormat::Pdf));
        assert_eq!(parsed.document.metadata.page_count, Some(1));
        assert_eq!(parsed.diagnostics.format, DocumentFormat::Pdf);
        assert_eq!(parsed.diagnostics.section_count, 1);
    }
}

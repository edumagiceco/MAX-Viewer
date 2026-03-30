use max_viewer_core::{
    Block, Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata, FormatInspector,
    FormatSupport, PageLayout, Paragraph, ParagraphStyle, ParseError, Section, TextRun, TextStyle,
};

const A4_WIDTH: i32 = 59_528;
const A4_HEIGHT: i32 = 84_188;
const A4_MARGIN_LEFT_RIGHT: i32 = 8_504;
const A4_MARGIN_TOP: i32 = 5_668;
const A4_MARGIN_BOTTOM: i32 = 4_252;
const A4_MARGIN_HEADER_FOOTER: i32 = 4_252;
const BODY_FONT_SIZE: i32 = 1_050;

#[derive(Debug, Default)]
pub struct TextInspector;

#[derive(Debug, Clone)]
pub struct TextParseResult {
    pub document: Document,
    pub diagnostics: DocumentDiagnostics,
}

impl TextInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Text,
            status: "active".to_string(),
            implemented: vec![
                "UTF-8 plain text parsing".to_string(),
                "line-preserving document preview".to_string(),
                "page-ready paragraph mapping for editor round-trips".to_string(),
            ],
            planned: vec![
                "encoding detection for non-UTF-8 text files".to_string(),
                "tab width and wrap preferences".to_string(),
            ],
        }
    }

    pub fn parse_bytes(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
    ) -> Result<TextParseResult, ParseError> {
        let diagnostics = self.inspect_bytes(bytes)?;
        let text = std::str::from_utf8(bytes)
            .map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
        let blocks = normalized
            .split('\n')
            .map(|line| Block::Paragraph(paragraph_for_line(line)))
            .collect();

        Ok(TextParseResult {
            document: Document {
                format: Some(DocumentFormat::Text),
                metadata: DocumentMetadata {
                    title: fallback_title.map(ToOwned::to_owned),
                    language: Some("text/plain".to_string()),
                    ..DocumentMetadata::default()
                },
                sections: vec![Section {
                    id: 0,
                    blocks,
                    page_layout: Some(default_page_layout()),
                    headers: Vec::new(),
                    footers: Vec::new(),
                    page_start_number: None,
                }],
                assets: Vec::new(),
            },
            diagnostics,
        })
    }
}

impl FormatInspector for TextInspector {
    fn format(&self) -> DocumentFormat {
        DocumentFormat::Text
    }

    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError> {
        let text = std::str::from_utf8(bytes)
            .map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let normalized = text.replace("\r\n", "\n").replace('\r', "\n");
        let line_count = normalized.split('\n').count();

        Ok(DocumentDiagnostics {
            format: DocumentFormat::Text,
            entry_count: line_count,
            section_count: 1,
            asset_count: 0,
            is_encrypted: false,
            version_hint: Some("UTF-8".to_string()),
            notes: vec![
                "Plain text content was read directly without markup parsing.".to_string(),
                format!("Detected {line_count} source lines."),
            ],
        })
    }
}

fn default_page_layout() -> PageLayout {
    PageLayout {
        width: Some(A4_WIDTH),
        height: Some(A4_HEIGHT),
        landscape: false,
        margin_left: Some(A4_MARGIN_LEFT_RIGHT),
        margin_right: Some(A4_MARGIN_LEFT_RIGHT),
        margin_top: Some(A4_MARGIN_TOP),
        margin_bottom: Some(A4_MARGIN_BOTTOM),
        margin_header: Some(A4_MARGIN_HEADER_FOOTER),
        margin_footer: Some(A4_MARGIN_HEADER_FOOTER),
        margin_gutter: Some(0),
        page_border: None,
    }
}

fn paragraph_for_line(line: &str) -> Paragraph {
    Paragraph {
        marker: None,
        runs: vec![TextRun {
            text: line.to_string(),
            style: Some(base_text_style()),
        }],
        style: Some(default_paragraph_style()),
        line_segment_count: Some(1),
        layout_height_hint: None,
        page_break_before: false,
    }
}

fn default_paragraph_style() -> ParagraphStyle {
    ParagraphStyle {
        align: Some("LEFT".to_string()),
        indent: None,
        margin_left: Some(0),
        margin_right: Some(0),
        margin_prev: Some(0),
        margin_next: Some(0),
        line_spacing_type: Some("PERCENT".to_string()),
        line_spacing: Some(145),
        heading_type: None,
        heading_id_ref: None,
        heading_level: None,
        marker_align: None,
        marker_width_adjust: None,
        marker_text_offset_type: None,
        marker_text_offset: None,
        keep_with_next: false,
        keep_lines: false,
    }
}

fn base_text_style() -> TextStyle {
    TextStyle {
        font_family: Some("IBM Plex Mono".to_string()),
        font_size: Some(BODY_FONT_SIZE),
        text_color: Some("#111827".to_string()),
        background_color: None,
        underline_color: None,
        width_ratio: None,
        letter_spacing: None,
        relative_size: None,
        baseline_offset: None,
        use_font_space: true,
        use_kerning: false,
        bold: false,
        italic: false,
        underline: false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_plain_text_with_blank_lines() {
        let inspector = TextInspector;
        let parsed = inspector
            .parse_bytes(b"alpha\r\n\r\nbeta\n", Some("notes.txt"))
            .expect("text should parse");

        assert_eq!(parsed.document.format, Some(DocumentFormat::Text));
        assert_eq!(parsed.diagnostics.format, DocumentFormat::Text);
        assert_eq!(parsed.document.sections.len(), 1);
        assert_eq!(parsed.document.sections[0].blocks.len(), 4);

        let blocks = &parsed.document.sections[0].blocks;
        match &blocks[0] {
            Block::Paragraph(paragraph) => assert_eq!(paragraph.runs[0].text, "alpha"),
            _ => panic!("expected paragraph"),
        }
        match &blocks[1] {
            Block::Paragraph(paragraph) => assert_eq!(paragraph.runs[0].text, ""),
            _ => panic!("expected paragraph"),
        }
    }

    #[test]
    fn rejects_non_utf8_text() {
        let inspector = TextInspector;
        let error = inspector
            .parse_bytes(&[0xff, 0xfe, 0xfd], Some("broken.txt"))
            .expect_err("invalid utf-8 should fail");

        assert!(matches!(error, ParseError::InvalidData(_)));
    }
}

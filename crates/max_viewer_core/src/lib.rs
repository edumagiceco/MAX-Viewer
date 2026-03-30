use serde::{Deserialize, Serialize};
use thiserror::Error;

pub const APP_NAME: &str = "MAX Viewer";

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "lowercase")]
pub enum DocumentFormat {
    Hwp,
    Hwpx,
    Pdf,
    Markdown,
    Text,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct DocumentMetadata {
    pub title: Option<String>,
    pub author: Option<String>,
    pub page_count: Option<u32>,
    pub language: Option<String>,
    pub created_at: Option<String>,
    pub modified_at: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextStyle {
    pub font_family: Option<String>,
    pub font_size: Option<i32>,
    pub text_color: Option<String>,
    pub background_color: Option<String>,
    pub underline_color: Option<String>,
    pub width_ratio: Option<i32>,
    pub letter_spacing: Option<i32>,
    pub relative_size: Option<i32>,
    pub baseline_offset: Option<i32>,
    pub use_font_space: bool,
    pub use_kerning: bool,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TextRun {
    pub text: String,
    pub style: Option<TextStyle>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphStyle {
    pub align: Option<String>,
    pub indent: Option<i32>,
    pub margin_left: Option<i32>,
    pub margin_right: Option<i32>,
    pub margin_prev: Option<i32>,
    pub margin_next: Option<i32>,
    pub line_spacing_type: Option<String>,
    pub line_spacing: Option<i32>,
    pub heading_type: Option<String>,
    pub heading_id_ref: Option<u32>,
    pub heading_level: Option<u32>,
    pub marker_align: Option<String>,
    pub marker_width_adjust: Option<i32>,
    pub marker_text_offset_type: Option<String>,
    pub marker_text_offset: Option<i32>,
    pub keep_with_next: bool,
    pub keep_lines: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Paragraph {
    pub marker: Option<TextRun>,
    pub runs: Vec<TextRun>,
    pub style: Option<ParagraphStyle>,
    pub line_segment_count: Option<u32>,
    pub layout_height_hint: Option<i32>,
    pub page_break_before: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableBorder {
    pub style: Option<String>,
    pub width: Option<String>,
    pub color: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableDiagonal {
    pub style: Option<String>,
    pub width: Option<String>,
    pub color: Option<String>,
    pub slash_type: Option<String>,
    pub back_slash_type: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableCellStyle {
    pub background_color: Option<String>,
    pub background_image: Option<String>,
    pub border_left: Option<TableBorder>,
    pub border_right: Option<TableBorder>,
    pub border_top: Option<TableBorder>,
    pub border_bottom: Option<TableBorder>,
    pub diagonal: Option<TableDiagonal>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableCell {
    pub text: String,
    pub blocks: Vec<Block>,
    pub col_span: Option<u32>,
    pub row_span: Option<u32>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub padding_left: Option<i32>,
    pub padding_right: Option<i32>,
    pub padding_top: Option<i32>,
    pub padding_bottom: Option<i32>,
    pub style: Option<TableCellStyle>,
    pub is_header: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableRow {
    pub cells: Vec<TableCell>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct TableBlock {
    pub rows: Vec<TableRow>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub cell_spacing: Option<i32>,
    pub style: Option<TableCellStyle>,
    pub no_adjust: bool,
    pub repeat_header: bool,
    pub header_row_count: Option<u32>,
    pub page_break_before: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct ImageBlock {
    pub kind: String,
    pub asset_id: String,
    pub alt_text: Option<String>,
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub width_rel_to: Option<String>,
    pub height_rel_to: Option<String>,
    pub treat_as_char: bool,
    pub text_wrap: Option<String>,
    pub z_order: Option<i32>,
    pub vert_rel_to: Option<String>,
    pub horz_rel_to: Option<String>,
    pub vert_align: Option<String>,
    pub horz_align: Option<String>,
    pub vert_offset: Option<i32>,
    pub horz_offset: Option<i32>,
    pub distance_left: Option<i32>,
    pub distance_right: Option<i32>,
    pub distance_top: Option<i32>,
    pub distance_bottom: Option<i32>,
    pub rotation: Option<i32>,
    pub caption: Option<String>,
    pub page_break_before: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct FootnoteBlock {
    pub kind: String,
    pub number: Option<u32>,
    pub blocks: Vec<Block>,
    pub page_break_before: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct UnsupportedBlock {
    pub kind: String,
    pub reason: Option<String>,
    pub page_break_before: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", tag = "type", content = "value")]
pub enum Block {
    Paragraph(Paragraph),
    Table(TableBlock),
    Image(ImageBlock),
    Footnote(FootnoteBlock),
    Unsupported(UnsupportedBlock),
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct PageLayout {
    pub width: Option<i32>,
    pub height: Option<i32>,
    pub landscape: bool,
    pub margin_left: Option<i32>,
    pub margin_right: Option<i32>,
    pub margin_top: Option<i32>,
    pub margin_bottom: Option<i32>,
    pub margin_header: Option<i32>,
    pub margin_footer: Option<i32>,
    pub margin_gutter: Option<i32>,
    pub page_border: Option<TableCellStyle>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct HeaderFooter {
    pub apply_page_type: Option<String>,
    pub blocks: Vec<Block>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Section {
    pub id: usize,
    pub blocks: Vec<Block>,
    pub page_layout: Option<PageLayout>,
    pub headers: Vec<HeaderFooter>,
    pub footers: Vec<HeaderFooter>,
    pub page_start_number: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct AssetRef {
    pub id: String,
    pub media_type: String,
    pub source_path: Option<String>,
    pub data_uri: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct Document {
    pub format: Option<DocumentFormat>,
    pub metadata: DocumentMetadata,
    pub sections: Vec<Section>,
    pub assets: Vec<AssetRef>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentDiagnostics {
    pub format: DocumentFormat,
    pub entry_count: usize,
    pub section_count: usize,
    pub asset_count: usize,
    pub is_encrypted: bool,
    pub version_hint: Option<String>,
    pub notes: Vec<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormatSupport {
    pub format: DocumentFormat,
    pub status: String,
    pub implemented: Vec<String>,
    pub planned: Vec<String>,
}

pub trait FormatInspector {
    fn format(&self) -> DocumentFormat;
    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError>;
}

#[derive(Debug, Error)]
pub enum ParseError {
    #[error("unsupported format: {0}")]
    UnsupportedFormat(String),
    #[error("invalid container: {0}")]
    InvalidContainer(String),
    #[error("invalid data: {0}")]
    InvalidData(String),
}

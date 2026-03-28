use std::io::{Cursor, Read, Seek, SeekFrom};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64};
use cfb::CompoundFile;
use flate2::read::{DeflateDecoder, ZlibDecoder};
use max_viewer_core::{
    AssetRef, Block, Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata,
    FormatInspector, FormatSupport, ImageBlock, PageLayout, Paragraph, ParagraphStyle, ParseError,
    Section, TableBlock, TableBorder, TableCell, TableCellStyle, TableRow, TextRun, TextStyle,
    UnsupportedBlock,
};

pub const CFB_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const FILE_HEADER_STREAM: &str = "/FileHeader";
const DOC_INFO_STREAM: &str = "/DocInfo";
const PREVIEW_TEXT_STREAM: &str = "/PrvText";
const BODY_TEXT_STORAGE: &str = "/BodyText";

const HWPTAG_BEGIN: u16 = 0x010;
const HWPTAG_BIN_DATA: u16 = HWPTAG_BEGIN + 2;
const HWPTAG_FACE_NAME: u16 = HWPTAG_BEGIN + 3;
const HWPTAG_BORDER_FILL: u16 = HWPTAG_BEGIN + 4;
const HWPTAG_CHAR_SHAPE: u16 = HWPTAG_BEGIN + 5;
const HWPTAG_PARA_SHAPE: u16 = HWPTAG_BEGIN + 9;
const HWPTAG_STYLE: u16 = HWPTAG_BEGIN + 10;
const HWPTAG_PARA_HEADER: u16 = HWPTAG_BEGIN + 50;
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51;
const HWPTAG_PARA_CHAR_SHAPE: u16 = HWPTAG_BEGIN + 52;
const HWPTAG_CTRL_HEADER: u16 = HWPTAG_BEGIN + 54;
const HWPTAG_LIST_HEADER: u16 = HWPTAG_BEGIN + 55;
const HWPTAG_PAGE_DEF: u16 = HWPTAG_BEGIN + 56;
const HWPTAG_TABLE: u16 = HWPTAG_BEGIN + 59;
const HWPTAG_SHAPE_COMPONENT: u16 = HWPTAG_BEGIN + 60;

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
pub struct HwpParseResult {
    pub document: Document,
    pub diagnostics: DocumentDiagnostics,
}

#[derive(Debug, Clone, Default)]
struct HwpHeader {
    version_hint: Option<String>,
    attributes: u32,
}

#[derive(Debug, Default)]
pub struct HwpInspector;

// ---------------------------------------------------------------------------
// Record iterator
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
struct Record<'a> {
    tag_id: u16,
    level: u16,
    data: &'a [u8],
}

fn iter_records(bytes: &[u8]) -> Vec<Record<'_>> {
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= bytes.len() {
        let header = u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap_or([0; 4]));
        offset += 4;
        let tag_id = (header & 0x3ff) as u16;
        let level = ((header >> 10) & 0x3ff) as u16;
        let size_bits = ((header >> 20) & 0xfff) as usize;
        let size = if size_bits == 0x0fff {
            if offset + 4 > bytes.len() {
                break;
            }
            let ext =
                u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap_or([0; 4])) as usize;
            offset += 4;
            ext
        } else {
            size_bits
        };
        if offset + size > bytes.len() {
            break;
        }
        records.push(Record {
            tag_id,
            level,
            data: &bytes[offset..offset + size],
        });
        offset += size;
    }
    records
}

// ---------------------------------------------------------------------------
// DocInfo structures
// ---------------------------------------------------------------------------

#[derive(Debug, Clone, Default)]
struct FaceName {
    name: String,
}

#[derive(Debug, Clone, Default)]
struct CharShape {
    face_name_ids: [u16; 7],
    ratios: [u8; 7],
    spacings: [i8; 7],
    height: u32,
    attributes: u32,
    text_color: u32,
}

#[derive(Debug, Clone, Default)]
struct ParaShape {
    attributes: u32,
    margin_left: i32,
    margin_right: i32,
    indent: i32,
    margin_prev: i32,
    margin_next: i32,
    line_spacing_type: u32,
    line_spacing: i32,
}

#[derive(Debug, Clone, Default)]
struct HwpStyle {
    #[allow(dead_code)]
    name: String,
    para_shape_id: u16,
    char_shape_id: u16,
}

#[derive(Debug, Clone, Default)]
struct BinDataRef {
    storage_type: u16,
    extension: String,
}

#[derive(Debug, Clone, Default)]
struct HwpBorderFill {
    border_left: Option<TableBorder>,
    border_right: Option<TableBorder>,
    border_top: Option<TableBorder>,
    border_bottom: Option<TableBorder>,
    background_color: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct DocInfoStore {
    face_names: Vec<FaceName>,
    char_shapes: Vec<CharShape>,
    para_shapes: Vec<ParaShape>,
    styles: Vec<HwpStyle>,
    border_fills: Vec<HwpBorderFill>,
    bin_data_refs: Vec<BinDataRef>,
}

// ---------------------------------------------------------------------------
// DocInfo parsing (steps 1-3)
// ---------------------------------------------------------------------------

fn read_doc_info(compound_file: &mut CompoundFile<Cursor<&[u8]>>) -> DocInfoStore {
    if !compound_file.exists(DOC_INFO_STREAM) {
        return DocInfoStore::default();
    }
    let bytes = match read_stream_bytes(compound_file, DOC_INFO_STREAM) {
        Ok(b) => b,
        Err(_) => return DocInfoStore::default(),
    };
    let decoded = try_decode_stream(&bytes);
    let records = iter_records(&decoded);
    let mut store = DocInfoStore::default();

    for rec in &records {
        match rec.tag_id {
            HWPTAG_FACE_NAME => {
                if let Some(name) = parse_face_name(rec.data) {
                    store.face_names.push(name);
                }
            }
            HWPTAG_CHAR_SHAPE => {
                store.char_shapes.push(parse_char_shape(rec.data));
            }
            HWPTAG_PARA_SHAPE => {
                store.para_shapes.push(parse_para_shape(rec.data));
            }
            HWPTAG_STYLE => {
                if let Some(style) = parse_style(rec.data) {
                    store.styles.push(style);
                }
            }
            HWPTAG_BORDER_FILL => {
                store.border_fills.push(parse_border_fill(rec.data));
            }
            HWPTAG_BIN_DATA => {
                store.bin_data_refs.push(parse_bin_data_ref(rec.data));
            }
            _ => {}
        }
    }

    store
}

fn parse_face_name(data: &[u8]) -> Option<FaceName> {
    if data.len() < 3 {
        return None;
    }
    let name_len = u16::from_le_bytes(data[1..3].try_into().ok()?) as usize;
    let name_bytes_start = 3;
    let name_bytes_end = name_bytes_start + name_len * 2;
    if data.len() < name_bytes_end {
        return None;
    }
    let units: Vec<u16> = data[name_bytes_start..name_bytes_end]
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    Some(FaceName {
        name: String::from_utf16_lossy(&units),
    })
}

fn parse_char_shape(data: &[u8]) -> CharShape {
    let mut cs = CharShape::default();
    if data.len() < 72 {
        return cs;
    }
    for i in 0..7 {
        cs.face_name_ids[i] =
            u16::from_le_bytes(data[i * 2..i * 2 + 2].try_into().unwrap_or([0; 2]));
    }
    for i in 0..7 {
        cs.ratios[i] = data[14 + i];
    }
    for i in 0..7 {
        cs.spacings[i] = data[21 + i] as i8;
    }
    // rel_sizes at 28..35, offsets at 35..42 — skip for now
    cs.height = u32::from_le_bytes(data[42..46].try_into().unwrap_or([0; 4]));
    cs.attributes = u32::from_le_bytes(data[46..50].try_into().unwrap_or([0; 4]));
    // shadow_gap1(1) shadow_gap2(1) -> 50..52
    cs.text_color = u32::from_le_bytes(data[52..56].try_into().unwrap_or([0; 4]));
    cs
}

fn parse_para_shape(data: &[u8]) -> ParaShape {
    let mut ps = ParaShape::default();
    if data.len() < 30 {
        return ps;
    }
    ps.attributes = u32::from_le_bytes(data[0..4].try_into().unwrap_or([0; 4]));
    ps.margin_left = i32::from_le_bytes(data[4..8].try_into().unwrap_or([0; 4]));
    ps.margin_right = i32::from_le_bytes(data[8..12].try_into().unwrap_or([0; 4]));
    ps.indent = i32::from_le_bytes(data[12..16].try_into().unwrap_or([0; 4]));
    ps.margin_prev = i32::from_le_bytes(data[16..20].try_into().unwrap_or([0; 4]));
    ps.margin_next = i32::from_le_bytes(data[20..24].try_into().unwrap_or([0; 4]));
    ps.line_spacing_type = u32::from_le_bytes(data[24..28].try_into().unwrap_or([0; 4])) & 0x1f;
    ps.line_spacing = i32::from_le_bytes(data[28..32].try_into().unwrap_or([0; 4]));
    ps
}

fn parse_style(data: &[u8]) -> Option<HwpStyle> {
    if data.len() < 6 {
        return None;
    }
    let name_len = u16::from_le_bytes(data[0..2].try_into().ok()?) as usize;
    let after_name = 2 + name_len * 2;
    if data.len() < after_name + 2 {
        return None;
    }
    let name_units: Vec<u16> = data[2..after_name]
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let name = String::from_utf16_lossy(&name_units);
    // english name length
    let en_name_len =
        u16::from_le_bytes(data[after_name..after_name + 2].try_into().ok()?) as usize;
    let after_en = after_name + 2 + en_name_len * 2;
    // 1 byte type, 1 byte next_style_id
    let props_start = after_en + 2;
    if data.len() < props_start + 4 {
        return Some(HwpStyle {
            name,
            ..Default::default()
        });
    }
    let para_shape_id = u16::from_le_bytes(
        data[props_start..props_start + 2]
            .try_into()
            .unwrap_or([0; 2]),
    );
    let char_shape_id = u16::from_le_bytes(
        data[props_start + 2..props_start + 4]
            .try_into()
            .unwrap_or([0; 2]),
    );
    Some(HwpStyle {
        name,
        para_shape_id,
        char_shape_id,
    })
}

fn parse_border_fill(data: &[u8]) -> HwpBorderFill {
    let mut bf = HwpBorderFill::default();
    // HWP BORDER_FILL record: attributes(2) + border sides(each: type(1) + width(1) + color(4)) × 4 + diagonal...
    // Simplified: read 4 border sides starting at offset 2
    if data.len() < 26 {
        return bf;
    }
    let read_border = |offset: usize| -> Option<TableBorder> {
        if offset + 6 > data.len() {
            return None;
        }
        let border_type = data[offset];
        let border_width = data[offset + 1];
        let color = u32::from_le_bytes(data[offset + 2..offset + 6].try_into().unwrap_or([0; 4]));
        if border_type == 0 && border_width == 0 {
            return None;
        }
        let style_name = match border_type {
            0 => "NONE",
            1 => "SOLID",
            2 => "DASHED",
            3 => "DOTTED",
            4 => "DASH_DOT",
            _ => "SOLID",
        };
        let width_mm = format!("{:.2} mm", border_width as f64 * 0.1);
        Some(TableBorder {
            style: Some(style_name.to_string()),
            width: Some(width_mm),
            color: Some(colorref_to_hex(color)),
        })
    };
    bf.border_left = read_border(2);
    bf.border_right = read_border(8);
    bf.border_top = read_border(14);
    bf.border_bottom = read_border(20);
    // Fill info comes after borders — try to read background color
    // The fill section is variable-length; look for a COLORREF at the tail
    if data.len() >= 30 {
        let fill_offset = 26;
        if fill_offset + 4 <= data.len() {
            let fill_color = u32::from_le_bytes(
                data[fill_offset..fill_offset + 4].try_into().unwrap_or([0; 4]),
            );
            if fill_color != 0 && fill_color != 0x00FFFFFF {
                bf.background_color = Some(colorref_to_hex(fill_color));
            }
        }
    }
    bf
}

fn border_fill_to_cell_style(bf: &HwpBorderFill) -> TableCellStyle {
    TableCellStyle {
        background_color: bf.background_color.clone(),
        border_left: bf.border_left.clone(),
        border_right: bf.border_right.clone(),
        border_top: bf.border_top.clone(),
        border_bottom: bf.border_bottom.clone(),
        diagonal: None,
    }
}

fn parse_bin_data_ref(data: &[u8]) -> BinDataRef {
    let mut bdr = BinDataRef::default();
    if data.len() < 2 {
        return bdr;
    }
    bdr.storage_type = u16::from_le_bytes(data[0..2].try_into().unwrap_or([0; 2]));
    // For EMBEDDING (type 0), after type there's abs_path(utf16 len-prefixed), rel_path, then bin_data_id, extension
    // For simplicity, try to find the extension at the tail of the record
    if bdr.storage_type == 0 && data.len() >= 4 {
        let ext_len = u16::from_le_bytes(
            data[data.len().saturating_sub(2 + 10)..data.len().saturating_sub(10)]
                .try_into()
                .unwrap_or([0; 2]),
        ) as usize;
        if ext_len > 0 && ext_len <= 10 {
            let ext_start = data.len().saturating_sub(ext_len * 2);
            if ext_start + ext_len * 2 <= data.len() {
                let units: Vec<u16> = data[ext_start..ext_start + ext_len * 2]
                    .chunks_exact(2)
                    .map(|c| u16::from_le_bytes([c[0], c[1]]))
                    .collect();
                bdr.extension = String::from_utf16_lossy(&units).to_lowercase();
            }
        }
    }
    bdr
}

// ---------------------------------------------------------------------------
// BinData asset loading (step 3)
// ---------------------------------------------------------------------------

fn load_bin_data_assets(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
    bin_data_refs: &[BinDataRef],
) -> Vec<AssetRef> {
    let entry_paths = collect_entry_paths(compound_file);
    let mut bin_paths: Vec<String> = entry_paths
        .into_iter()
        .filter(|p| p.starts_with("/BinData/"))
        .collect();
    bin_paths.sort();

    let mut assets = Vec::new();
    for (index, path) in bin_paths.iter().enumerate() {
        let bytes = match read_stream_bytes(compound_file, path) {
            Ok(b) => b,
            Err(_) => continue,
        };
        if bytes.is_empty() {
            continue;
        }

        let ext = bin_data_refs
            .get(index)
            .map(|b| b.extension.as_str())
            .unwrap_or("");
        let media_type = guess_media_type(ext, &bytes);
        let data_uri = format!("data:{};base64,{}", media_type, BASE64.encode(&bytes));
        let id = path
            .rsplit('/')
            .next()
            .unwrap_or(path)
            .to_string();

        assets.push(AssetRef {
            id,
            media_type,
            source_path: Some(path.clone()),
            data_uri: Some(data_uri),
        });
    }

    assets
}

fn guess_media_type(ext: &str, bytes: &[u8]) -> String {
    match ext.trim_start_matches('.') {
        "png" => return "image/png".to_string(),
        "jpg" | "jpeg" => return "image/jpeg".to_string(),
        "gif" => return "image/gif".to_string(),
        "bmp" => return "image/bmp".to_string(),
        "wmf" => return "image/x-wmf".to_string(),
        "emf" => return "image/x-emf".to_string(),
        _ => {}
    }
    if bytes.starts_with(&[0x89, 0x50, 0x4E, 0x47]) {
        "image/png".to_string()
    } else if bytes.starts_with(&[0xFF, 0xD8, 0xFF]) {
        "image/jpeg".to_string()
    } else if bytes.starts_with(b"GIF") {
        "image/gif".to_string()
    } else if bytes.starts_with(b"BM") {
        "image/bmp".to_string()
    } else {
        "application/octet-stream".to_string()
    }
}

// ---------------------------------------------------------------------------
// DocInfo → core model converters (step 4)
// ---------------------------------------------------------------------------

fn char_shape_to_text_style(cs: &CharShape, face_names: &[FaceName]) -> TextStyle {
    TextStyle {
        font_family: face_names.get(cs.face_name_ids[0] as usize).map(|f| f.name.clone()),
        font_size: if cs.height > 0 {
            Some(cs.height as i32 / 100)
        } else {
            None
        },
        bold: cs.attributes & (1 << 1) != 0,
        italic: cs.attributes & (1 << 0) != 0,
        underline: (cs.attributes >> 2) & 0x3 != 0,
        text_color: if cs.text_color != 0 {
            Some(colorref_to_hex(cs.text_color))
        } else {
            None
        },
        width_ratio: if cs.ratios[0] != 0 && cs.ratios[0] != 100 {
            Some(cs.ratios[0] as i32)
        } else {
            None
        },
        letter_spacing: if cs.spacings[0] != 0 {
            Some(cs.spacings[0] as i32)
        } else {
            None
        },
        ..TextStyle::default()
    }
}

fn para_shape_to_paragraph_style(ps: &ParaShape) -> ParagraphStyle {
    let align_bits = ps.attributes & 0x7;
    let align = match align_bits {
        0 => Some("JUSTIFY".to_string()),
        1 => Some("LEFT".to_string()),
        2 => Some("RIGHT".to_string()),
        3 => Some("CENTER".to_string()),
        4 => Some("DISTRIBUTE".to_string()),
        _ => None,
    };
    ParagraphStyle {
        align,
        indent: if ps.indent != 0 {
            Some(ps.indent)
        } else {
            None
        },
        margin_left: if ps.margin_left != 0 {
            Some(ps.margin_left)
        } else {
            None
        },
        margin_right: if ps.margin_right != 0 {
            Some(ps.margin_right)
        } else {
            None
        },
        margin_prev: if ps.margin_prev != 0 {
            Some(ps.margin_prev)
        } else {
            None
        },
        margin_next: if ps.margin_next != 0 {
            Some(ps.margin_next)
        } else {
            None
        },
        line_spacing_type: Some(format!("{}", ps.line_spacing_type)),
        line_spacing: if ps.line_spacing != 0 {
            Some(ps.line_spacing)
        } else {
            None
        },
        ..ParagraphStyle::default()
    }
}

fn colorref_to_hex(c: u32) -> String {
    let r = c & 0xff;
    let g = (c >> 8) & 0xff;
    let b = (c >> 16) & 0xff;
    format!("#{:02x}{:02x}{:02x}", r, g, b)
}

// ---------------------------------------------------------------------------
// BodyText parsing with DocInfo (steps 4-7)
// ---------------------------------------------------------------------------

fn parse_body_section_with_docinfo(
    bytes: &[u8],
    doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> Vec<Block> {
    if bytes.is_empty() {
        return Vec::new();
    }
    let decoded = try_decode_stream(bytes);
    let records = iter_records(&decoded);
    if records.is_empty() {
        let fallback_records = iter_records(bytes);
        if fallback_records.is_empty() {
            return Vec::new();
        }
        return parse_records_to_blocks(&fallback_records, doc_info, assets);
    }

    parse_records_to_blocks(&records, doc_info, assets)
}

fn parse_records_to_blocks(
    records: &[Record<'_>],
    doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> Vec<Block> {
    let mut blocks = Vec::new();
    let mut page_layout: Option<PageLayout> = None;
    let mut header_blocks: Vec<Block> = Vec::new();
    let mut footer_blocks: Vec<Block> = Vec::new();

    let mut idx = 0;
    while idx < records.len() {
        let rec = &records[idx];
        match rec.tag_id {
            HWPTAG_PARA_HEADER => {
                let (para, consumed) =
                    parse_paragraph_with_style(&records[idx..], doc_info, assets);
                if let Some(para) = para {
                    blocks.push(Block::Paragraph(para));
                }
                idx += consumed.max(1);
            }
            HWPTAG_CTRL_HEADER => {
                let ctrl_type = read_ctrl_type(rec.data);
                match ctrl_type.as_deref() {
                    Some("tbl ") => {
                        let (table, consumed) =
                            parse_table_control(&records[idx..], doc_info, assets);
                        if let Some(table) = table {
                            blocks.push(Block::Table(table));
                        }
                        idx += consumed.max(1);
                    }
                    Some("gso ") => {
                        let (image, consumed) =
                            parse_gso_control(&records[idx..], doc_info, assets);
                        if let Some(img) = image {
                            blocks.push(Block::Image(img));
                        } else {
                            blocks.push(Block::Unsupported(UnsupportedBlock {
                                kind: "hwp-gso".to_string(),
                                reason: Some("drawing object".to_string()),
                            }));
                        }
                        idx += consumed.max(1);
                    }
                    Some("secd") => {
                        if let Some(pd) = find_page_def(&records[idx + 1..]) {
                            page_layout = Some(pd);
                        }
                        idx += 1;
                    }
                    Some("daeh") => {
                        // Header control
                        let (hf_blocks, consumed) = parse_header_footer_control(&records[idx..], doc_info, assets);
                        header_blocks = hf_blocks;
                        idx += consumed.max(1);
                    }
                    Some("toof") => {
                        // Footer control
                        let (hf_blocks, consumed) = parse_header_footer_control(&records[idx..], doc_info, assets);
                        footer_blocks = hf_blocks;
                        idx += consumed.max(1);
                    }
                    _ => {
                        idx += 1;
                    }
                }
            }
            HWPTAG_PAGE_DEF => {
                page_layout = parse_page_def(rec.data);
                idx += 1;
            }
            _ => {
                idx += 1;
            }
        }
    }

    // Encode header/footer as sentinel blocks
    if !header_blocks.is_empty() {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "__header_blocks__".to_string(),
            reason: Some(format!("{}", header_blocks.len())),
        }));
        blocks.extend(header_blocks.into_iter().map(|b| {
            // Wrap in a sentinel-tagged unsupported to mark as header content
            Block::Unsupported(UnsupportedBlock {
                kind: "__header_content__".to_string(),
                reason: match &b {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
                    _ => Some("block".to_string()),
                },
            })
        }));
    }
    if !footer_blocks.is_empty() {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "__footer_blocks__".to_string(),
            reason: Some(format!("{}", footer_blocks.len())),
        }));
        blocks.extend(footer_blocks.into_iter().map(|b| {
            Block::Unsupported(UnsupportedBlock {
                kind: "__footer_content__".to_string(),
                reason: match &b {
                    Block::Paragraph(p) => Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>()),
                    _ => Some("block".to_string()),
                },
            })
        }));
    }

    if let Some(layout) = page_layout {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "__page_layout__".to_string(),
            reason: Some(serde_page_layout(&layout)),
        }));
    }

    blocks
}

fn serde_page_layout(layout: &PageLayout) -> String {
    // Encode as simple key=value pairs
    format!(
        "w={},h={},l={},r={},t={},b={},hdr={},ftr={},g={},ls={}",
        layout.width.unwrap_or(0),
        layout.height.unwrap_or(0),
        layout.margin_left.unwrap_or(0),
        layout.margin_right.unwrap_or(0),
        layout.margin_top.unwrap_or(0),
        layout.margin_bottom.unwrap_or(0),
        layout.margin_header.unwrap_or(0),
        layout.margin_footer.unwrap_or(0),
        layout.margin_gutter.unwrap_or(0),
        if layout.landscape { 1 } else { 0 },
    )
}

fn deserialize_page_layout(s: &str) -> Option<PageLayout> {
    let mut layout = PageLayout::default();
    for pair in s.split(',') {
        let mut kv = pair.splitn(2, '=');
        let key = kv.next()?;
        let val: i32 = kv.next()?.parse().ok()?;
        match key {
            "w" => layout.width = Some(val),
            "h" => layout.height = Some(val),
            "l" => layout.margin_left = Some(val),
            "r" => layout.margin_right = Some(val),
            "t" => layout.margin_top = Some(val),
            "b" => layout.margin_bottom = Some(val),
            "hdr" => layout.margin_header = Some(val),
            "ftr" => layout.margin_footer = Some(val),
            "g" => layout.margin_gutter = Some(val),
            "ls" => layout.landscape = val != 0,
            _ => {}
        }
    }
    Some(layout)
}

// --- Paragraph with style (step 4) ---

fn parse_paragraph_with_style<'a>(
    records: &[Record<'a>],
    doc_info: &DocInfoStore,
    _assets: &[AssetRef],
) -> (Option<Paragraph>, usize) {
    let header = &records[0];
    if header.tag_id != HWPTAG_PARA_HEADER {
        return (None, 1);
    }
    let data = header.data;
    let char_count = if data.len() >= 4 {
        Some(u32::from_le_bytes(data[0..4].try_into().unwrap_or([0; 4])) & 0x7fff_ffff)
    } else {
        None
    };
    let para_shape_id = if data.len() >= 10 {
        Some(u16::from_le_bytes(data[8..10].try_into().unwrap_or([0; 2])))
    } else {
        None
    };
    let style_id = if data.len() >= 11 { Some(data[10]) } else { None };

    let mut text_data: Option<&[u8]> = None;
    let mut char_shape_positions: Vec<(u32, u32)> = Vec::new();
    let mut consumed = 1;

    let base_level = header.level;
    for rec in records.iter().skip(1) {
        let is_para_child = rec.tag_id == HWPTAG_PARA_TEXT
            || rec.tag_id == HWPTAG_PARA_CHAR_SHAPE
            || rec.level > base_level;
        if !is_para_child {
            break;
        }
        consumed += 1;
        match rec.tag_id {
            HWPTAG_PARA_TEXT => {
                text_data = Some(rec.data);
            }
            HWPTAG_PARA_CHAR_SHAPE => {
                char_shape_positions = parse_char_shape_positions(rec.data);
            }
            HWPTAG_CTRL_HEADER | HWPTAG_LIST_HEADER | HWPTAG_TABLE | HWPTAG_SHAPE_COMPONENT => {
                // These belong to child controls — stop paragraph parsing
                consumed -= 1;
                break;
            }
            _ => {}
        }
    }

    let text_data = match text_data {
        Some(d) if !d.is_empty() => d,
        _ => return (None, consumed),
    };

    let units = extract_text_units(text_data, char_count);
    if units.iter().all(|u| *u <= 0x001f) {
        return (None, consumed);
    }

    // Build runs with char shape splits
    let runs = build_styled_runs(&units, &char_shape_positions, doc_info);
    if runs.is_empty() {
        return (None, consumed);
    }

    // Resolve paragraph style
    let para_style = resolve_para_style(para_shape_id, style_id, doc_info);

    (
        Some(Paragraph {
            marker: None,
            runs,
            style: para_style,
            line_segment_count: None,
            layout_height_hint: None,
            page_break_before: false,
        }),
        consumed,
    )
}

fn extract_text_units(data: &[u8], char_count: Option<u32>) -> Vec<u16> {
    if data.len() % 2 != 0 {
        return Vec::new();
    }
    let mut units: Vec<u16> = data
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    if let Some(count) = char_count {
        let expected = count as usize;
        if expected > 0 && expected < units.len() {
            units.truncate(expected);
        }
    }
    units
}

fn parse_char_shape_positions(data: &[u8]) -> Vec<(u32, u32)> {
    let mut positions = Vec::new();
    let mut offset = 0;
    while offset + 8 <= data.len() {
        let pos = u32::from_le_bytes(data[offset..offset + 4].try_into().unwrap_or([0; 4]));
        let id = u32::from_le_bytes(data[offset + 4..offset + 8].try_into().unwrap_or([0; 4]));
        positions.push((pos, id));
        offset += 8;
    }
    positions
}

fn build_styled_runs(
    units: &[u16],
    char_shape_positions: &[(u32, u32)],
    doc_info: &DocInfoStore,
) -> Vec<TextRun> {
    if char_shape_positions.is_empty() {
        let text = units_to_text(units);
        if text.trim().is_empty() {
            return Vec::new();
        }
        let default_style = doc_info.char_shapes.first().map(|cs| char_shape_to_text_style(cs, &doc_info.face_names));
        return vec![TextRun {
            text,
            style: default_style,
        }];
    }

    let mut runs = Vec::new();
    for (i, &(pos, cs_id)) in char_shape_positions.iter().enumerate() {
        let start = pos as usize;
        let end = char_shape_positions
            .get(i + 1)
            .map(|(next_pos, _)| *next_pos as usize)
            .unwrap_or(units.len());
        if start >= units.len() {
            break;
        }
        let slice = &units[start..end.min(units.len())];
        let text = units_to_text(slice);
        if text.is_empty() {
            continue;
        }
        let style = doc_info
            .char_shapes
            .get(cs_id as usize)
            .map(|cs| char_shape_to_text_style(cs, &doc_info.face_names));
        runs.push(TextRun { text, style });
    }

    if runs.is_empty() {
        let text = units_to_text(units);
        if !text.trim().is_empty() {
            runs.push(TextRun { text, style: None });
        }
    }

    runs
}

fn units_to_text(units: &[u16]) -> String {
    let mut text = String::new();
    for &unit in units {
        match unit {
            0x0009 => text.push('\t'),
            0x000a => text.push('\n'),
            0x000d | 0x0000..=0x001f => {}
            _ => {
                if let Some(ch) = char::from_u32(unit as u32) {
                    text.push(ch);
                }
            }
        }
    }
    text
}

fn resolve_para_style(
    para_shape_id: Option<u16>,
    style_id: Option<u8>,
    doc_info: &DocInfoStore,
) -> Option<ParagraphStyle> {
    // Direct para_shape_id takes priority
    if let Some(ps_id) = para_shape_id {
        if let Some(ps) = doc_info.para_shapes.get(ps_id as usize) {
            return Some(para_shape_to_paragraph_style(ps));
        }
    }
    // Fallback to style
    if let Some(s_id) = style_id {
        if let Some(style) = doc_info.styles.get(s_id as usize) {
            if let Some(ps) = doc_info.para_shapes.get(style.para_shape_id as usize) {
                return Some(para_shape_to_paragraph_style(ps));
            }
        }
    }
    None
}

// --- Control parsing helpers ---

fn read_ctrl_type(data: &[u8]) -> Option<String> {
    if data.len() < 4 {
        return None;
    }
    // Control type is stored as little-endian 4 bytes representing reversed ASCII
    let bytes = [data[3], data[2], data[1], data[0]];
    Some(String::from_utf8_lossy(&bytes).to_string())
}

fn find_page_def(records: &[Record<'_>]) -> Option<PageLayout> {
    for rec in records {
        if rec.tag_id == HWPTAG_PAGE_DEF {
            return parse_page_def(rec.data);
        }
        if rec.tag_id == HWPTAG_PARA_HEADER || rec.tag_id == HWPTAG_CTRL_HEADER {
            break;
        }
    }
    None
}

// --- Table control (step 5) ---

fn parse_table_control<'a>(
    records: &[Record<'a>],
    doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> (Option<TableBlock>, usize) {
    // records[0] = CTRL_HEADER for "tbl "
    let ctrl = &records[0];
    if ctrl.tag_id != HWPTAG_CTRL_HEADER {
        return (None, 1);
    }

    let mut table_def: Option<(u16, u16, Option<u16>)> = None; // (row_count, col_count, border_fill_id)
    let mut cell_defs: Vec<CellDef> = Vec::new();
    let mut consumed = 1;

    let base_level = ctrl.level;
    for rec in records.iter().skip(1) {
        if rec.tag_id == HWPTAG_CTRL_HEADER && rec.level <= base_level {
            break;
        }
        if rec.tag_id == HWPTAG_PARA_HEADER && rec.level <= base_level {
            break;
        }
        consumed += 1;

        match rec.tag_id {
            HWPTAG_TABLE => {
                if rec.data.len() >= 8 {
                    let row_count = u16::from_le_bytes(
                        rec.data[4..6].try_into().unwrap_or([0; 2]),
                    );
                    let col_count = u16::from_le_bytes(
                        rec.data[6..8].try_into().unwrap_or([0; 2]),
                    );
                    let bf_id = if rec.data.len() >= 14 {
                        Some(u16::from_le_bytes(rec.data[12..14].try_into().unwrap_or([0; 2])))
                    } else {
                        None
                    };
                    table_def = Some((row_count, col_count, bf_id));
                }
            }
            HWPTAG_LIST_HEADER => {
                let cell = parse_list_header_as_cell(rec.data);
                // Collect paragraphs belonging to this cell
                let cell_level = rec.level;
                let mut cell_blocks = Vec::new();
                while consumed < records.len() {
                    let next = &records[consumed];
                    if next.level <= cell_level {
                        break;
                    }
                    if next.tag_id == HWPTAG_PARA_HEADER {
                        let (para, pconsumed) =
                            parse_paragraph_with_style(&records[consumed..], doc_info, assets);
                        if let Some(p) = para {
                            cell_blocks.push(Block::Paragraph(p));
                        }
                        consumed += pconsumed;
                    } else {
                        consumed += 1;
                    }
                }
                cell_defs.push(CellDef {
                    col_addr: cell.0,
                    row_addr: cell.1,
                    col_span: cell.2,
                    row_span: cell.3,
                    width: cell.4,
                    height: cell.5,
                    blocks: cell_blocks,
                });
            }
            _ => {}
        }
    }

    let (row_count, col_count, bf_id) = match table_def {
        Some(t) => t,
        None => return (None, consumed),
    };

    let table_style = bf_id
        .and_then(|id| doc_info.border_fills.get(id as usize))
        .map(border_fill_to_cell_style);
    let mut table = assemble_table(row_count, col_count, cell_defs);
    table.style = table_style;
    (Some(table), consumed)
}

#[derive(Debug)]
struct CellDef {
    col_addr: u16,
    row_addr: u16,
    col_span: u16,
    row_span: u16,
    width: u32,
    height: u32,
    blocks: Vec<Block>,
}

fn parse_list_header_as_cell(data: &[u8]) -> (u16, u16, u16, u16, u32, u32) {
    // para_count(2) + attributes(4) + ... cell info starts at offset 6
    if data.len() < 22 {
        return (0, 0, 1, 1, 0, 0);
    }
    let col_addr = u16::from_le_bytes(data[6..8].try_into().unwrap_or([0; 2]));
    let row_addr = u16::from_le_bytes(data[8..10].try_into().unwrap_or([0; 2]));
    let col_span = u16::from_le_bytes(data[10..12].try_into().unwrap_or([0; 2])).max(1);
    let row_span = u16::from_le_bytes(data[12..14].try_into().unwrap_or([0; 2])).max(1);
    let width = u32::from_le_bytes(data[14..18].try_into().unwrap_or([0; 4]));
    let height = u32::from_le_bytes(data[18..22].try_into().unwrap_or([0; 4]));
    (col_addr, row_addr, col_span, row_span, width, height)
}

fn assemble_table(row_count: u16, _col_count: u16, cell_defs: Vec<CellDef>) -> TableBlock {
    let mut rows: Vec<TableRow> = (0..row_count).map(|_| TableRow { cells: Vec::new() }).collect();

    for cell in cell_defs {
        let r = cell.row_addr as usize;
        if r >= rows.len() {
            continue;
        }
        let text = cell
            .blocks
            .iter()
            .filter_map(|b| match b {
                Block::Paragraph(p) => Some(
                    p.runs
                        .iter()
                        .map(|r| r.text.as_str())
                        .collect::<String>(),
                ),
                _ => None,
            })
            .collect::<Vec<_>>()
            .join("\n");

        rows[r].cells.push(TableCell {
            text,
            blocks: cell.blocks,
            col_span: if cell.col_span > 1 {
                Some(cell.col_span as u32)
            } else {
                None
            },
            row_span: if cell.row_span > 1 {
                Some(cell.row_span as u32)
            } else {
                None
            },
            width: if cell.width > 0 {
                Some(cell.width as i32)
            } else {
                None
            },
            height: if cell.height > 0 {
                Some(cell.height as i32)
            } else {
                None
            },
            ..TableCell::default()
        });
    }

    // Remove empty trailing rows
    while rows.last().map_or(false, |r| r.cells.is_empty()) {
        rows.pop();
    }

    TableBlock {
        rows,
        no_adjust: false,
        ..TableBlock::default()
    }
}

// --- Header/Footer control ---

fn parse_header_footer_control<'a>(
    records: &[Record<'a>],
    doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> (Vec<Block>, usize) {
    let ctrl = &records[0];
    let base_level = ctrl.level;
    let mut consumed = 1;
    let mut hf_blocks = Vec::new();

    for rec in records.iter().skip(1) {
        if rec.tag_id == HWPTAG_CTRL_HEADER && rec.level <= base_level {
            break;
        }
        if rec.tag_id == HWPTAG_PARA_HEADER && rec.level <= base_level {
            break;
        }
        if rec.tag_id == HWPTAG_PARA_HEADER {
            let (para, pconsumed) =
                parse_paragraph_with_style(&records[consumed..], doc_info, assets);
            if let Some(p) = para {
                hf_blocks.push(Block::Paragraph(p));
            }
            consumed += pconsumed;
        } else {
            consumed += 1;
        }
    }

    (hf_blocks, consumed)
}

// --- GSO control / image (step 6) ---

fn parse_gso_control<'a>(
    records: &[Record<'a>],
    _doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> (Option<ImageBlock>, usize) {
    let ctrl = &records[0];
    if ctrl.tag_id != HWPTAG_CTRL_HEADER || ctrl.data.len() < 32 {
        return (None, 1);
    }

    // Parse common object attributes from CTRL_HEADER
    // offset 4..8: attributes
    let obj_attr = u32::from_le_bytes(ctrl.data[4..8].try_into().unwrap_or([0; 4]));
    let treat_as_char = (obj_attr >> 18) & 1 != 0;

    // offset 20..24: width, 24..28: height
    let width = if ctrl.data.len() >= 24 {
        Some(i32::from_le_bytes(
            ctrl.data[20..24].try_into().unwrap_or([0; 4]),
        ))
    } else {
        None
    };
    let height = if ctrl.data.len() >= 28 {
        Some(i32::from_le_bytes(
            ctrl.data[24..28].try_into().unwrap_or([0; 4]),
        ))
    } else {
        None
    };

    let base_level = ctrl.level;
    let mut consumed = 1;
    let mut bin_item_id: Option<u16> = None;
    let mut shape_type = "shape";

    for rec in records.iter().skip(1) {
        if rec.tag_id == HWPTAG_CTRL_HEADER && rec.level <= base_level {
            break;
        }
        if rec.tag_id == HWPTAG_PARA_HEADER && rec.level <= base_level {
            break;
        }
        consumed += 1;

        if rec.tag_id == HWPTAG_SHAPE_COMPONENT && rec.data.len() >= 4 {
            // shape component: first 4 bytes = shape type id
            // Look for $pic marker — if data contains a BinData reference
            if rec.data.len() >= 8 {
                // Search for BIN reference in shape data
                // The BinDataID is typically after the shape type identification
                // Try to find it by scanning for a plausible index
                for offset in (4..rec.data.len().saturating_sub(1)).step_by(2) {
                    if offset + 2 > rec.data.len() {
                        break;
                    }
                    let candidate =
                        u16::from_le_bytes(rec.data[offset..offset + 2].try_into().unwrap_or([0; 2]));
                    // Check if this index maps to an existing asset
                    if candidate > 0 && (candidate as usize) <= assets.len() {
                        bin_item_id = Some(candidate);
                        shape_type = "pic";
                        break;
                    }
                }
            }
        }
    }

    if let Some(bin_id) = bin_item_id {
        let asset_index = (bin_id as usize).saturating_sub(1);
        let asset_id = assets
            .get(asset_index)
            .map(|a| a.id.clone())
            .unwrap_or_else(|| format!("BIN{:04X}", bin_id));

        return (
            Some(ImageBlock {
                kind: shape_type.to_string(),
                asset_id,
                alt_text: None,
                width,
                height,
                treat_as_char,
                ..ImageBlock::default()
            }),
            consumed,
        );
    }

    (None, consumed)
}

// --- Page definition (step 7) ---

fn parse_page_def(data: &[u8]) -> Option<PageLayout> {
    if data.len() < 40 {
        return None;
    }
    let width = i32::from_le_bytes(data[0..4].try_into().ok()?);
    let height = i32::from_le_bytes(data[4..8].try_into().ok()?);
    let margin_left = i32::from_le_bytes(data[8..12].try_into().ok()?);
    let margin_right = i32::from_le_bytes(data[12..16].try_into().ok()?);
    let margin_top = i32::from_le_bytes(data[16..20].try_into().ok()?);
    let margin_bottom = i32::from_le_bytes(data[20..24].try_into().ok()?);
    let margin_header = i32::from_le_bytes(data[24..28].try_into().ok()?);
    let margin_footer = i32::from_le_bytes(data[28..32].try_into().ok()?);
    let margin_gutter = i32::from_le_bytes(data[32..36].try_into().ok()?);
    let attributes = u32::from_le_bytes(data[36..40].try_into().ok()?);
    let landscape = attributes & 1 != 0;

    Some(PageLayout {
        width: Some(width),
        height: Some(height),
        margin_left: Some(margin_left),
        margin_right: Some(margin_right),
        margin_top: Some(margin_top),
        margin_bottom: Some(margin_bottom),
        margin_header: Some(margin_header),
        margin_footer: Some(margin_footer),
        margin_gutter: Some(margin_gutter),
        landscape,
        page_border: None,
    })
}

// ---------------------------------------------------------------------------
// Stream decoding helper
// ---------------------------------------------------------------------------

fn try_decode_stream(bytes: &[u8]) -> Vec<u8> {
    if bytes.is_empty() {
        return Vec::new();
    }
    // Try decompression first — compressed streams may have headers that look like records
    if let Some(decoded) = try_decode_zlib(bytes) {
        if !decoded.is_empty() && looks_like_record_stream(&decoded) {
            return decoded;
        }
    }
    if let Some(decoded) = try_decode_deflate(bytes) {
        if !decoded.is_empty() && looks_like_record_stream(&decoded) {
            return decoded;
        }
    }
    bytes.to_vec()
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

impl HwpInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Hwp,
            status: "basic".to_string(),
            implemented: vec![
                "CFB signature probing".to_string(),
                "FileHeader version and attribute decoding".to_string(),
                "DocInfo record decoding (FaceName, CharShape, ParaShape, Style, BinData)".to_string(),
                "BodyText paragraph reconstruction with style application".to_string(),
                "Table control parsing (CTRL_HEADER, TABLE, LIST_HEADER)".to_string(),
                "Image/shape control parsing (gso, SHAPE_COMPONENT)".to_string(),
                "Page definition (PAGE_DEF) and section layout".to_string(),
                "BinData asset loading with base64 data URI".to_string(),
                "PrvText Unicode preview extraction as fallback".to_string(),
            ],
            planned: vec![
                "Border/fill style application".to_string(),
                "Footnote and endnote reconstruction".to_string(),
                "Advanced drawing object restoration".to_string(),
            ],
        }
    }

    pub fn parse_bytes(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
    ) -> Result<HwpParseResult, ParseError> {
        let diagnostics = self.inspect_bytes(bytes)?;
        let mut compound_file = open_compound_file(bytes)?;

        // Step 1-3: read DocInfo
        let doc_info = read_doc_info(&mut compound_file);

        // Step 3: load BinData assets
        let assets = load_bin_data_assets(&mut compound_file, &doc_info.bin_data_refs);

        // Steps 4-7: read body sections with style and control parsing
        let body_sections = if !diagnostics.is_encrypted && compound_file.exists(BODY_TEXT_STORAGE)
        {
            read_body_sections_with_docinfo(&mut compound_file, &doc_info, &assets)?
        } else {
            Vec::new()
        };

        let preview_text = read_preview_text(&mut compound_file)?;

        let mut sections = body_sections;
        if sections.is_empty() {
            sections.push(Section {
                id: 0,
                blocks: preview_blocks(preview_text.as_deref()),
                page_layout: None,
                headers: Vec::new(),
                footers: Vec::new(),
                page_start_number: None,
            });
        }

        let page_count = if diagnostics.section_count > 0 {
            diagnostics.section_count
        } else {
            sections.len()
        };

        Ok(HwpParseResult {
            document: Document {
                format: Some(DocumentFormat::Hwp),
                metadata: DocumentMetadata {
                    title: fallback_title.map(ToOwned::to_owned),
                    page_count: Some(page_count as u32),
                    ..DocumentMetadata::default()
                },
                sections,
                assets,
            },
            diagnostics,
        })
    }
}

fn read_body_sections_with_docinfo(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
    doc_info: &DocInfoStore,
    assets: &[AssetRef],
) -> Result<Vec<Section>, ParseError> {
    let mut section_paths = collect_entry_paths(compound_file)
        .into_iter()
        .filter(|path| path.starts_with("/BodyText/Section"))
        .collect::<Vec<_>>();
    section_paths.sort_by_key(|path| section_index(path));

    let mut sections = Vec::new();
    for (index, path) in section_paths.iter().enumerate() {
        let bytes = read_stream_bytes(compound_file, path)?;
        let mut blocks = parse_body_section_with_docinfo(&bytes, doc_info, assets);

        // Extract page_layout and header/footer from sentinel blocks
        let mut page_layout = None;
        let mut header_texts: Vec<String> = Vec::new();
        let mut footer_texts: Vec<String> = Vec::new();
        let mut collecting_header = false;
        let mut collecting_footer = false;
        blocks.retain(|b| {
            if let Block::Unsupported(u) = b {
                match u.kind.as_str() {
                    "__page_layout__" => {
                        if let Some(ref reason) = u.reason {
                            page_layout = deserialize_page_layout(reason);
                        }
                        collecting_header = false;
                        collecting_footer = false;
                        return false;
                    }
                    "__header_blocks__" => {
                        collecting_header = true;
                        collecting_footer = false;
                        return false;
                    }
                    "__footer_blocks__" => {
                        collecting_footer = true;
                        collecting_header = false;
                        return false;
                    }
                    "__header_content__" => {
                        if let Some(ref reason) = u.reason {
                            header_texts.push(reason.clone());
                        }
                        return false;
                    }
                    "__footer_content__" => {
                        if let Some(ref reason) = u.reason {
                            footer_texts.push(reason.clone());
                        }
                        return false;
                    }
                    _ => {}
                }
            }
            collecting_header = false;
            collecting_footer = false;
            true
        });

        let headers = if header_texts.is_empty() {
            Vec::new()
        } else {
            vec![max_viewer_core::HeaderFooter {
                apply_page_type: Some("BOTH".to_string()),
                blocks: header_texts.into_iter().map(|text| {
                    Block::Paragraph(Paragraph {
                        runs: vec![TextRun { text, style: None }],
                        ..Paragraph::default()
                    })
                }).collect(),
            }]
        };
        let footers = if footer_texts.is_empty() {
            Vec::new()
        } else {
            vec![max_viewer_core::HeaderFooter {
                apply_page_type: Some("BOTH".to_string()),
                blocks: footer_texts.into_iter().map(|text| {
                    Block::Paragraph(Paragraph {
                        runs: vec![TextRun { text, style: None }],
                        ..Paragraph::default()
                    })
                }).collect(),
            }]
        };

        if !blocks.is_empty() || page_layout.is_some() {
            sections.push(Section {
                id: index,
                blocks,
                page_layout,
                headers,
                footers,
                page_start_number: None,
            });
        }
    }

    Ok(sections)
}

// ---------------------------------------------------------------------------
// Inspector (unchanged)
// ---------------------------------------------------------------------------

impl FormatInspector for HwpInspector {
    fn format(&self) -> DocumentFormat {
        DocumentFormat::Hwp
    }

    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError> {
        if bytes.len() < CFB_SIGNATURE.len() {
            return Err(ParseError::InvalidContainer(
                "buffer is too short for a CFB header".to_string(),
            ));
        }

        if bytes[..CFB_SIGNATURE.len()] != CFB_SIGNATURE {
            return Err(ParseError::UnsupportedFormat(
                "missing Microsoft Compound File signature".to_string(),
            ));
        }

        let mut compound_file = open_compound_file(bytes)?;
        let header = read_file_header(&mut compound_file)?;
        let entry_paths = collect_entry_paths(&compound_file);
        let section_count = entry_paths
            .iter()
            .filter(|path| path.starts_with("/BodyText/Section"))
            .count();
        let asset_count = entry_paths
            .iter()
            .filter(|path| path.starts_with("/BinData/"))
            .count();
        let has_preview = entry_paths.iter().any(|path| path == PREVIEW_TEXT_STREAM);

        let mut notes = vec![
            "Legacy HWP files are detected via the Compound File signature.".to_string(),
        ];

        if has_preview {
            notes.push("PrvText preview stream is present.".to_string());
        }
        if header.attributes & (1 << 0) != 0 {
            notes.push("Body streams are marked as compressed.".to_string());
        }
        if header.attributes & (1 << 2) != 0 {
            notes.push("Document is marked as distributable.".to_string());
        }
        if header.attributes & (1 << 4) != 0 {
            notes.push("Document is marked as DRM protected.".to_string());
        }

        Ok(DocumentDiagnostics {
            format: DocumentFormat::Hwp,
            entry_count: entry_paths.len(),
            section_count,
            asset_count,
            is_encrypted: header.attributes & (1 << 1) != 0,
            version_hint: header.version_hint,
            notes,
        })
    }
}

// ---------------------------------------------------------------------------
// Utility functions (preserved)
// ---------------------------------------------------------------------------

fn open_compound_file(bytes: &[u8]) -> Result<CompoundFile<Cursor<&[u8]>>, ParseError> {
    CompoundFile::open(Cursor::new(bytes))
        .map_err(|error| ParseError::InvalidContainer(error.to_string()))
}

fn collect_entry_paths(compound_file: &CompoundFile<Cursor<&[u8]>>) -> Vec<String> {
    compound_file
        .walk()
        .filter(|entry| entry.is_stream())
        .map(|entry| entry.path().display().to_string())
        .collect()
}

fn read_file_header(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
) -> Result<HwpHeader, ParseError> {
    if !compound_file.exists(FILE_HEADER_STREAM) {
        return Ok(HwpHeader::default());
    }
    let mut stream = compound_file
        .open_stream(FILE_HEADER_STREAM)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    let mut bytes = Vec::new();
    stream
        .read_to_end(&mut bytes)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;

    if bytes.len() < 40 {
        return Err(ParseError::InvalidData(
            "FileHeader stream is shorter than the documented fixed header size.".to_string(),
        ));
    }

    let signature = bytes[..32]
        .iter()
        .copied()
        .take_while(|byte| *byte != 0)
        .collect::<Vec<_>>();
    let signature_text = String::from_utf8_lossy(&signature).trim().to_string();
    if !signature_text.contains("HWP Document File") {
        return Err(ParseError::UnsupportedFormat(format!(
            "unexpected FileHeader signature: {signature_text}"
        )));
    }

    let version_raw = u32::from_le_bytes(bytes[32..36].try_into().unwrap_or([0, 0, 0, 0]));
    let attributes = u32::from_le_bytes(bytes[36..40].try_into().unwrap_or([0, 0, 0, 0]));

    Ok(HwpHeader {
        version_hint: Some(format_version(version_raw)),
        attributes,
    })
}

fn format_version(version: u32) -> String {
    format!(
        "{}.{}.{}.{}",
        (version >> 24) & 0xff,
        (version >> 16) & 0xff,
        (version >> 8) & 0xff,
        version & 0xff
    )
}

fn read_preview_text(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
) -> Result<Option<String>, ParseError> {
    if !compound_file.exists(PREVIEW_TEXT_STREAM) {
        return Ok(None);
    }
    let mut stream = compound_file
        .open_stream(PREVIEW_TEXT_STREAM)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    let mut bytes = Vec::new();
    stream
        .read_to_end(&mut bytes)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;

    if bytes.is_empty() {
        return Ok(None);
    }
    let text = decode_preview_text(&bytes);
    if text.trim().is_empty() {
        Ok(None)
    } else {
        Ok(Some(text))
    }
}

fn preview_blocks(preview_text: Option<&str>) -> Vec<Block> {
    let mut blocks = Vec::new();
    if let Some(text) = preview_text {
        for paragraph in split_preview_paragraphs(text) {
            blocks.push(Block::Paragraph(Paragraph {
                marker: None,
                runs: vec![TextRun {
                    text: paragraph.to_string(),
                    style: None,
                }],
                style: None,
                line_segment_count: None,
                layout_height_hint: None,
                page_break_before: false,
            }));
        }
    }
    if blocks.is_empty() {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "hwp-preview".to_string(),
            reason: Some(
                "PrvText preview stream was missing, so only container diagnostics are available."
                    .to_string(),
            ),
        }));
    }
    blocks
}

fn read_stream_bytes(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
    path: &str,
) -> Result<Vec<u8>, ParseError> {
    let mut stream = compound_file
        .open_stream(path)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    stream
        .seek(SeekFrom::Start(0))
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    let mut bytes = Vec::new();
    stream
        .read_to_end(&mut bytes)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    Ok(bytes)
}

fn section_index(path: &str) -> usize {
    path.rsplit('/')
        .next()
        .unwrap_or(path)
        .trim_start_matches("Section")
        .parse::<usize>()
        .unwrap_or(usize::MAX)
}

fn try_decode_zlib(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut decoder = ZlibDecoder::new(bytes);
    let mut decoded = Vec::new();
    decoder.read_to_end(&mut decoded).ok()?;
    (!decoded.is_empty()).then_some(decoded)
}

fn try_decode_deflate(bytes: &[u8]) -> Option<Vec<u8>> {
    let mut decoder = DeflateDecoder::new(bytes);
    let mut decoded = Vec::new();
    decoder.read_to_end(&mut decoded).ok()?;
    (!decoded.is_empty()).then_some(decoded)
}

fn looks_like_record_stream(bytes: &[u8]) -> bool {
    if bytes.len() < 4 {
        return false;
    }
    let header = u32::from_le_bytes([bytes[0], bytes[1], bytes[2], bytes[3]]);
    let tag_id = (header & 0x3ff) as u16;
    tag_id >= HWPTAG_BEGIN && tag_id < 0x400
}

fn decode_preview_text(bytes: &[u8]) -> String {
    if bytes.len() >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe {
        return decode_utf16_le(&bytes[2..]);
    }
    if bytes.len() >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff {
        return decode_utf16_be(&bytes[2..]);
    }
    if bytes.len().is_multiple_of(2) {
        let text = decode_utf16_le(bytes);
        if !text.trim_matches('\u{0}').is_empty() {
            return text;
        }
    }
    String::from_utf8_lossy(bytes)
        .trim_matches('\u{0}')
        .to_string()
}

fn decode_utf16_le(bytes: &[u8]) -> String {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&units)
        .trim_matches('\u{0}')
        .to_string()
}

fn decode_utf16_be(bytes: &[u8]) -> String {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    String::from_utf16_lossy(&units)
        .trim_matches('\u{0}')
        .to_string()
}

fn split_preview_paragraphs(text: &str) -> Vec<&str> {
    text.lines()
        .map(str::trim)
        .filter(|line| !line.is_empty())
        .collect()
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::{Compression, write::ZlibEncoder};
    use std::io::Write;

    fn write_record(output: &mut Vec<u8>, tag_id: u16, level: u16, data: &[u8]) {
        let size = data.len();
        let header = if size >= 0x0fff {
            ((0x0fffu32) << 20) | ((level as u32) << 10) | (tag_id as u32)
        } else {
            ((size as u32) << 20) | ((level as u32) << 10) | (tag_id as u32)
        };
        output.extend_from_slice(&header.to_le_bytes());
        if size >= 0x0fff {
            output.extend_from_slice(&(size as u32).to_le_bytes());
        }
        output.extend_from_slice(data);
    }

    fn make_file_header_stream(compound: &mut CompoundFile<Cursor<Vec<u8>>>, attributes: u32) {
        let mut stream = compound.create_stream(FILE_HEADER_STREAM).unwrap();
        let mut header = vec![0u8; 32];
        header[..17].copy_from_slice(b"HWP Document File");
        stream.write_all(&header).unwrap();
        stream.write_all(&0x05000300u32.to_le_bytes()).unwrap();
        stream.write_all(&attributes.to_le_bytes()).unwrap();
    }

    #[test]
    fn parses_hwp_preview_text() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        make_file_header_stream(&mut compound, 0);

        {
            let mut stream = compound.create_stream(PREVIEW_TEXT_STREAM).unwrap();
            for unit in "첫째 줄\n둘째 줄".encode_utf16() {
                stream.write_all(&unit.to_le_bytes()).unwrap();
            }
        }

        let bytes = compound.into_inner().into_inner();
        let parsed = HwpInspector
            .parse_bytes(&bytes, Some("fixture.hwp"))
            .expect("preview extraction should work");

        assert_eq!(parsed.diagnostics.format, DocumentFormat::Hwp);
        assert_eq!(parsed.document.sections.len(), 1);
        assert_eq!(
            parsed.document.metadata.title.as_deref(),
            Some("fixture.hwp")
        );
        assert_eq!(parsed.document.sections[0].blocks.len(), 2);
    }

    #[test]
    fn parses_hwp_bodytext_section() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        make_file_header_stream(&mut compound, 0);

        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();
            let mut para_header = vec![0u8; 22];
            para_header[..4].copy_from_slice(&3u32.to_le_bytes());
            write_record(&mut payload, HWPTAG_PARA_HEADER, 0, &para_header);

            let text_units = "본문"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>();
            write_record(&mut payload, HWPTAG_PARA_TEXT, 1, &text_units);

            let mut stream = compound.create_stream("/BodyText/Section0").unwrap();
            stream.write_all(&payload).unwrap();
        }

        let bytes = compound.into_inner().into_inner();
        let parsed = HwpInspector
            .parse_bytes(&bytes, Some("bodytext.hwp"))
            .expect("body text should parse");

        assert_eq!(parsed.document.sections.len(), 1);
        assert!(!parsed.document.sections[0].blocks.is_empty());
        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert!(paragraph.runs[0].text.contains("본문"));
            }
            _ => panic!("expected paragraph block"),
        }
    }

    #[test]
    fn parses_compressed_hwp_bodytext_section() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        make_file_header_stream(&mut compound, 1);

        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();
            let mut para_header = vec![0u8; 22];
            para_header[..4].copy_from_slice(&2u32.to_le_bytes());
            write_record(&mut payload, HWPTAG_PARA_HEADER, 0, &para_header);

            let text_units = "압축"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>();
            write_record(&mut payload, HWPTAG_PARA_TEXT, 1, &text_units);

            let mut encoder = ZlibEncoder::new(Vec::new(), Compression::default());
            encoder.write_all(&payload).unwrap();
            let compressed = encoder.finish().unwrap();

            let mut stream = compound.create_stream("/BodyText/Section0").unwrap();
            stream.write_all(&compressed).unwrap();
        }

        let bytes = compound.into_inner().into_inner();
        let parsed = HwpInspector
            .parse_bytes(&bytes, Some("compressed.hwp"))
            .expect("compressed body text should parse");

        let section = &parsed.document.sections[0];
        assert!(
            !section.blocks.is_empty(),
            "blocks should not be empty, got {} sections with {} blocks in first",
            parsed.document.sections.len(),
            section.blocks.len()
        );
        match &section.blocks[0] {
            Block::Paragraph(paragraph) => {
                assert!(paragraph.runs[0].text.contains("압축"));
            }
            other => panic!("expected paragraph block, got {:?}", other),
        }
    }

    #[test]
    fn parses_docinfo_char_shape_and_applies_style() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        make_file_header_stream(&mut compound, 0);

        // Create DocInfo with a face name and char shape
        {
            let mut docinfo_payload = Vec::new();

            // FACE_NAME: attribute(1) + name_len(2) + name(UTF16)
            let face_name = "함초롬바탕";
            let name_units: Vec<u8> = face_name
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect();
            let mut face_data = Vec::new();
            face_data.push(0); // attributes
            face_data.extend_from_slice(&(face_name.encode_utf16().count() as u16).to_le_bytes());
            face_data.extend_from_slice(&name_units);
            write_record(&mut docinfo_payload, HWPTAG_FACE_NAME, 0, &face_data);

            // CHAR_SHAPE: 72 bytes minimum
            let mut cs_data = vec![0u8; 72];
            // face_name_ids[0] = 0 (index of face above)
            cs_data[0..2].copy_from_slice(&0u16.to_le_bytes());
            // height = 1000 (10pt in 1/100 pt)
            cs_data[42..46].copy_from_slice(&1000u32.to_le_bytes());
            // attributes: bold (bit 1)
            cs_data[46..50].copy_from_slice(&2u32.to_le_bytes());
            // text_color: red = 0x0000FF (COLORREF BGR)
            cs_data[52..56].copy_from_slice(&0x0000FFu32.to_le_bytes());
            write_record(&mut docinfo_payload, HWPTAG_CHAR_SHAPE, 0, &cs_data);

            let mut stream = compound.create_stream(DOC_INFO_STREAM).unwrap();
            stream.write_all(&docinfo_payload).unwrap();
        }

        // Create BodyText with a paragraph referencing char shape 0
        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();

            // PARA_HEADER: 22 bytes, paraShapeId at offset 8
            let mut ph = vec![0u8; 22];
            ph[..4].copy_from_slice(&4u32.to_le_bytes()); // char count
            write_record(&mut payload, HWPTAG_PARA_HEADER, 0, &ph);

            // PARA_TEXT
            let text_units = "스타일"
                .encode_utf16()
                .flat_map(u16::to_le_bytes)
                .collect::<Vec<_>>();
            write_record(&mut payload, HWPTAG_PARA_TEXT, 1, &text_units);

            // PARA_CHAR_SHAPE: position 0, charShapeId 0
            let mut pcs = Vec::new();
            pcs.extend_from_slice(&0u32.to_le_bytes()); // position
            pcs.extend_from_slice(&0u32.to_le_bytes()); // charShapeId
            write_record(&mut payload, HWPTAG_PARA_CHAR_SHAPE, 1, &pcs);

            let mut stream = compound.create_stream("/BodyText/Section0").unwrap();
            stream.write_all(&payload).unwrap();
        }

        let bytes = compound.into_inner().into_inner();
        let parsed = HwpInspector
            .parse_bytes(&bytes, Some("styled.hwp"))
            .expect("styled body text should parse");

        assert!(!parsed.document.sections[0].blocks.is_empty());
        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert!(paragraph.runs[0].text.contains("스타일"));
                let style = paragraph.runs[0].style.as_ref().expect("should have style");
                assert_eq!(style.font_family.as_deref(), Some("함초롬바탕"));
                assert_eq!(style.font_size, Some(10));
                assert!(style.bold);
                assert_eq!(style.text_color.as_deref(), Some("#ff0000"));
            }
            _ => panic!("expected paragraph block"),
        }
    }

    #[test]
    fn handles_empty_and_corrupt_input_gracefully() {
        // Empty bytes
        assert!(HwpInspector.parse_bytes(&[], None).is_err());

        // Too short
        assert!(HwpInspector.parse_bytes(&[0x00, 0x01], None).is_err());

        // Wrong signature
        assert!(HwpInspector.parse_bytes(&[0xFF; 100], None).is_err());

        // Valid CFB but truncated FileHeader
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        {
            let mut stream = compound.create_stream(FILE_HEADER_STREAM).unwrap();
            stream.write_all(&[0u8; 20]).unwrap(); // Too short for valid header
        }
        let bytes = compound.into_inner().into_inner();
        assert!(HwpInspector.parse_bytes(&bytes, None).is_err());
    }
}

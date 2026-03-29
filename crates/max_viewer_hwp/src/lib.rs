use std::io::{Cursor, Read, Seek, SeekFrom};
use std::path::Path;

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
const HWPTAG_PARA_LINE_SEG: u16 = HWPTAG_BEGIN + 53;
const HWPTAG_PARA_RANGE_TAG: u16 = HWPTAG_BEGIN + 54;
const HWPTAG_CTRL_HEADER: u16 = HWPTAG_BEGIN + 55;
const HWPTAG_LIST_HEADER: u16 = HWPTAG_BEGIN + 56;
const HWPTAG_PAGE_DEF: u16 = HWPTAG_BEGIN + 57;
const HWPTAG_FOOTNOTE_SHAPE: u16 = HWPTAG_BEGIN + 58;
const HWPTAG_PAGE_BORDER_FILL: u16 = HWPTAG_BEGIN + 59;
const HWPTAG_SHAPE_COMPONENT: u16 = HWPTAG_BEGIN + 60;
const HWPTAG_TABLE: u16 = HWPTAG_BEGIN + 61;
const HWPTAG_SHAPE_COMPONENT_PICTURE: u16 = HWPTAG_BEGIN + 69;

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

#[derive(Debug, Clone, Copy, Default)]
struct HwpLineSegment {
    chpos: i32,
    y: i32,
    height: i32,
    space_below: i32,
}

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
    underline_color: u32,
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
    background_image: Option<String>,
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
    cs.underline_color = u32::from_le_bytes(data[56..60].try_into().unwrap_or([0; 4]));
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
    if data.len() < 26 {
        return bf;
    }
    let read_border = |offset: usize| -> Option<TableBorder> {
        let border_type = data.get(offset)? & 0x0f;
        let border_width = data.get(offset + 1)? & 0x0f;
        let color = u32::from_le_bytes(data[offset + 2..offset + 6].try_into().ok()?);

        let style_name = match border_type {
            0 => "NONE",
            1 => "SOLID",
            2 => "DASH",
            3 => "DOT",
            4 => "DASH_DOT",
            5 => "DASH_DOT_DOT",
            6 => "LONG_DASH",
            7 => "DOT",
            8 => "DOUBLE",
            9 => "DOUBLE_SLIM",
            10 => "SLIM_THICK",
            11 => "THICK_SLIM",
            12 => "SOLID",
            13 => "DOUBLE",
            14 => "INSET",
            15 => "OUTSET",
            16 => "GROOVE",
            17 => "RIDGE",
            _ => "SOLID",
        };
        let width_mm = match border_width {
            0 => 0.10,
            1 => 0.12,
            2 => 0.15,
            3 => 0.20,
            4 => 0.25,
            5 => 0.30,
            6 => 0.40,
            7 => 0.50,
            8 => 0.60,
            9 => 0.70,
            10 => 1.00,
            11 => 1.50,
            12 => 2.00,
            13 => 3.00,
            14 => 4.00,
            15 => 5.00,
            _ => 0.10,
        };

        Some(TableBorder {
            style: Some(style_name.to_string()),
            width: Some(format!("{width_mm:.2} mm")),
            color: Some(colorref_to_hex(color)),
        })
    };
    // NOTE:
    // HWP 5.0 BORDER_FILL is serialized as:
    //   attribute(u16)
    //   left/right/top/bottom Border(6 bytes each)
    //   diagonal Border(6 bytes)
    //   Fill
    // Some public references list separate border type/width/color arrays, but
    // actual HWP files and the reference parser on docs.rs read repeated border
    // structs. Using array-style offsets misreads diagonal/fill bytes as colors
    // and turns decorative cover tables into black/gray boxes.
    bf.border_left = read_border(2);
    bf.border_right = read_border(8);
    bf.border_top = read_border(14);
    bf.border_bottom = read_border(20);
    let fill_offset = 32;
    if data.len() >= fill_offset + 4 {
        let fill_type = u32::from_le_bytes(
            data[fill_offset..fill_offset + 4]
                .try_into()
                .unwrap_or([0; 4]),
        );
        if fill_type & 0x0000_0001 != 0 {
            if data.len() >= fill_offset + 8 {
                let background_color = u32::from_le_bytes(
                    data[fill_offset + 4..fill_offset + 8]
                        .try_into()
                        .unwrap_or([0; 4]),
                );
                bf.background_color = Some(colorref_to_hex(background_color));
            }
        }
        if fill_type & 0x0000_0004 != 0 {
            bf.background_image = parse_gradation_fill_css(&data[fill_offset + 4..]);
        }
    }
    bf
}

fn border_fill_to_cell_style(bf: &HwpBorderFill) -> TableCellStyle {
    TableCellStyle {
        background_color: bf.background_color.clone(),
        background_image: bf.background_image.clone(),
        border_left: bf.border_left.clone(),
        border_right: bf.border_right.clone(),
        border_top: bf.border_top.clone(),
        border_bottom: bf.border_bottom.clone(),
        diagonal: None,
    }
}

fn resolve_border_fill_style(
    doc_info: &DocInfoStore,
    border_fill_id: Option<u16>,
) -> Option<TableCellStyle> {
    border_fill_id
        .and_then(|id| id.checked_sub(1))
        .and_then(|index| doc_info.border_fills.get(index as usize))
        .map(border_fill_to_cell_style)
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
            .filter(|ext| !ext.is_empty())
            .unwrap_or_else(|| {
                Path::new(path)
                    .extension()
                    .and_then(|ext| ext.to_str())
                    .unwrap_or("")
            });
        let media_type = guess_media_type(ext, &bytes);
        let data_uri = format!("data:{};base64,{}", media_type, BASE64.encode(&bytes));
        let id = path.rsplit('/').next().unwrap_or(path).to_string();

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
    let underline_kind = (cs.attributes >> 2) & 0x3;
    let has_visible_underline = matches!(underline_kind, 1 | 3);

    TextStyle {
        font_family: face_names
            .get(cs.face_name_ids[0] as usize)
            .map(|f| f.name.clone()),
        font_size: if cs.height > 0 {
            Some(cs.height as i32)
        } else {
            None
        },
        bold: cs.attributes & (1 << 1) != 0,
        italic: cs.attributes & (1 << 0) != 0,
        underline: has_visible_underline,
        text_color: if cs.text_color != 0 {
            Some(colorref_to_hex(cs.text_color))
        } else {
            None
        },
        underline_color: if has_visible_underline {
            Some(colorref_to_hex(cs.underline_color))
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
    let align_bits = (ps.attributes >> 2) & 0x7;
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

fn read_u8(cursor: &mut Cursor<&[u8]>) -> Option<u8> {
    let mut buf = [0u8; 1];
    cursor.read_exact(&mut buf).ok()?;
    Some(buf[0])
}

fn read_i32(cursor: &mut Cursor<&[u8]>) -> Option<i32> {
    let mut buf = [0u8; 4];
    cursor.read_exact(&mut buf).ok()?;
    Some(i32::from_le_bytes(buf))
}

fn read_u32(cursor: &mut Cursor<&[u8]>) -> Option<u32> {
    let mut buf = [0u8; 4];
    cursor.read_exact(&mut buf).ok()?;
    Some(u32::from_le_bytes(buf))
}

fn parse_gradation_fill_css(data: &[u8]) -> Option<String> {
    let mut cursor = Cursor::new(data);
    let gradation_type = read_u8(&mut cursor)?;
    let angle = read_i32(&mut cursor)?;
    let _center_x = read_i32(&mut cursor)?;
    let _center_y = read_i32(&mut cursor)?;
    let _step = read_i32(&mut cursor)?;
    let color_count = usize::try_from(read_i32(&mut cursor)?).ok()?;
    if color_count < 2 {
        return None;
    }

    for _ in 0..color_count.saturating_sub(2) {
        read_i32(&mut cursor)?;
    }

    let mut colors = Vec::with_capacity(color_count);
    for _ in 0..color_count {
        colors.push(colorref_to_hex(read_u32(&mut cursor)?));
    }

    let _shape_type = read_u32(&mut cursor)?;
    let _blur_center = read_u8(&mut cursor)?;

    let stops = colors.join(", ");
    let css_angle = if angle == 0 { 90 } else { angle };

    match gradation_type {
        1 => Some(format!("linear-gradient({css_angle}deg, {stops})")),
        2 | 3 | 4 => Some(format!("linear-gradient(90deg, {stops})")),
        _ => None,
    }
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
                                page_break_before: false,
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
                        let (hf_blocks, consumed) =
                            parse_header_footer_control(&records[idx..], doc_info, assets);
                        header_blocks = hf_blocks;
                        idx += consumed.max(1);
                    }
                    Some("toof") => {
                        // Footer control
                        let (hf_blocks, consumed) =
                            parse_header_footer_control(&records[idx..], doc_info, assets);
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
            page_break_before: false,
        }));
        blocks.extend(header_blocks.into_iter().map(|b| {
            // Wrap in a sentinel-tagged unsupported to mark as header content
            Block::Unsupported(UnsupportedBlock {
                kind: "__header_content__".to_string(),
                reason: match &b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
                    _ => Some("block".to_string()),
                },
                page_break_before: false,
            })
        }));
    }
    if !footer_blocks.is_empty() {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "__footer_blocks__".to_string(),
            reason: Some(format!("{}", footer_blocks.len())),
            page_break_before: false,
        }));
        blocks.extend(footer_blocks.into_iter().map(|b| {
            Block::Unsupported(UnsupportedBlock {
                kind: "__footer_content__".to_string(),
                reason: match &b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
                    _ => Some("block".to_string()),
                },
                page_break_before: false,
            })
        }));
    }

    if let Some(layout) = page_layout {
        blocks.push(Block::Unsupported(UnsupportedBlock {
            kind: "__page_layout__".to_string(),
            reason: Some(serde_page_layout(&layout)),
            page_break_before: false,
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
    let style_id = if data.len() >= 11 {
        Some(data[10])
    } else {
        None
    };

    let mut text_data: Option<&[u8]> = None;
    let mut char_shape_positions: Vec<(u32, u32)> = Vec::new();
    let mut line_segments: Vec<HwpLineSegment> = Vec::new();
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
            HWPTAG_PARA_LINE_SEG => {
                line_segments = parse_para_line_segs(rec.data);
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
    let runs = apply_line_seg_breaks_to_runs(
        build_styled_runs(&units, &char_shape_positions, doc_info),
        &line_segments,
    );
    if runs.is_empty() {
        return (None, consumed);
    }

    // Resolve paragraph style
    let para_style = resolve_para_style(para_shape_id, style_id, doc_info);
    let line_segment_count = (!line_segments.is_empty()).then_some(line_segments.len() as u32);
    let layout_height_hint = paragraph_layout_height_hint(&line_segments);

    (
        Some(Paragraph {
            marker: None,
            runs,
            style: para_style,
            line_segment_count,
            layout_height_hint,
            page_break_before: false,
        }),
        consumed,
    )
}

fn parse_para_line_segs(data: &[u8]) -> Vec<HwpLineSegment> {
    const LINE_SEG_SIZE: usize = 36;
    if data.len() < LINE_SEG_SIZE {
        return Vec::new();
    }

    data.chunks_exact(LINE_SEG_SIZE)
        .map(|chunk| HwpLineSegment {
            chpos: i32::from_le_bytes(chunk[0..4].try_into().unwrap_or([0; 4])),
            y: i32::from_le_bytes(chunk[4..8].try_into().unwrap_or([0; 4])),
            height: i32::from_le_bytes(chunk[8..12].try_into().unwrap_or([0; 4])),
            space_below: i32::from_le_bytes(chunk[20..24].try_into().unwrap_or([0; 4])),
        })
        .collect()
}

fn paragraph_layout_height_hint(line_segments: &[HwpLineSegment]) -> Option<i32> {
    line_segments
        .iter()
        .map(|seg| {
            seg.y
                .saturating_add(seg.height)
                .saturating_add(seg.space_below)
        })
        .max()
        .filter(|height| *height > 0)
}

fn apply_line_seg_breaks_to_runs(
    runs: Vec<TextRun>,
    line_segments: &[HwpLineSegment],
) -> Vec<TextRun> {
    let mut break_positions: Vec<usize> = line_segments
        .iter()
        .skip(1)
        .filter_map(|seg| usize::try_from(seg.chpos).ok())
        .collect();
    break_positions.sort_unstable();
    break_positions.dedup();
    break_positions.retain(|pos| *pos > 0);

    if break_positions.is_empty() {
        return runs;
    }

    let mut output = Vec::with_capacity(runs.len() + break_positions.len());
    let mut consumed_chars = 0usize;
    let mut next_break_index = 0usize;

    for run in runs {
        let chars: Vec<char> = run.text.chars().collect();
        if chars.is_empty() {
            output.push(run);
            continue;
        }

        let mut local_start = 0usize;
        let run_end = consumed_chars + chars.len();

        while let Some(&break_pos) = break_positions.get(next_break_index) {
            if break_pos > run_end {
                break;
            }

            let local_break = break_pos.saturating_sub(consumed_chars).min(chars.len());
            if local_break > local_start {
                output.push(TextRun {
                    text: chars[local_start..local_break].iter().collect(),
                    style: run.style.clone(),
                });
            }

            output.push(TextRun {
                text: "\n".to_string(),
                style: None,
            });
            local_start = local_break;
            next_break_index += 1;
        }

        if local_start < chars.len() {
            output.push(TextRun {
                text: chars[local_start..].iter().collect(),
                style: run.style,
            });
        }

        consumed_chars = run_end;
    }

    output
}

fn extract_text_units(data: &[u8], char_count: Option<u32>) -> Vec<u16> {
    if data.len() % 2 != 0 {
        return Vec::new();
    }
    let raw_units: Vec<u16> = data
        .chunks_exact(2)
        .map(|c| u16::from_le_bytes([c[0], c[1]]))
        .collect();
    let mut units = Vec::with_capacity(raw_units.len());
    let mut index = 0usize;
    while index < raw_units.len() {
        if let Some(skip_len) = extended_control_span(&raw_units, index) {
            index += skip_len;
            continue;
        }
        units.push(raw_units[index]);
        index += 1;
    }
    if let Some(count) = char_count {
        let expected = count as usize;
        if expected > 0 && expected < units.len() {
            units.truncate(expected);
        }
    }
    units
}

fn extended_control_span(units: &[u16], index: usize) -> Option<usize> {
    let Some(&control_char) = units.get(index) else {
        return None;
    };

    if !(0x01..=0x1f).contains(&control_char) || matches!(control_char, 0x09 | 0x0a | 0x0d) {
        return None;
    }

    // HWP PARA_TEXT stores inline control payloads inline with text as a short
    // control marker + binary payload + matching tail marker sequence. The exact
    // payload width varies across controls, so keep this parser-side and strip
    // only the compact HWP control envelope instead of letting it leak into the UI.
    for end in index + 1..usize::min(index + 8, units.len()) {
        if units[end] == control_char {
            return Some(end + 1 - index);
        }
    }

    // If the control marker is malformed, still drop the lone control code so
    // it cannot surface as mojibake in rendered text.
    Some(1)
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
        let default_style = doc_info
            .char_shapes
            .first()
            .map(|cs| char_shape_to_text_style(cs, &doc_info.face_names));
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
    normalize_hwp_text(text)
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
    let object_width = ctrl
        .data
        .get(16..20)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value > 0);
    let object_height = ctrl
        .data
        .get(20..24)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value > 0);

    let mut table_def: Option<(u16, u16, Option<u16>)> = None; // (row_count, col_count, border_fill_id)
    let mut cell_defs: Vec<CellDef> = Vec::new();
    let base_level = ctrl.level;
    let mut cursor = 1usize;

    while cursor < records.len() {
        let rec = &records[cursor];
        if rec.tag_id == HWPTAG_CTRL_HEADER && rec.level <= base_level {
            break;
        }
        if rec.tag_id == HWPTAG_PARA_HEADER && rec.level <= base_level {
            break;
        }

        match rec.tag_id {
            HWPTAG_TABLE => {
                if rec.data.len() >= 8 {
                    let row_count = u16::from_le_bytes(rec.data[4..6].try_into().unwrap_or([0; 2]));
                    let col_count = u16::from_le_bytes(rec.data[6..8].try_into().unwrap_or([0; 2]));
                    let bf_id = if rec.data.len() >= 14 {
                        Some(u16::from_le_bytes(
                            rec.data[12..14].try_into().unwrap_or([0; 2]),
                        ))
                    } else {
                        None
                    };
                    table_def = Some((row_count, col_count, bf_id));
                }
                cursor += 1;
            }
            HWPTAG_LIST_HEADER => {
                let cell = parse_list_header_as_cell(rec.data);
                // Collect paragraphs belonging to this cell
                let cell_level = rec.level;
                let mut cell_blocks = Vec::new();
                cursor += 1;
                while cursor < records.len() {
                    let next = &records[cursor];
                    if next.tag_id == HWPTAG_LIST_HEADER && next.level <= cell_level {
                        break;
                    }
                    if next.tag_id == HWPTAG_CTRL_HEADER && next.level <= base_level {
                        break;
                    }
                    if next.tag_id == HWPTAG_TABLE && next.level <= base_level {
                        break;
                    }
                    if next.tag_id == HWPTAG_PARA_HEADER {
                        let (para, pconsumed) =
                            parse_paragraph_with_style(&records[cursor..], doc_info, assets);
                        if let Some(p) = para {
                            cell_blocks.push(Block::Paragraph(p));
                        }
                        cursor += pconsumed.max(1);
                    } else if next.tag_id == HWPTAG_CTRL_HEADER {
                        match read_ctrl_type(next.data).as_deref() {
                            Some("gso ") => {
                                let (image, iconsumed) =
                                    parse_gso_control(&records[cursor..], doc_info, assets);
                                if let Some(img) = image {
                                    cell_blocks.push(Block::Image(img));
                                } else {
                                    cell_blocks.push(Block::Unsupported(UnsupportedBlock {
                                        kind: "hwp-gso".to_string(),
                                        reason: Some("drawing object in table cell".to_string()),
                                        page_break_before: false,
                                    }));
                                }
                                cursor += iconsumed.max(1);
                            }
                            Some("tbl ") => {
                                let (table, tconsumed) =
                                    parse_table_control(&records[cursor..], doc_info, assets);
                                if let Some(table) = table {
                                    cell_blocks.push(Block::Table(table));
                                }
                                cursor += tconsumed.max(1);
                            }
                            _ => {
                                cursor += 1;
                            }
                        }
                    } else {
                        cursor += 1;
                    }
                }
                cell_defs.push(CellDef {
                    col_addr: cell.0,
                    row_addr: cell.1,
                    col_span: cell.2,
                    row_span: cell.3,
                    width: cell.4,
                    height: cell.5,
                    margin_left: cell.6,
                    margin_right: cell.7,
                    margin_top: cell.8,
                    margin_bottom: cell.9,
                    border_fill_id: cell.10.filter(|id| *id > 0),
                    blocks: cell_blocks,
                });
            }
            _ => {
                cursor += 1;
            }
        }
    }

    let (row_count, col_count, bf_id) = match table_def {
        Some(t) => t,
        None => return (None, cursor.max(1)),
    };

    let mut table = assemble_table(row_count, col_count, cell_defs, doc_info);
    table.width = object_width;
    table.height = object_height;
    table.style = resolve_border_fill_style(doc_info, bf_id);
    (Some(table), cursor.max(1))
}

#[derive(Debug)]
struct CellDef {
    col_addr: u16,
    row_addr: u16,
    col_span: u16,
    row_span: u16,
    width: u32,
    height: u32,
    margin_left: u16,
    margin_right: u16,
    margin_top: u16,
    margin_bottom: u16,
    border_fill_id: Option<u16>,
    blocks: Vec<Block>,
}

fn parse_list_header_as_cell(
    data: &[u8],
) -> (
    u16,
    u16,
    u16,
    u16,
    u32,
    u32,
    u16,
    u16,
    u16,
    u16,
    Option<u16>,
) {
    // HWP 5.x table-cell LIST_HEADER stores a 4-byte paragraph count and a
    // 4-byte attribute field before the 26-byte cell descriptor.
    // Using the legacy 6-byte offset misreads widths/heights into huge or
    // negative numbers and turns ordinary tables into page-sized black blocks.
    const CELL_DESCRIPTOR_OFFSET: usize = 8;
    if data.len() < CELL_DESCRIPTOR_OFFSET + 16 {
        return (0, 0, 1, 1, 0, 0, 0, 0, 0, 0, None);
    }
    let base = CELL_DESCRIPTOR_OFFSET;
    let col_addr = u16::from_le_bytes(data[base..base + 2].try_into().unwrap_or([0; 2]));
    let row_addr = u16::from_le_bytes(data[base + 2..base + 4].try_into().unwrap_or([0; 2]));
    let col_span = u16::from_le_bytes(data[base + 4..base + 6].try_into().unwrap_or([0; 2])).max(1);
    let row_span = u16::from_le_bytes(data[base + 6..base + 8].try_into().unwrap_or([0; 2])).max(1);
    let width = u32::from_le_bytes(data[base + 8..base + 12].try_into().unwrap_or([0; 4]));
    let height = u32::from_le_bytes(data[base + 12..base + 16].try_into().unwrap_or([0; 4]));
    let margin_left = if data.len() >= base + 18 {
        u16::from_le_bytes(data[base + 16..base + 18].try_into().unwrap_or([0; 2]))
    } else {
        0
    };
    let margin_right = if data.len() >= base + 20 {
        u16::from_le_bytes(data[base + 18..base + 20].try_into().unwrap_or([0; 2]))
    } else {
        0
    };
    let margin_top = if data.len() >= base + 22 {
        u16::from_le_bytes(data[base + 20..base + 22].try_into().unwrap_or([0; 2]))
    } else {
        0
    };
    let margin_bottom = if data.len() >= base + 24 {
        u16::from_le_bytes(data[base + 22..base + 24].try_into().unwrap_or([0; 2]))
    } else {
        0
    };
    let border_fill_id = if data.len() >= base + 26 {
        Some(u16::from_le_bytes(
            data[base + 24..base + 26].try_into().unwrap_or([0; 2]),
        ))
    } else {
        None
    };
    (
        col_addr,
        row_addr,
        col_span,
        row_span,
        width,
        height,
        margin_left,
        margin_right,
        margin_top,
        margin_bottom,
        border_fill_id,
    )
}

fn assemble_table(
    row_count: u16,
    col_count: u16,
    cell_defs: Vec<CellDef>,
    doc_info: &DocInfoStore,
) -> TableBlock {
    let mut rows: Vec<TableRow> = (0..row_count)
        .map(|_| TableRow { cells: Vec::new() })
        .collect();
    let total_cols = usize::from(col_count).max(
        cell_defs
            .iter()
            .map(|cell| usize::from(cell.col_addr.saturating_add(cell.col_span.max(1))))
            .max()
            .unwrap_or(0),
    );
    let mut cells_by_row: Vec<Vec<CellDef>> = (0..row_count).map(|_| Vec::new()).collect();

    for cell in cell_defs {
        let row_index = usize::from(cell.row_addr);
        if row_index < cells_by_row.len() {
            cells_by_row[row_index].push(cell);
        }
    }

    for row_cells in &mut cells_by_row {
        row_cells.sort_by_key(|cell| (cell.col_addr, cell.col_span, cell.row_span));
    }

    let mut occupied = vec![vec![false; total_cols]; usize::from(row_count)];

    for (r, row_cells) in cells_by_row.into_iter().enumerate() {
        if r >= rows.len() {
            break;
        }
        let mut current_col = 0usize;

        for cell in row_cells {
            let target_col = usize::from(cell.col_addr);
            while current_col < target_col && current_col < total_cols {
                if !occupied[r][current_col] {
                    rows[r].cells.push(TableCell::default());
                }
                current_col += 1;
            }
            while current_col < total_cols && occupied[r][current_col] {
                current_col += 1;
            }

            let placement_col = current_col.max(target_col.min(total_cols));
            let col_span = usize::from(cell.col_span.max(1));
            let row_span = usize::from(cell.row_span.max(1));
            let text = cell
                .blocks
                .iter()
                .filter_map(|b| match b {
                    Block::Paragraph(p) => {
                        Some(p.runs.iter().map(|r| r.text.as_str()).collect::<String>())
                    }
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
                padding_left: if cell.margin_left > 0 {
                    Some(cell.margin_left as i32)
                } else {
                    None
                },
                padding_right: if cell.margin_right > 0 {
                    Some(cell.margin_right as i32)
                } else {
                    None
                },
                padding_top: if cell.margin_top > 0 {
                    Some(cell.margin_top as i32)
                } else {
                    None
                },
                padding_bottom: if cell.margin_bottom > 0 {
                    Some(cell.margin_bottom as i32)
                } else {
                    None
                },
                style: resolve_border_fill_style(doc_info, cell.border_fill_id),
                // HWP cell heights are useful for small cover/layout tables, but some
                // documents also contain page-sized bogus values. Keep only sane hints
                // so title-page layout survives without reintroducing giant black blocks.
                height: (cell.height > 0 && cell.height <= 20_000).then_some(cell.height as i32),
                ..TableCell::default()
            });

            for future_row in r + 1..usize::min(rows.len(), r + row_span) {
                for future_col in placement_col..usize::min(total_cols, placement_col + col_span) {
                    occupied[future_row][future_col] = true;
                }
            }
            current_col = placement_col.saturating_add(col_span);
        }

        while current_col < total_cols {
            if !occupied[r][current_col] {
                rows[r].cells.push(TableCell::default());
            }
            current_col += 1;
        }
    }

    // Remove empty trailing rows
    while rows.last().map_or(false, |r| r.cells.is_empty()) {
        rows.pop();
    }

    distribute_inline_images_across_empty_cells(&mut rows);

    TableBlock {
        rows,
        no_adjust: false,
        page_break_before: false,
        ..TableBlock::default()
    }
}

fn distribute_inline_images_across_empty_cells(rows: &mut [TableRow]) {
    for row in rows {
        if row.cells.len() < 2 {
            continue;
        }

        let populated_indices: Vec<usize> = row
            .cells
            .iter()
            .enumerate()
            .filter_map(|(index, cell)| {
                (!cell.blocks.is_empty() || !cell.text.trim().is_empty()).then_some(index)
            })
            .collect();
        if populated_indices.len() != 1 {
            continue;
        }

        let donor_index = populated_indices[0];
        let donor = &row.cells[donor_index];
        if donor.blocks.len() < 2 || !donor.blocks.iter().all(|block| matches!(block, Block::Image(_)))
        {
            continue;
        }

        let empty_indices: Vec<usize> = row
            .cells
            .iter()
            .enumerate()
            .filter_map(|(index, cell)| {
                (index != donor_index && cell.blocks.is_empty() && cell.text.trim().is_empty())
                    .then_some(index)
            })
            .collect();
        if donor.blocks.len() > empty_indices.len() + 1 {
            continue;
        }

        let donor_width = donor.width;
        let donor_height = donor.height;
        let donor_padding_left = donor.padding_left;
        let donor_padding_right = donor.padding_right;
        let donor_padding_top = donor.padding_top;
        let donor_padding_bottom = donor.padding_bottom;
        let donor_style = donor.style.clone();

        let donor_blocks = std::mem::take(&mut row.cells[donor_index].blocks);
        let Some(first_block) = donor_blocks.get(0).cloned() else {
            continue;
        };
        row.cells[donor_index].blocks = vec![first_block];

        for (target_index, block) in empty_indices.into_iter().zip(donor_blocks.into_iter().skip(1)) {
            let target = &mut row.cells[target_index];
            target.blocks = vec![block];
            if target.width.is_none() {
                target.width = donor_width;
            }
            if target.height.is_none() {
                target.height = donor_height;
            }
            if target.padding_left.is_none() {
                target.padding_left = donor_padding_left;
            }
            if target.padding_right.is_none() {
                target.padding_right = donor_padding_right;
            }
            if target.padding_top.is_none() {
                target.padding_top = donor_padding_top;
            }
            if target.padding_bottom.is_none() {
                target.padding_bottom = donor_padding_bottom;
            }
            if target.style.is_none() {
                target.style = donor_style.clone();
            }
        }
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
        } else if rec.tag_id == HWPTAG_CTRL_HEADER {
            match read_ctrl_type(rec.data).as_deref() {
                Some("gso ") => {
                    let (image, iconsumed) =
                        parse_gso_control(&records[consumed..], doc_info, assets);
                    if let Some(img) = image {
                        hf_blocks.push(Block::Image(img));
                    }
                    consumed += iconsumed.max(1);
                }
                Some("tbl ") => {
                    let (table, tconsumed) =
                        parse_table_control(&records[consumed..], doc_info, assets);
                    if let Some(table) = table {
                        hf_blocks.push(Block::Table(table));
                    }
                    consumed += tconsumed.max(1);
                }
                _ => {
                    consumed += 1;
                }
            }
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

    let obj_attr = u32::from_le_bytes(ctrl.data[4..8].try_into().unwrap_or([0; 4]));
    let treat_as_char = obj_attr & 1 != 0;
    let vert_offset = ctrl
        .data
        .get(8..12)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value != 0);
    let horz_offset = ctrl
        .data
        .get(12..16)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value != 0);
    let width = ctrl
        .data
        .get(16..20)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value > 0);
    let height = ctrl
        .data
        .get(20..24)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value > 0);
    let z_order = ctrl
        .data
        .get(24..28)
        .map(|bytes| i32::from_le_bytes(bytes.try_into().unwrap_or([0; 4])))
        .filter(|value| *value != 0);
    let distance_left = ctrl
        .data
        .get(28..30)
        .map(|bytes| i16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32)
        .filter(|value| *value > 0);
    let distance_right = ctrl
        .data
        .get(30..32)
        .map(|bytes| i16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32)
        .filter(|value| *value > 0);
    let distance_top = ctrl
        .data
        .get(32..34)
        .map(|bytes| i16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32)
        .filter(|value| *value > 0);
    let distance_bottom = ctrl
        .data
        .get(34..36)
        .map(|bytes| i16::from_le_bytes(bytes.try_into().unwrap_or([0; 2])) as i32)
        .filter(|value| *value > 0);
    let vert_rel_to = match (obj_attr >> 3) & 0b11 {
        0 => Some("PAPER".to_string()),
        1 => Some("PAGE".to_string()),
        2 => Some("PARA".to_string()),
        _ => None,
    };
    let vert_align = match (obj_attr >> 5) & 0b111 {
        0 => Some("TOP".to_string()),
        1 => Some("CENTER".to_string()),
        2 => Some("BOTTOM".to_string()),
        3 => Some("INSIDE".to_string()),
        4 => Some("OUTSIDE".to_string()),
        _ => None,
    };
    let horz_rel_to = match (obj_attr >> 8) & 0b11 {
        0 => Some("PAPER".to_string()),
        1 => Some("PAGE".to_string()),
        2 => Some("COLUMN".to_string()),
        3 => Some("PARA".to_string()),
        _ => None,
    };
    let horz_align = match (obj_attr >> 10) & 0b111 {
        0 => Some("LEFT".to_string()),
        1 => Some("CENTER".to_string()),
        2 => Some("RIGHT".to_string()),
        3 => Some("INSIDE".to_string()),
        4 => Some("OUTSIDE".to_string()),
        _ => None,
    };
    let width_rel_to = match (obj_attr >> 15) & 0b111 {
        0 => Some("PAPER".to_string()),
        1 => Some("PAGE".to_string()),
        2 => Some("COLUMN".to_string()),
        3 => Some("PARA".to_string()),
        4 => Some("ABSOLUTE".to_string()),
        _ => None,
    };
    let height_rel_to = match (obj_attr >> 18) & 0b11 {
        0 => Some("PAPER".to_string()),
        1 => Some("PAGE".to_string()),
        2 => Some("ABSOLUTE".to_string()),
        _ => None,
    };
    let text_wrap = match (obj_attr >> 21) & 0b111 {
        0 => Some("SQUARE".to_string()),
        1 => Some("TIGHT".to_string()),
        2 => Some("THROUGH".to_string()),
        3 => Some("TOP_AND_BOTTOM".to_string()),
        4 => Some("BEHIND_TEXT".to_string()),
        5 => Some("IN_FRONT_OF_TEXT".to_string()),
        _ => None,
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

        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE && rec.data.len() >= 73 {
            let candidate = u16::from_le_bytes(rec.data[71..73].try_into().unwrap_or([0; 2]));
            if candidate > 0 && (candidate as usize) <= assets.len() {
                bin_item_id = Some(candidate);
                shape_type = "pic";
            }
        } else if rec.tag_id == HWPTAG_SHAPE_COMPONENT && rec.data.len() >= 4 {
            let candidate_type = rec.data.get(0..4).map(read_ctrl_type_from_reversed);
            if matches!(candidate_type.as_deref(), Some("$pic")) {
                shape_type = "pic";
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
                width_rel_to,
                height_rel_to,
                treat_as_char,
                text_wrap,
                z_order,
                vert_rel_to,
                horz_rel_to,
                vert_align,
                horz_align,
                vert_offset,
                horz_offset,
                distance_left,
                distance_right,
                distance_top,
                distance_bottom,
                page_break_before: false,
                ..ImageBlock::default()
            }),
            consumed,
        );
    }

    (None, consumed)
}

fn read_ctrl_type_from_reversed(bytes: &[u8]) -> String {
    if bytes.len() < 4 {
        return String::new();
    }
    String::from_utf8_lossy(&[bytes[3], bytes[2], bytes[1], bytes[0]]).to_string()
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
                "DocInfo record decoding (FaceName, CharShape, ParaShape, Style, BinData)"
                    .to_string(),
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
                blocks: header_texts
                    .into_iter()
                    .map(|text| {
                        Block::Paragraph(Paragraph {
                            runs: vec![TextRun { text, style: None }],
                            ..Paragraph::default()
                        })
                    })
                    .collect(),
            }]
        };
        let footers = if footer_texts.is_empty() {
            Vec::new()
        } else {
            vec![max_viewer_core::HeaderFooter {
                apply_page_type: Some("BOTH".to_string()),
                blocks: footer_texts
                    .into_iter()
                    .map(|text| {
                        Block::Paragraph(Paragraph {
                            runs: vec![TextRun { text, style: None }],
                            ..Paragraph::default()
                        })
                    })
                    .collect(),
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

        let mut notes =
            vec!["Legacy HWP files are detected via the Compound File signature.".to_string()];

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
            page_break_before: false,
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
    normalize_hwp_text(String::from_utf16_lossy(&units).trim_matches('\u{0}'))
}

fn decode_utf16_be(bytes: &[u8]) -> String {
    let units = bytes
        .chunks_exact(2)
        .map(|chunk| u16::from_be_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();
    normalize_hwp_text(String::from_utf16_lossy(&units).trim_matches('\u{0}'))
}

fn normalize_hwp_text(text: impl AsRef<str>) -> String {
    let text = text.as_ref();
    if !text.chars().any(|ch| matches!(ch, '\u{f53a}')) {
        return text.to_string();
    }

    let mut normalized = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            // HWP text can contain legacy Hanyang PUA syllables directly.
            // pyhwp/hypua2jamo map U+F53A to the historic reading "ᄒᆞᆫ", but
            // current WebKit does not shape that sequence like Hanword here.
            // Collapse it to the rendered modern syllable so the page matches
            // the source document's visible "한글" wording.
            '\u{f53a}' => normalized.push('한'),
            _ => normalized.push(ch),
        }
    }
    normalized
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
    fn normalizes_known_hanyang_pua_text() {
        let bytes = [0x3a, 0xf5, 0x00, 0xae];
        assert_eq!(decode_utf16_le(&bytes), "한글");
    }

    #[test]
    fn parses_para_line_segments_and_height_hint() {
        let mut payload = Vec::new();
        for (chpos, y, height, space_below) in [(0i32, 0i32, 1000i32, 120i32), (12, 1120, 1000, 80)]
        {
            payload.extend_from_slice(&chpos.to_le_bytes());
            payload.extend_from_slice(&y.to_le_bytes());
            payload.extend_from_slice(&height.to_le_bytes());
            payload.extend_from_slice(&1000i32.to_le_bytes()); // height_text
            payload.extend_from_slice(&850i32.to_le_bytes()); // height_baseline
            payload.extend_from_slice(&space_below.to_le_bytes());
            payload.extend_from_slice(&0i32.to_le_bytes()); // x
            payload.extend_from_slice(&4000i32.to_le_bytes()); // width
            payload.extend_from_slice(&0u32.to_le_bytes()); // flags
        }

        let line_segments = parse_para_line_segs(&payload);
        assert_eq!(line_segments.len(), 2);
        assert_eq!(line_segments[1].chpos, 12);
        assert_eq!(paragraph_layout_height_hint(&line_segments), Some(2200));
    }

    #[test]
    fn injects_line_breaks_from_line_segments() {
        let runs = vec![TextRun {
            text: "사업계획서 작성 및 제출 요령".to_string(),
            style: Some(TextStyle::default()),
        }];
        let line_segments = vec![
            HwpLineSegment {
                chpos: 0,
                y: 0,
                height: 1000,
                space_below: 0,
            },
            HwpLineSegment {
                chpos: 8,
                y: 1000,
                height: 1000,
                space_below: 0,
            },
        ];

        let broken = apply_line_seg_breaks_to_runs(runs, &line_segments);
        let joined = broken
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>();
        assert_eq!(joined, "사업계획서 작성\n 및 제출 요령");
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
                assert_eq!(style.font_size, Some(1000));
                assert!(style.bold);
                assert_eq!(style.text_color.as_deref(), Some("#ff0000"));
                assert_eq!(style.underline_color, None);
            }
            _ => panic!("expected paragraph block"),
        }
    }

    #[test]
    fn parses_char_shape_underline_color() {
        let mut cs_data = vec![0u8; 72];
        cs_data[42..46].copy_from_slice(&1100u32.to_le_bytes());
        cs_data[46..50].copy_from_slice(&(1u32 << 2).to_le_bytes());
        cs_data[56..60].copy_from_slice(&0x00ff00u32.to_le_bytes());

        let cs = parse_char_shape(&cs_data);
        let style = char_shape_to_text_style(&cs, &[]);

        assert!(style.underline);
        assert_eq!(style.underline_color.as_deref(), Some("#00ff00"));
    }

    #[test]
    fn ignores_reserved_underline_kind() {
        let mut cs_data = vec![0u8; 72];
        cs_data[46..50].copy_from_slice(&(2u32 << 2).to_le_bytes());
        cs_data[56..60].copy_from_slice(&0x0000ffu32.to_le_bytes());

        let cs = parse_char_shape(&cs_data);
        let style = char_shape_to_text_style(&cs, &[]);

        assert!(!style.underline);
        assert_eq!(style.underline_color, None);
    }

    #[test]
    fn paragraph_alignment_uses_bits_two_to_four() {
        let style = para_shape_to_paragraph_style(&ParaShape {
            attributes: 3u32 << 2,
            ..ParaShape::default()
        });

        assert_eq!(style.align.as_deref(), Some("CENTER"));
    }

    #[test]
    fn parses_page_def_from_section_control_with_real_tag_ids() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();
        make_file_header_stream(&mut compound, 0);

        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();
            write_record(&mut payload, HWPTAG_CTRL_HEADER, 0, b"dces");
            write_record(&mut payload, HWPTAG_LIST_HEADER, 1, &[0u8; 14]);

            let mut page_def = Vec::new();
            page_def.extend_from_slice(&12240i32.to_le_bytes());
            page_def.extend_from_slice(&15840i32.to_le_bytes());
            page_def.extend_from_slice(&1800i32.to_le_bytes());
            page_def.extend_from_slice(&1800i32.to_le_bytes());
            page_def.extend_from_slice(&1440i32.to_le_bytes());
            page_def.extend_from_slice(&1440i32.to_le_bytes());
            page_def.extend_from_slice(&720i32.to_le_bytes());
            page_def.extend_from_slice(&720i32.to_le_bytes());
            page_def.extend_from_slice(&0i32.to_le_bytes());
            page_def.extend_from_slice(&0u32.to_le_bytes());
            write_record(&mut payload, HWPTAG_PAGE_DEF, 1, &page_def);

            let mut para_header = vec![0u8; 22];
            para_header[..4].copy_from_slice(&4u32.to_le_bytes());
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
            .parse_bytes(&bytes, Some("page-def.hwp"))
            .expect("page def should parse");

        let layout = parsed.document.sections[0]
            .page_layout
            .as_ref()
            .expect("section should expose page layout");
        assert_eq!(layout.width, Some(12240));
        assert_eq!(layout.height, Some(15840));
        assert_eq!(layout.margin_left, Some(1800));
        assert_eq!(layout.margin_top, Some(1440));
        assert!(!layout.landscape);
    }

    #[test]
    fn strips_extended_control_payloads_from_para_text() {
        let bytes = [
            0x02, 0x00, 0x64, 0x63, 0x65, 0x73, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00,
            0x21, 0x00, 0x0d, 0x00,
        ];
        let units = extract_text_units(&bytes, None);
        assert_eq!(units, vec![0x21, 0x0d]);
    }

    #[test]
    fn parses_table_cell_descriptor_from_hwp_list_header() {
        let data = [
            0x01, 0x00, 0x00, 0x00, 0x20, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x06, 0x00,
            0x01, 0x00, 0x95, 0xbb, 0x00, 0x00, 0xa9, 0x0a, 0x00, 0x00, 0x8d, 0x00, 0x8d, 0x00,
            0x8d, 0x00, 0x8d, 0x00, 0x22, 0x00, 0x95, 0xbb, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x00,
        ];
        let (
            col,
            row,
            col_span,
            row_span,
            width,
            height,
            margin_left,
            margin_right,
            margin_top,
            margin_bottom,
            border_fill_id,
        ) = parse_list_header_as_cell(&data);
        assert_eq!((col, row, col_span, row_span), (0, 0, 6, 1));
        assert_eq!(width, 48021);
        assert_eq!(height, 2729);
        assert_eq!(
            (margin_left, margin_right, margin_top, margin_bottom),
            (141, 141, 141, 141)
        );
        assert_eq!(border_fill_id, Some(34));
    }

    #[test]
    fn does_not_treat_diagonal_bytes_as_background_fill() {
        let mut data = vec![0u8; 32];
        data[26] = 1;
        data[27] = 0;
        data[28..32].copy_from_slice(&0x0000_0100u32.to_le_bytes());

        let border_fill = parse_border_fill(&data);
        assert_eq!(border_fill.background_color, None);
    }

    #[test]
    fn reads_solid_fill_background_from_fill_stream() {
        let mut data = vec![0u8; 49];
        data[32..36].copy_from_slice(&0x0000_0001u32.to_le_bytes());
        data[36..40].copy_from_slice(&0x00ff_0000u32.to_le_bytes());

        let border_fill = parse_border_fill(&data);
        assert_eq!(border_fill.background_color.as_deref(), Some("#0000ff"));
    }

    #[test]
    fn reads_gradation_fill_as_css_background_image() {
        let mut data = vec![0u8; 70];
        data[32..36].copy_from_slice(&0x0000_0004u32.to_le_bytes());
        data[36] = 2;
        data[37..41].copy_from_slice(&90i32.to_le_bytes());
        data[41..45].copy_from_slice(&0i32.to_le_bytes());
        data[45..49].copy_from_slice(&0i32.to_le_bytes());
        data[49..53].copy_from_slice(&0i32.to_le_bytes());
        data[53..57].copy_from_slice(&2i32.to_le_bytes());
        data[57..61].copy_from_slice(&0x0000_00ffu32.to_le_bytes());
        data[61..65].copy_from_slice(&0x00ff_0000u32.to_le_bytes());
        data[65..69].copy_from_slice(&1u32.to_le_bytes());
        data[69] = 50;

        let border_fill = parse_border_fill(&data);
        assert_eq!(
            border_fill.background_image.as_deref(),
            Some("linear-gradient(90deg, #ff0000, #0000ff)")
        );
    }

    #[test]
    fn parses_border_fill_from_repeated_border_structs() {
        let mut data = vec![0u8; 49];
        data[2] = 8; // left: double
        data[3] = 8; // left width = 0.60mm
        data[4..8].copy_from_slice(&0x0000_00ffu32.to_le_bytes()); // left red

        data[8] = 1; // right: solid
        data[9] = 0; // right width = 0.10mm
        data[10..14].copy_from_slice(&0x0000_ff00u32.to_le_bytes()); // right green

        data[14] = 6; // top: long dash
        data[15] = 10; // top width = 1.00mm
        data[16..20].copy_from_slice(&0x00ff_0000u32.to_le_bytes()); // top blue

        data[20] = 3; // bottom: dot
        data[21] = 3; // bottom width = 0.20mm
        data[22..26].copy_from_slice(&0x00ff_ffffu32.to_le_bytes()); // bottom yellow

        let border_fill = parse_border_fill(&data);
        assert_eq!(
            border_fill
                .border_left
                .as_ref()
                .and_then(|b| b.style.as_deref()),
            Some("DOUBLE")
        );
        assert_eq!(
            border_fill
                .border_left
                .as_ref()
                .and_then(|b| b.width.as_deref()),
            Some("0.60 mm")
        );
        assert_eq!(
            border_fill
                .border_left
                .as_ref()
                .and_then(|b| b.color.as_deref()),
            Some("#ff0000")
        );
        assert_eq!(
            border_fill
                .border_top
                .as_ref()
                .and_then(|b| b.style.as_deref()),
            Some("LONG_DASH")
        );
        assert_eq!(
            border_fill
                .border_bottom
                .as_ref()
                .and_then(|b| b.style.as_deref()),
            Some("DOT")
        );
    }

    #[test]
    fn masks_hwp_border_style_flags_to_low_nibble() {
        let mut data = vec![0u8; 49];
        data[2] = 0x11; // low nibble = 1 => solid
        data[3] = 0x06; // low nibble = 6 => 0.40mm
        data[8] = 0x00; // none
        data[9] = 0x01; // 0.12mm

        let border_fill = parse_border_fill(&data);
        assert_eq!(
            border_fill
                .border_left
                .as_ref()
                .and_then(|b| b.style.as_deref()),
            Some("SOLID")
        );
        assert_eq!(
            border_fill
                .border_left
                .as_ref()
                .and_then(|b| b.width.as_deref()),
            Some("0.40 mm")
        );
        assert_eq!(
            border_fill
                .border_right
                .as_ref()
                .and_then(|b| b.style.as_deref()),
            Some("NONE")
        );
    }

    #[test]
    fn assemble_table_respects_column_addresses() {
        let table = assemble_table(
            2,
            3,
            vec![
                CellDef {
                    col_addr: 1,
                    row_addr: 0,
                    col_span: 1,
                    row_span: 1,
                    width: 1000,
                    height: 0,
                    margin_left: 0,
                    margin_right: 0,
                    margin_top: 0,
                    margin_bottom: 0,
                    border_fill_id: None,
                    blocks: vec![Block::Paragraph(Paragraph {
                        runs: vec![TextRun {
                            text: "middle".to_string(),
                            style: None,
                        }],
                        ..Paragraph::default()
                    })],
                },
                CellDef {
                    col_addr: 0,
                    row_addr: 1,
                    col_span: 1,
                    row_span: 1,
                    width: 1000,
                    height: 0,
                    margin_left: 0,
                    margin_right: 0,
                    margin_top: 0,
                    margin_bottom: 0,
                    border_fill_id: None,
                    blocks: vec![Block::Paragraph(Paragraph {
                        runs: vec![TextRun {
                            text: "left".to_string(),
                            style: None,
                        }],
                        ..Paragraph::default()
                    })],
                },
            ],
            &DocInfoStore::default(),
        );

        assert_eq!(table.rows.len(), 2);
        assert_eq!(table.rows[0].cells.len(), 3);
        assert_eq!(table.rows[0].cells[0].text, "");
        assert_eq!(table.rows[0].cells[1].text, "middle");
        assert_eq!(table.rows[0].cells[2].text, "");
        assert_eq!(table.rows[1].cells.len(), 3);
        assert_eq!(table.rows[1].cells[0].text, "left");
    }

    #[test]
    fn distributes_multiple_inline_images_across_empty_cells() {
        let table = assemble_table(
            1,
            2,
            vec![CellDef {
                col_addr: 0,
                row_addr: 0,
                col_span: 1,
                row_span: 1,
                width: 23809,
                height: 6798,
                margin_left: 510,
                margin_right: 510,
                margin_top: 141,
                margin_bottom: 141,
                border_fill_id: None,
                blocks: vec![
                    Block::Image(ImageBlock {
                        asset_id: "BIN0001.jpg".to_string(),
                        treat_as_char: true,
                        ..ImageBlock::default()
                    }),
                    Block::Image(ImageBlock {
                        asset_id: "BIN0002.jpg".to_string(),
                        treat_as_char: true,
                        ..ImageBlock::default()
                    }),
                ],
            }],
            &DocInfoStore::default(),
        );

        assert_eq!(table.rows.len(), 1);
        assert_eq!(table.rows[0].cells.len(), 2);
        assert_eq!(table.rows[0].cells[0].blocks.len(), 1);
        assert_eq!(table.rows[0].cells[1].blocks.len(), 1);
        let first = match &table.rows[0].cells[0].blocks[0] {
            Block::Image(image) => image.asset_id.as_str(),
            _ => panic!("expected image in first cell"),
        };
        let second = match &table.rows[0].cells[1].blocks[0] {
            Block::Image(image) => image.asset_id.as_str(),
            _ => panic!("expected image in second cell"),
        };
        assert_eq!(first, "BIN0001.jpg");
        assert_eq!(second, "BIN0002.jpg");
    }

    #[test]
    fn parses_gso_picture_dimensions_from_common_object_properties() {
        let mut ctrl = vec![0u8; 46];
        ctrl[..4].copy_from_slice(b" gso");
        ctrl[4..8].copy_from_slice(&0u32.to_le_bytes());
        ctrl[16..20].copy_from_slice(&15153i32.to_le_bytes());
        ctrl[20..24].copy_from_slice(&4345i32.to_le_bytes());
        ctrl[24..28].copy_from_slice(&3i32.to_le_bytes());

        let mut picture = vec![0u8; 91];
        picture[71..73].copy_from_slice(&1u16.to_le_bytes());

        let assets = vec![AssetRef {
            id: "BIN0001.jpg".to_string(),
            media_type: "image/jpeg".to_string(),
            source_path: Some("/BinData/BIN0001.jpg".to_string()),
            data_uri: None,
        }];
        let records = vec![
            Record {
                tag_id: HWPTAG_CTRL_HEADER,
                level: 3,
                data: &ctrl,
            },
            Record {
                tag_id: HWPTAG_SHAPE_COMPONENT_PICTURE,
                level: 4,
                data: &picture,
            },
        ];

        let (image, consumed) = parse_gso_control(&records, &DocInfoStore::default(), &assets);
        assert_eq!(consumed, 2);
        let image = image.expect("picture should parse");
        assert_eq!(image.asset_id, "BIN0001.jpg");
        assert_eq!(image.width, Some(15153));
        assert_eq!(image.height, Some(4345));
        assert_eq!(image.z_order, Some(3));
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

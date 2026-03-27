use std::io::{Cursor, Read, Seek, SeekFrom};

use cfb::CompoundFile;
use flate2::read::{DeflateDecoder, ZlibDecoder};
use max_viewer_core::{
    Block, Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata, FormatInspector,
    FormatSupport, Paragraph, ParseError, Section, TextRun, UnsupportedBlock,
};

pub const CFB_SIGNATURE: [u8; 8] = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
const FILE_HEADER_STREAM: &str = "/FileHeader";
const PREVIEW_TEXT_STREAM: &str = "/PrvText";
const BODY_TEXT_STORAGE: &str = "/BodyText";
const HWPTAG_BEGIN: u16 = 0x010;
const HWPTAG_PARA_HEADER: u16 = HWPTAG_BEGIN + 50;
const HWPTAG_PARA_TEXT: u16 = HWPTAG_BEGIN + 51;

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

impl HwpInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Hwp,
            status: "preview".to_string(),
            implemented: vec![
                "CFB signature probing".to_string(),
                "FileHeader version and attribute decoding".to_string(),
                "BodyText section probing and minimal paragraph reconstruction".to_string(),
                "PrvText Unicode preview extraction".to_string(),
            ],
            planned: vec![
                "DocInfo record decoding".to_string(),
                "style and control interpretation".to_string(),
                "table and drawing object reconstruction".to_string(),
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
        let body_sections = read_body_sections(&mut compound_file, diagnostics.is_encrypted)?;
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
                assets: Vec::new(),
            },
            diagnostics,
        })
    }
}

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
            "This stage uses the official PrvText preview stream for readable fallback output."
                .to_string(),
        ];

        if has_preview {
            notes.push("PrvText preview stream is present.".to_string());
        } else {
            notes.push("PrvText preview stream is missing.".to_string());
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

fn read_body_sections(
    compound_file: &mut CompoundFile<Cursor<&[u8]>>,
    is_encrypted: bool,
) -> Result<Vec<Section>, ParseError> {
    if is_encrypted || !compound_file.exists(BODY_TEXT_STORAGE) {
        return Ok(Vec::new());
    }

    let mut section_paths = collect_entry_paths(compound_file)
        .into_iter()
        .filter(|path| path.starts_with("/BodyText/Section"))
        .collect::<Vec<_>>();

    section_paths.sort_by_key(|path| section_index(path));

    let mut sections = Vec::new();
    for (index, path) in section_paths.iter().enumerate() {
        let bytes = read_stream_bytes(compound_file, path)?;
        let blocks = parse_body_section_blocks(&bytes)?;
        if !blocks.is_empty() {
            sections.push(Section {
                id: index,
                blocks,
                page_layout: None,
                headers: Vec::new(),
                footers: Vec::new(),
                page_start_number: None,
            });
        }
    }

    Ok(sections)
}

fn section_index(path: &str) -> usize {
    path.rsplit('/')
        .next()
        .unwrap_or(path)
        .trim_start_matches("Section")
        .parse::<usize>()
        .unwrap_or(usize::MAX)
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

fn parse_body_section_blocks(bytes: &[u8]) -> Result<Vec<Block>, ParseError> {
    if bytes.is_empty() {
        return Ok(Vec::new());
    }

    if looks_like_record_stream(bytes) {
        if let Ok(blocks) = parse_body_text_records(bytes) {
            return Ok(blocks);
        }
    }

    if let Some(zlib_bytes) = try_decode_zlib(bytes) {
        if let Ok(blocks) = parse_body_text_records(&zlib_bytes) {
            return Ok(blocks);
        }
    }

    if let Some(deflate_bytes) = try_decode_deflate(bytes) {
        if let Ok(blocks) = parse_body_text_records(&deflate_bytes) {
            return Ok(blocks);
        }
    }

    parse_body_text_records(bytes)
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

fn parse_body_text_records(bytes: &[u8]) -> Result<Vec<Block>, ParseError> {
    let mut offset = 0usize;
    let mut blocks = Vec::new();
    let mut current_chars = None::<u32>;

    while offset + 4 <= bytes.len() {
        let header = u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap_or([0; 4]));
        offset += 4;

        let tag_id = (header & 0x3ff) as u16;
        let size_bits = ((header >> 20) & 0xfff) as usize;
        let size = if size_bits == 0x0fff {
            if offset + 4 > bytes.len() {
                return Err(ParseError::InvalidData(
                    "extended HWP record length header is truncated".to_string(),
                ));
            }
            let extended =
                u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap_or([0; 4])) as usize;
            offset += 4;
            extended
        } else {
            size_bits
        };

        if offset + size > bytes.len() {
            return Err(ParseError::InvalidData(
                "HWP record size exceeds remaining stream bytes".to_string(),
            ));
        }

        let data = &bytes[offset..offset + size];
        offset += size;

        match tag_id {
            HWPTAG_PARA_HEADER => {
                if data.len() >= 4 {
                    let raw_chars = u32::from_le_bytes(data[0..4].try_into().unwrap_or([0; 4]));
                    current_chars = Some(raw_chars & 0x7fff_ffff);
                }
            }
            HWPTAG_PARA_TEXT => {
                let paragraph = parse_para_text(data, current_chars)?;
                if !paragraph.runs.is_empty() {
                    blocks.push(Block::Paragraph(paragraph));
                }
                current_chars = None;
            }
            _ => {}
        }
    }

    Ok(blocks)
}

fn parse_para_text(data: &[u8], current_chars: Option<u32>) -> Result<Paragraph, ParseError> {
    if data.len() % 2 != 0 {
        return Err(ParseError::InvalidData(
            "paragraph text record length is not aligned to UTF-16 code units".to_string(),
        ));
    }

    let mut units = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect::<Vec<_>>();

    if let Some(chars) = current_chars {
        let expected = chars as usize;
        if expected > 0 && expected < units.len() {
            units.truncate(expected);
        }
    }

    let mut text = String::new();
    for unit in units {
        match unit {
            0x0009 => text.push('\t'),
            0x000a => text.push('\n'),
            0x000d => {}
            0x0000..=0x001f => {}
            _ => {
                if let Some(ch) = char::from_u32(unit as u32) {
                    text.push(ch);
                }
            }
        }
    }

    Ok(Paragraph {
        marker: None,
        runs: if text.trim().is_empty() {
            Vec::new()
        } else {
            vec![TextRun { text, style: None }]
        },
        style: None,
        page_break_before: false,
    })
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

    String::from_utf8_lossy(bytes).trim_matches('\u{0}').to_string()
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

#[cfg(test)]
mod tests {
    use super::*;
    use flate2::{write::ZlibEncoder, Compression};
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

    #[test]
    fn parses_hwp_preview_text() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();

        {
            let mut stream = compound.create_stream(FILE_HEADER_STREAM).unwrap();
            let mut header = vec![0u8; 32];
            header[..17].copy_from_slice(b"HWP Document File");
            stream.write_all(&header).unwrap();
            stream.write_all(&0x05000300u32.to_le_bytes()).unwrap();
            stream.write_all(&0u32.to_le_bytes()).unwrap();
        }

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
        assert_eq!(parsed.document.metadata.title.as_deref(), Some("fixture.hwp"));
        assert_eq!(parsed.document.sections[0].blocks.len(), 2);
    }

    #[test]
    fn parses_hwp_bodytext_section() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();

        {
            let mut stream = compound.create_stream(FILE_HEADER_STREAM).unwrap();
            let mut header = vec![0u8; 32];
            header[..17].copy_from_slice(b"HWP Document File");
            stream.write_all(&header).unwrap();
            stream.write_all(&0x05000300u32.to_le_bytes()).unwrap();
            stream.write_all(&0u32.to_le_bytes()).unwrap();
        }

        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();
            let mut para_header = Vec::new();
            para_header.extend_from_slice(&3u32.to_le_bytes());
            para_header.extend_from_slice(&0u32.to_le_bytes());
            para_header.extend_from_slice(&0u16.to_le_bytes());
            para_header.push(0);
            para_header.push(0);
            para_header.extend_from_slice(&0u16.to_le_bytes());
            write_record(&mut payload, HWPTAG_PARA_HEADER, 0, &para_header);

            let text_units = "본문".encode_utf16().flat_map(u16::to_le_bytes).collect::<Vec<_>>();
            write_record(&mut payload, HWPTAG_PARA_TEXT, 0, &text_units);

            let mut stream = compound.create_stream("/BodyText/Section0").unwrap();
            stream.write_all(&payload).unwrap();
        }

        let bytes = compound.into_inner().into_inner();
        let parsed = HwpInspector
            .parse_bytes(&bytes, Some("bodytext.hwp"))
            .expect("body text should parse");

        assert_eq!(parsed.document.sections.len(), 1);
        assert_eq!(parsed.document.sections[0].blocks.len(), 1);
        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(paragraph.runs[0].text, "본문");
            }
            _ => panic!("expected paragraph block"),
        }
    }

    #[test]
    fn parses_compressed_hwp_bodytext_section() {
        let mut compound = CompoundFile::create(Cursor::new(Vec::new())).unwrap();

        {
            let mut stream = compound.create_stream(FILE_HEADER_STREAM).unwrap();
            let mut header = vec![0u8; 32];
            header[..17].copy_from_slice(b"HWP Document File");
            stream.write_all(&header).unwrap();
            stream.write_all(&0x05000300u32.to_le_bytes()).unwrap();
            stream.write_all(&(1u32).to_le_bytes()).unwrap();
        }

        {
            compound.create_storage(BODY_TEXT_STORAGE).unwrap();

            let mut payload = Vec::new();
            let mut para_header = Vec::new();
            para_header.extend_from_slice(&2u32.to_le_bytes());
            para_header.extend_from_slice(&0u32.to_le_bytes());
            para_header.extend_from_slice(&0u16.to_le_bytes());
            para_header.push(0);
            para_header.push(0);
            para_header.extend_from_slice(&0u16.to_le_bytes());
            write_record(&mut payload, HWPTAG_PARA_HEADER, 0, &para_header);

            let text_units = "압축".encode_utf16().flat_map(u16::to_le_bytes).collect::<Vec<_>>();
            write_record(&mut payload, HWPTAG_PARA_TEXT, 0, &text_units);

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

        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(paragraph.runs[0].text, "압축");
            }
            _ => panic!("expected paragraph block"),
        }
    }
}

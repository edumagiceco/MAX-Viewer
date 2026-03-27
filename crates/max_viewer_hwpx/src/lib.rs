use std::{
    collections::{BTreeMap, BTreeSet},
    io::{Cursor, Read},
    path::Path,
};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use max_viewer_core::{
    AssetRef, Block, Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata,
    FormatInspector, FormatSupport, HeaderFooter, ImageBlock, PageLayout, Paragraph,
    ParagraphStyle, ParseError, Section, TableBlock, TableCell, TableRow, TextRun, TextStyle,
};
use roxmltree::{Document as XmlDocument, Node};
use zip::read::ZipArchive;

#[derive(Debug, Default)]
pub struct HwpxInspector;

#[derive(Debug, Clone)]
pub struct HwpxParseResult {
    pub document: Document,
    pub diagnostics: DocumentDiagnostics,
}

#[derive(Debug, Clone, Default)]
struct HwpxStyleResolver {
    char_styles: BTreeMap<u32, TextStyle>,
    para_styles: BTreeMap<u32, ParagraphStyle>,
    named_styles: BTreeMap<u32, NamedStyleRef>,
    numberings: BTreeMap<u32, NumberingDefinition>,
    bullets: BTreeMap<u32, BulletDefinition>,
    begin_page: Option<u32>,
}

#[derive(Debug, Clone, Default)]
struct NamedStyleRef {
    para_pr_id: Option<u32>,
    char_pr_id: Option<u32>,
}

#[derive(Debug, Clone, Default)]
struct NumberingDefinition {
    start: u32,
    levels: BTreeMap<u32, ParaHeadDefinition>,
}

#[derive(Debug, Clone, Default)]
struct BulletDefinition {
    bullet_char: Option<String>,
    para_head: Option<ParaHeadDefinition>,
}

#[derive(Debug, Clone, Default)]
struct ParaHeadDefinition {
    level: u32,
    start: Option<u32>,
    pattern: Option<String>,
    num_format: Option<String>,
    align: Option<String>,
    width_adjust: Option<i32>,
    text_offset_type: Option<String>,
    text_offset: Option<i32>,
    char_pr_id_ref: Option<u32>,
}

#[derive(Debug, Default)]
struct SectionParseState {
    numbering_counters: BTreeMap<u32, Vec<u32>>,
}

#[derive(Debug, Clone)]
struct ParagraphMarkerResolution {
    marker: TextRun,
}

#[derive(Debug, Clone, Default)]
struct AssetCollection {
    assets: Vec<AssetRef>,
    lookup: BTreeMap<String, String>,
}

impl HwpxInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Hwpx,
            status: "active".to_string(),
            implemented: vec![
                "ZIP container probing".to_string(),
                "mimetype presence check".to_string(),
                "content.hpf section resolution".to_string(),
                "header.xml metadata extraction".to_string(),
                "header.xml style table loading".to_string(),
                "section and BinData parsing into the shared document model".to_string(),
                "page layout and basic paragraph/text style reconstruction".to_string(),
            ],
            planned: vec![
                "header and footer rendering".to_string(),
                "inline image positioning".to_string(),
                "numbering and bullet fidelity".to_string(),
            ],
        }
    }

    pub fn parse_bytes(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
    ) -> Result<HwpxParseResult, ParseError> {
        let diagnostics = self.inspect_bytes(bytes)?;
        let mut archive = open_archive(bytes)?;

        let content_manifest = read_optional_utf8(&mut archive, "Contents/content.hpf")?;
        let header_xml = read_optional_utf8(&mut archive, "Contents/header.xml")?;

        let section_entries = resolve_section_entries(content_manifest.as_deref(), &mut archive)?;
        let asset_collection = collect_assets(content_manifest.as_deref(), &mut archive)?;
        let metadata = parse_header_metadata(header_xml.as_deref(), fallback_title, section_entries.len());
        let style_resolver = parse_style_resolver(header_xml.as_deref());

        let mut sections = Vec::with_capacity(section_entries.len());
        for (index, entry_name) in section_entries.iter().enumerate() {
            let xml = read_required_utf8(&mut archive, entry_name)?;
            sections.push(parse_section_document(
                &xml,
                &style_resolver,
                &asset_collection.lookup,
                index,
            )?);
        }

        Ok(HwpxParseResult {
            document: Document {
                format: Some(DocumentFormat::Hwpx),
                metadata,
                sections,
                assets: asset_collection.assets,
            },
            diagnostics,
        })
    }
}

impl FormatInspector for HwpxInspector {
    fn format(&self) -> DocumentFormat {
        DocumentFormat::Hwpx
    }

    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError> {
        let mut archive = open_archive(bytes)?;

        let entry_count = archive.len();
        if entry_count == 0 {
            return Err(ParseError::InvalidContainer(
                "empty HWPX container".to_string(),
            ));
        }

        let mut section_count = 0usize;
        let mut asset_count = 0usize;
        let mut is_encrypted = false;
        let mut has_header = false;

        for index in 0..entry_count {
            let file = archive
                .by_index(index)
                .map_err(|error| ParseError::InvalidContainer(error.to_string()))?;
            let name = file.name().to_string();

            if name.starts_with("Contents/section") && name.ends_with(".xml") {
                section_count += 1;
            }
            if name.starts_with("BinData/") && !name.ends_with('/') {
                asset_count += 1;
            }
            if name == "Contents/header.xml" {
                has_header = true;
            }
            if name.contains("encryption") {
                is_encrypted = true;
            }
        }

        let mut notes = vec![
            "HWPX 문서 컨테이너를 정상적으로 열었습니다.".to_string(),
        ];

        if !has_header {
            notes.push("Contents/header.xml is missing.".to_string());
        }

        let mut version_hint = None;
        if let Ok(mut mimetype_file) = archive.by_name("mimetype") {
            let mut mimetype = String::new();
            mimetype_file
                .read_to_string(&mut mimetype)
                .map_err(|error| ParseError::InvalidData(error.to_string()))?;
            let normalized_mimetype = normalize_hwpx_mimetype(&mimetype);
            if !matches!(
                normalized_mimetype.as_str(),
                "application/hwp+zip" | "application/hwpx+zip"
            ) {
                return Err(ParseError::UnsupportedFormat(format!(
                    "unexpected mimetype: {}",
                    mimetype.trim()
                )));
            }
            version_hint = Some(normalized_mimetype);
        } else {
            notes.push("mimetype 항목이 없어 ZIP 구조 기준으로 문서를 해석했습니다.".to_string());
        }

        Ok(DocumentDiagnostics {
            format: DocumentFormat::Hwpx,
            entry_count,
            section_count,
            asset_count,
            is_encrypted,
            version_hint,
            notes,
        })
    }
}

fn normalize_hwpx_mimetype(mimetype: &str) -> String {
    mimetype.trim().to_ascii_lowercase()
}

fn open_archive(bytes: &[u8]) -> Result<ZipArchive<Cursor<&[u8]>>, ParseError> {
    ZipArchive::new(Cursor::new(bytes))
        .map_err(|error| ParseError::InvalidContainer(error.to_string()))
}

fn read_optional_utf8(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    entry_name: &str,
) -> Result<Option<String>, ParseError> {
    let mut file = match archive.by_name(entry_name) {
        Ok(file) => file,
        Err(_) => return Ok(None),
    };

    let mut content = String::new();
    file.read_to_string(&mut content)
        .map_err(|error| ParseError::InvalidData(error.to_string()))?;
    Ok(Some(content))
}

fn read_required_utf8(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    entry_name: &str,
) -> Result<String, ParseError> {
    read_optional_utf8(archive, entry_name)?.ok_or_else(|| {
        ParseError::InvalidData(format!("missing required HWPX entry: {entry_name}"))
    })
}

fn resolve_section_entries(
    content_manifest: Option<&str>,
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> Result<Vec<String>, ParseError> {
    let mut ordered = if let Some(xml) = content_manifest {
        parse_spine_entries(xml)?
    } else {
        Vec::new()
    };

    if ordered.is_empty() {
        ordered = collect_archive_section_entries(archive)?;
    }

    Ok(ordered)
}

fn parse_spine_entries(xml: &str) -> Result<Vec<String>, ParseError> {
    let document =
        XmlDocument::parse(xml).map_err(|error| ParseError::InvalidData(error.to_string()))?;
    let manifest = parse_manifest_entries_from_document(&document);

    let mut ordered = Vec::new();
    let mut seen = BTreeSet::new();

    for node in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "itemref")
    {
        if let Some(idref) = node.attribute("idref") {
            if let Some(entry) = manifest.get(idref) {
                if is_section_entry(entry) && seen.insert(entry.clone()) {
                    ordered.push(entry.clone());
                }
            }
        }
    }

    if ordered.is_empty() {
        for entry in manifest.values() {
            if is_section_entry(entry) && seen.insert(entry.clone()) {
                ordered.push(entry.clone());
            }
        }
    }

    Ok(ordered)
}

fn collect_archive_section_entries(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> Result<Vec<String>, ParseError> {
    let mut sections = Vec::new();

    for index in 0..archive.len() {
        let file = archive
            .by_index(index)
            .map_err(|error| ParseError::InvalidContainer(error.to_string()))?;
        let name = file.name().to_string();
        if is_section_entry(&name) {
            sections.push(name);
        }
    }

    sections.sort_by(section_path_sort_key);
    sections.dedup();

    Ok(sections)
}

fn section_path_sort_key(left: &String, right: &String) -> std::cmp::Ordering {
    section_index(left)
        .cmp(&section_index(right))
        .then_with(|| left.cmp(right))
}

fn section_index(entry: &str) -> usize {
    let stem = Path::new(entry)
        .file_stem()
        .and_then(|stem| stem.to_str())
        .unwrap_or(entry);
    stem.trim_start_matches("section")
        .parse::<usize>()
        .unwrap_or(usize::MAX)
}

fn normalize_entry_path(href: &str) -> String {
    let mut trimmed = href.trim();
    while let Some(rest) = trimmed.strip_prefix("./") {
        trimmed = rest;
    }
    while let Some(rest) = trimmed.strip_prefix("../") {
        trimmed = rest;
    }

    if trimmed.starts_with("Contents/") || trimmed.starts_with("BinData/") {
        trimmed.to_string()
    } else {
        format!("Contents/{trimmed}")
    }
}

fn is_section_entry(entry: &str) -> bool {
    entry.starts_with("Contents/section") && entry.ends_with(".xml")
}

fn collect_assets(
    content_manifest: Option<&str>,
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> Result<AssetCollection, ParseError> {
    let manifest_entries = parse_manifest_entries(content_manifest)?;
    let mut assets = Vec::new();
    let mut lookup = BTreeMap::new();

    for index in 0..archive.len() {
        let mut file = archive
            .by_index(index)
            .map_err(|error| ParseError::InvalidContainer(error.to_string()))?;
        let name = file.name().to_string();

        if name.starts_with("BinData/") && !name.ends_with('/') {
            let media_type = guess_media_type(&name);
            let mut bytes = Vec::new();
            file.read_to_end(&mut bytes)
                .map_err(|error| ParseError::InvalidData(error.to_string()))?;

            let canonical_id = name.clone();
            let data_uri = media_type.starts_with("image/").then(|| {
                format!("data:{media_type};base64,{}", BASE64_STANDARD.encode(bytes))
            });

            register_asset_aliases(&mut lookup, &canonical_id, Some(&name));

            assets.push(AssetRef {
                id: canonical_id,
                media_type,
                source_path: Some(name),
                data_uri,
            });
        }
    }

    let known_ids = assets
        .iter()
        .map(|asset| asset.id.as_str())
        .collect::<BTreeSet<_>>();

    for (manifest_id, manifest_path) in manifest_entries {
        let normalized_path = normalize_entry_path(&manifest_path);
        if known_ids.contains(normalized_path.as_str()) {
            lookup.insert(manifest_id.clone(), normalized_path.clone());
            register_asset_aliases(&mut lookup, &normalized_path, Some(&manifest_id));
        }
    }

    Ok(AssetCollection { assets, lookup })
}

fn parse_manifest_entries(content_manifest: Option<&str>) -> Result<BTreeMap<String, String>, ParseError> {
    let Some(xml) = content_manifest else {
        return Ok(BTreeMap::new());
    };

    let document =
        XmlDocument::parse(xml).map_err(|error| ParseError::InvalidData(error.to_string()))?;
    Ok(parse_manifest_entries_from_document(&document))
}

fn parse_manifest_entries_from_document(document: &XmlDocument<'_>) -> BTreeMap<String, String> {
    let mut manifest = BTreeMap::new();
    for node in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "item")
    {
        if let (Some(id), Some(href)) = (node.attribute("id"), node.attribute("href")) {
            manifest.insert(id.to_string(), normalize_entry_path(href));
        }
    }
    manifest
}

fn register_asset_aliases(
    lookup: &mut BTreeMap<String, String>,
    canonical_id: &str,
    extra_alias: Option<&str>,
) {
    lookup.insert(canonical_id.to_string(), canonical_id.to_string());

    if let Some(path) = Path::new(canonical_id).file_name().and_then(|value| value.to_str()) {
        lookup.insert(path.to_string(), canonical_id.to_string());
    }

    if let Some(stem) = Path::new(canonical_id).file_stem().and_then(|value| value.to_str()) {
        lookup.insert(stem.to_string(), canonical_id.to_string());
    }

    if let Some(alias) = extra_alias.filter(|alias| !alias.trim().is_empty()) {
        lookup.insert(alias.to_string(), canonical_id.to_string());

        if let Some(stem) = Path::new(alias).file_stem().and_then(|value| value.to_str()) {
            lookup.insert(stem.to_string(), canonical_id.to_string());
        }
    }
}

fn guess_media_type(name: &str) -> String {
    match Path::new(name)
        .extension()
        .and_then(|extension| extension.to_str())
        .map(|extension| extension.to_ascii_lowercase())
        .as_deref()
    {
        Some("jpg") | Some("jpeg") => "image/jpeg".to_string(),
        Some("png") => "image/png".to_string(),
        Some("gif") => "image/gif".to_string(),
        Some("bmp") => "image/bmp".to_string(),
        Some("svg") => "image/svg+xml".to_string(),
        Some("webp") => "image/webp".to_string(),
        _ => "application/octet-stream".to_string(),
    }
}

fn parse_header_metadata(
    header_xml: Option<&str>,
    fallback_title: Option<&str>,
    section_count: usize,
) -> DocumentMetadata {
    let mut metadata = DocumentMetadata {
        title: fallback_title.map(ToOwned::to_owned),
        page_count: Some(section_count as u32),
        ..DocumentMetadata::default()
    };

    let Some(xml) = header_xml else {
        return metadata;
    };

    let Ok(document) = XmlDocument::parse(xml) else {
        return metadata;
    };

    if metadata.title.is_none() {
        metadata.title = first_text_for_tags(&document, &["title", "subject"]);
    }
    metadata.author = first_text_for_tags(&document, &["creator", "author"]);
    metadata.language = first_text_for_tags(&document, &["language", "lang"]);

    metadata
}

fn first_text_for_tags(document: &XmlDocument<'_>, tags: &[&str]) -> Option<String> {
    document
        .descendants()
        .filter(|node| node.is_element())
        .find_map(|node| {
            let name = node.tag_name().name();
            if tags.iter().any(|tag| *tag == name) {
                let text = node.text()?.trim();
                if text.is_empty() {
                    None
                } else {
                    Some(text.to_string())
                }
            } else {
                None
            }
        })
}

fn parse_style_resolver(header_xml: Option<&str>) -> HwpxStyleResolver {
    let Some(xml) = header_xml else {
        return HwpxStyleResolver::default();
    };

    let Ok(document) = XmlDocument::parse(xml) else {
        return HwpxStyleResolver::default();
    };

    let mut font_faces = BTreeMap::<String, BTreeMap<u32, String>>::new();
    for fontface in document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "fontface")
    {
        let lang = fontface
            .attribute("lang")
            .map(|value| value.trim().to_ascii_uppercase())
            .unwrap_or_else(|| "HANGUL".to_string());
        let language_fonts = font_faces.entry(lang).or_default();

        for font in fontface
            .children()
            .filter(|node| node.is_element() && node.tag_name().name() == "font")
        {
            let Some(id) = parse_u32_attribute(font, "id") else {
                continue;
            };
            let Some(face) = font.attribute("face").map(str::trim) else {
                continue;
            };
            language_fonts.insert(id, face.to_string());
        }
    }

    let char_styles = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "charPr")
        .filter_map(|node| Some((parse_u32_attribute(node, "id")?, parse_text_style(node, &font_faces))))
        .collect::<BTreeMap<_, _>>();

    let para_styles = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "paraPr")
        .filter_map(|node| {
            Some((
                parse_u32_attribute(node, "id")?,
                parse_paragraph_style_definition(node),
            ))
        })
        .collect::<BTreeMap<_, _>>();

    let named_styles = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "style")
        .filter_map(|node| {
            Some((
                parse_u32_attribute(node, "id")?,
                NamedStyleRef {
                    para_pr_id: parse_u32_attribute(node, "paraPrIDRef"),
                    char_pr_id: parse_u32_attribute(node, "charPrIDRef"),
                },
            ))
        })
        .collect::<BTreeMap<_, _>>();

    let numberings = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "numbering")
        .filter_map(|node| Some((parse_u32_attribute(node, "id")?, parse_numbering_definition(node))))
        .collect::<BTreeMap<_, _>>();

    let bullets = document
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == "bullet")
        .filter_map(|node| Some((parse_u32_attribute(node, "id")?, parse_bullet_definition(node))))
        .collect::<BTreeMap<_, _>>();

    let begin_page = document
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "beginNum")
        .and_then(|node| parse_u32_attribute(node, "page"));

    HwpxStyleResolver {
        char_styles,
        para_styles,
        named_styles,
        numberings,
        bullets,
        begin_page,
    }
}

fn parse_numbering_definition(node: Node<'_, '_>) -> NumberingDefinition {
    let mut levels = BTreeMap::new();

    for para_head in node
        .children()
        .filter(|child| child.is_element() && child.tag_name().name() == "paraHead")
    {
        let definition = parse_para_head_definition(para_head);
        levels.insert(definition.level, definition);
    }

    NumberingDefinition {
        start: parse_u32_attribute(node, "start").unwrap_or(1),
        levels,
    }
}

fn parse_bullet_definition(node: Node<'_, '_>) -> BulletDefinition {
    let para_head = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "paraHead")
        .map(parse_para_head_definition);

    BulletDefinition {
        bullet_char: node.attribute("char").map(|value| value.to_string()),
        para_head,
    }
}

fn parse_para_head_definition(node: Node<'_, '_>) -> ParaHeadDefinition {
    ParaHeadDefinition {
        level: parse_u32_attribute(node, "level").unwrap_or(0),
        start: parse_u32_attribute(node, "start"),
        pattern: node.text().map(str::trim).filter(|text| !text.is_empty()).map(str::to_string),
        num_format: node.attribute("numFormat").map(|value| value.to_string()),
        align: node.attribute("align").map(|value| value.to_string()),
        width_adjust: parse_i32_attribute(node, "widthAdjust"),
        text_offset_type: node.attribute("textOffsetType").map(|value| value.to_string()),
        text_offset: parse_i32_attribute(node, "textOffset"),
        char_pr_id_ref: parse_u32_attribute(node, "charPrIDRef"),
    }
}

fn parse_section_document(
    xml: &str,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    section_id: usize,
) -> Result<Section, ParseError> {
    let document =
        XmlDocument::parse(xml).map_err(|error| ParseError::InvalidData(error.to_string()))?;
    let root = document.root_element();
    let mut state = SectionParseState::default();
    let mut blocks = Vec::new();
    collect_blocks(root, style_resolver, asset_lookup, &mut state, &mut blocks);

    Ok(Section {
        id: section_id,
        blocks,
        page_layout: parse_page_layout(root),
        headers: parse_header_footers(root, "header", style_resolver),
        footers: parse_header_footers(root, "footer", style_resolver),
        page_start_number: style_resolver.begin_page,
    })
}

fn collect_blocks(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
    blocks: &mut Vec<Block>,
) {
    for child in node.children().filter(|node| node.is_element()) {
        match child.tag_name().name() {
            "p" => {
                if let Some(paragraph) = parse_paragraph(child, style_resolver, state) {
                    blocks.push(Block::Paragraph(paragraph));
                }
            }
            "tbl" => {
                let table = parse_table(child, style_resolver, asset_lookup, state);
                if !table.rows.is_empty() {
                    blocks.push(Block::Table(table));
                }
            }
            "pic" | "img" | "image" | "ole" | "rect" | "ellipse" | "line" | "connectLine"
            | "arc" | "curve" | "polygon" | "container" | "textart" | "equation" => {
                if let Some(image) = parse_image(child, asset_lookup) {
                    blocks.push(Block::Image(image));
                }
            }
            "header" | "footer" => {}
            _ => collect_blocks(child, style_resolver, asset_lookup, state, blocks),
        }
    }
}

fn parse_paragraph(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    state: &mut SectionParseState,
) -> Option<Paragraph> {
    let mut paragraph_style = resolve_paragraph_style(node, style_resolver);
    let default_run_style = resolve_default_text_style(node, style_resolver);
    let mut runs = parse_paragraph_runs(node, style_resolver, default_run_style);
    let line_break_positions = parse_line_break_positions(node);
    if !line_break_positions.is_empty() {
        runs = inject_line_breaks_into_runs(runs, &line_break_positions);
    }

    let marker_resolution = resolve_paragraph_marker(paragraph_style.as_mut(), style_resolver, state);
    Some(Paragraph {
        marker: marker_resolution.map(|resolution| resolution.marker),
        runs,
        style: paragraph_style,
        page_break_before: node
            .attribute("pageBreak")
            .map(|value| matches!(value.trim(), "1" | "true" | "TRUE"))
            .unwrap_or(false),
    })
}

fn parse_paragraph_runs(
    paragraph_node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    default_run_style: Option<TextStyle>,
) -> Vec<TextRun> {
    let mut runs = paragraph_node
        .children()
        .filter(|node| node.is_element() && node.tag_name().name() == "run")
        .filter_map(|run_node| {
            let text = collect_text(run_node);
            if text.is_empty() {
                return None;
            }

            let run_style = parse_u32_attribute(run_node, "charPrIDRef")
                .and_then(|id| style_resolver.char_styles.get(&id).cloned())
                .or_else(|| default_run_style.clone());

            Some(TextRun {
                text,
                style: run_style,
            })
        })
        .collect::<Vec<_>>();

    if runs.is_empty() {
        let text = collect_text(paragraph_node);
        if !text.is_empty() {
            runs.push(TextRun {
                text,
                style: default_run_style,
            });
        }
    }

    runs
}

fn parse_line_break_positions(paragraph_node: Node<'_, '_>) -> Vec<usize> {
    let mut positions = paragraph_node
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "linesegarray")
        .into_iter()
        .flat_map(|line_seg_array| {
            line_seg_array
                .children()
                .filter(|child| child.is_element() && child.tag_name().name() == "lineseg")
                .filter_map(|line_seg| parse_u32_attribute(line_seg, "textpos").map(|value| value as usize))
        })
        .collect::<Vec<_>>();

    positions.sort_unstable();
    positions.dedup();
    positions.retain(|position| *position > 0);
    positions
}

fn inject_line_breaks_into_runs(runs: Vec<TextRun>, line_break_positions: &[usize]) -> Vec<TextRun> {
    if runs.is_empty() || line_break_positions.is_empty() {
        return runs;
    }

    let mut normalized_breaks = line_break_positions.iter().copied().peekable();
    let total_chars = runs.iter().map(|run| run.text.chars().count()).sum::<usize>();
    let mut global_char_index = 0usize;
    let mut output = Vec::new();

    for run in runs {
        let characters = run.text.chars().collect::<Vec<_>>();
        let mut current = String::new();

        for ch in characters {
            while let Some(next_break) = normalized_breaks.peek().copied() {
                if next_break == global_char_index {
                    if !current.is_empty() {
                        output.push(TextRun {
                            text: current.clone(),
                            style: run.style.clone(),
                        });
                        current.clear();
                    }
                    output.push(TextRun {
                        text: "\n".to_string(),
                        style: None,
                    });
                    normalized_breaks.next();
                } else {
                    break;
                }
            }

            current.push(ch);
            global_char_index += 1;
        }

        if !current.is_empty() {
            output.push(TextRun {
                text: current,
                style: run.style.clone(),
            });
        }
    }

    while let Some(next_break) = normalized_breaks.peek().copied() {
        if next_break >= total_chars {
            break;
        }
        output.push(TextRun {
            text: "\n".to_string(),
            style: None,
        });
        normalized_breaks.next();
    }

    output
}

fn resolve_paragraph_style(
    paragraph_node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
) -> Option<ParagraphStyle> {
    let direct_para_style_id = parse_u32_attribute(paragraph_node, "paraPrIDRef");
    let named_style = parse_u32_attribute(paragraph_node, "styleIDRef")
        .and_then(|id| style_resolver.named_styles.get(&id));

    direct_para_style_id
        .or_else(|| named_style.and_then(|style| style.para_pr_id))
        .and_then(|id| style_resolver.para_styles.get(&id).cloned())
}

fn resolve_default_text_style(
    paragraph_node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
) -> Option<TextStyle> {
    parse_u32_attribute(paragraph_node, "styleIDRef")
        .and_then(|id| style_resolver.named_styles.get(&id))
        .and_then(|style| style.char_pr_id)
        .and_then(|id| style_resolver.char_styles.get(&id).cloned())
}

fn resolve_paragraph_marker(
    paragraph_style: Option<&mut ParagraphStyle>,
    style_resolver: &HwpxStyleResolver,
    state: &mut SectionParseState,
) -> Option<ParagraphMarkerResolution> {
    let style = paragraph_style?;
    let heading_type = style.heading_type.as_deref()?;
    let heading_id_ref = style.heading_id_ref?;
    let heading_level = style.heading_level.unwrap_or(0);

    match heading_type {
        "NUMBER" | "OUTLINE" => {
            let definition = style_resolver.numberings.get(&heading_id_ref)?;
            let para_head = definition.levels.get(&heading_level)?;
            let value = next_numbering_value(
                &mut state.numbering_counters,
                heading_id_ref,
                heading_level,
                para_head.start.unwrap_or(definition.start.max(1)),
            );
            let marker_text = render_numbering_marker(para_head, value)?;
            style.marker_align = para_head.align.clone();
            style.marker_width_adjust = para_head.width_adjust;
            style.marker_text_offset_type = para_head.text_offset_type.clone();
            style.marker_text_offset = para_head.text_offset;
            Some(ParagraphMarkerResolution {
                marker: TextRun {
                    text: format!("{marker_text} "),
                    style: para_head
                        .char_pr_id_ref
                        .and_then(|id| style_resolver.char_styles.get(&id).cloned()),
                },
            })
        }
        "BULLET" => {
            let definition = style_resolver.bullets.get(&heading_id_ref)?;
            let marker_text = definition
                .bullet_char
                .clone()
                .or_else(|| definition.para_head.as_ref().and_then(|para_head| para_head.pattern.clone()))
                .unwrap_or_else(|| "•".to_string());
            style.marker_align = definition.para_head.as_ref().and_then(|para_head| para_head.align.clone());
            style.marker_width_adjust = definition.para_head.as_ref().and_then(|para_head| para_head.width_adjust);
            style.marker_text_offset_type = definition
                .para_head
                .as_ref()
                .and_then(|para_head| para_head.text_offset_type.clone());
            style.marker_text_offset = definition.para_head.as_ref().and_then(|para_head| para_head.text_offset);
            Some(ParagraphMarkerResolution {
                marker: TextRun {
                    text: format!("{marker_text} "),
                    style: definition
                        .para_head
                        .as_ref()
                        .and_then(|para_head| para_head.char_pr_id_ref)
                        .and_then(|id| style_resolver.char_styles.get(&id).cloned()),
                },
            })
        }
        _ => None,
    }
}

fn next_numbering_value(
    numbering_counters: &mut BTreeMap<u32, Vec<u32>>,
    numbering_id: u32,
    level: u32,
    start: u32,
) -> u32 {
    let counters = numbering_counters.entry(numbering_id).or_default();
    let level_index = level as usize;
    if counters.len() <= level_index {
        counters.resize(level_index + 1, 0);
    }

    for counter in counters.iter_mut().skip(level_index + 1) {
        *counter = 0;
    }

    if counters[level_index] == 0 {
        counters[level_index] = start.max(1);
    } else {
        counters[level_index] += 1;
    }

    counters[level_index]
}

fn render_numbering_marker(para_head: &ParaHeadDefinition, value: u32) -> Option<String> {
    let marker_value = format_number(value, para_head.num_format.as_deref());
    let pattern = para_head.pattern.as_deref().unwrap_or("^1");

    if pattern.contains('^') {
        Some(pattern.replace("^1", &marker_value))
    } else if pattern.is_empty() {
        Some(marker_value)
    } else {
        Some(format!("{pattern}{marker_value}"))
    }
}

fn format_number(value: u32, format: Option<&str>) -> String {
    match format.unwrap_or("DIGIT") {
        "ROMAN_CAPITAL" => to_roman(value).to_uppercase(),
        "ROMAN_SMALL" => to_roman(value).to_lowercase(),
        "LATIN_CAPITAL" => to_latin(value, true),
        "LATIN_SMALL" => to_latin(value, false),
        "CIRCLED_DIGIT" => to_circled_digit(value),
        "HANGUL_JAMO" => to_hangul_jamo(value),
        "HANGUL_SYLLABLE" => to_hangul_syllable(value),
        _ => value.to_string(),
    }
}

fn to_roman(mut value: u32) -> String {
    let numerals = [
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ];
    let mut output = String::new();
    for (number, symbol) in numerals {
        while value >= number {
            value -= number;
            output.push_str(symbol);
        }
    }
    output
}

fn to_latin(value: u32, upper: bool) -> String {
    let mut number = value.max(1);
    let mut chars = Vec::new();
    while number > 0 {
        number -= 1;
        let base = if upper { b'A' } else { b'a' };
        chars.push((base + (number % 26) as u8) as char);
        number /= 26;
    }
    chars.into_iter().rev().collect()
}

fn to_circled_digit(value: u32) -> String {
    const CIRCLED: [&str; 20] = [
        "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
        "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳",
    ];
    if let Some(circled) = CIRCLED.get(value.saturating_sub(1) as usize) {
        (*circled).to_string()
    } else {
        value.to_string()
    }
}

fn to_hangul_jamo(value: u32) -> String {
    const JAMO: [&str; 14] = ["ㄱ", "ㄴ", "ㄷ", "ㄹ", "ㅁ", "ㅂ", "ㅅ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"];
    if let Some(jamo) = JAMO.get(value.saturating_sub(1) as usize) {
        (*jamo).to_string()
    } else {
        value.to_string()
    }
}

fn to_hangul_syllable(value: u32) -> String {
    const SYLLABLE: [&str; 14] = ["가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하"];
    if let Some(syllable) = SYLLABLE.get(value.saturating_sub(1) as usize) {
        (*syllable).to_string()
    } else {
        value.to_string()
    }
}

fn parse_header_footers(
    section_root: Node<'_, '_>,
    tag_name: &str,
    style_resolver: &HwpxStyleResolver,
) -> Vec<HeaderFooter> {
    section_root
        .descendants()
        .filter(|node| node.is_element() && node.tag_name().name() == tag_name)
        .filter_map(|node| {
            let sub_list = node
                .children()
                .find(|child| child.is_element() && child.tag_name().name() == "subList")?;

            let mut state = SectionParseState::default();
            let mut blocks = Vec::new();
            collect_blocks(sub_list, style_resolver, &BTreeMap::new(), &mut state, &mut blocks);

            Some(HeaderFooter {
                apply_page_type: node.attribute("applyPageType").map(|value| value.to_string()),
                blocks,
            })
        })
        .collect()
}

fn parse_page_layout(section_root: Node<'_, '_>) -> Option<PageLayout> {
    let sec_pr = section_root
        .descendants()
        .find(|node| node.is_element() && node.tag_name().name() == "secPr")?;
    let page_pr = sec_pr
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "pagePr")?;
    let margin = page_pr
        .children()
        .find(|node| node.is_element() && node.tag_name().name() == "margin");

    let mut width = parse_i32_attribute(page_pr, "width");
    let mut height = parse_i32_attribute(page_pr, "height");
    let landscape = matches!(
        page_pr.attribute("landscape").map(|value| value.trim()),
        Some("WIDELY")
    );

    if landscape {
        if let (Some(current_width), Some(current_height)) = (width, height) {
            if current_width < current_height {
                width = Some(current_height);
                height = Some(current_width);
            }
        }
    }

    Some(PageLayout {
        width,
        height,
        landscape,
        margin_left: margin.and_then(|node| parse_i32_attribute(node, "left")),
        margin_right: margin.and_then(|node| parse_i32_attribute(node, "right")),
        margin_top: margin.and_then(|node| parse_i32_attribute(node, "top")),
        margin_bottom: margin.and_then(|node| parse_i32_attribute(node, "bottom")),
        margin_header: margin.and_then(|node| parse_i32_attribute(node, "header")),
        margin_footer: margin.and_then(|node| parse_i32_attribute(node, "footer")),
        margin_gutter: margin.and_then(|node| parse_i32_attribute(node, "gutter")),
    })
}

fn parse_text_style(
    node: Node<'_, '_>,
    font_faces: &BTreeMap<String, BTreeMap<u32, String>>,
) -> TextStyle {
    let font_ref = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "fontRef");

    let font_family = font_ref
        .and_then(|font_ref| {
            [
                ("hangul", "HANGUL"),
                ("latin", "LATIN"),
                ("other", "OTHER"),
                ("hanja", "HANJA"),
                ("japanese", "JAPANESE"),
                ("symbol", "SYMBOL"),
                ("user", "USER"),
            ]
            .into_iter()
            .find_map(|(attribute_name, language_key)| {
                let font_id = parse_u32_attribute(font_ref, attribute_name)?;
                let face = font_faces
                    .get(language_key)
                    .and_then(|faces| faces.get(&font_id))
                    .cloned();
                face
            })
        })
        .or_else(|| {
            font_faces
                .values()
                .find_map(|faces| faces.values().next().cloned())
        });

    let underline = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "underline")
        .map(|underline| underline.attribute("type").map(|value| value != "NONE").unwrap_or(true))
        .unwrap_or(false);

    let ratio = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "ratio")
        .and_then(parse_language_i32_attribute);
    let spacing = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "spacing")
        .and_then(parse_language_i32_attribute);
    let relative_size = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "relSz")
        .and_then(parse_language_i32_attribute);
    let baseline_offset = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "offset")
        .and_then(parse_language_i32_attribute);

    TextStyle {
        font_family,
        font_size: parse_i32_attribute(node, "height"),
        text_color: node
            .attribute("textColor")
            .and_then(normalize_color_value),
        background_color: node
            .attribute("shadeColor")
            .and_then(normalize_color_value),
        width_ratio: ratio,
        letter_spacing: spacing,
        relative_size,
        baseline_offset,
        use_font_space: parse_bool_attribute(node, "useFontSpace"),
        use_kerning: parse_bool_attribute(node, "useKerning"),
        bold: node
            .children()
            .any(|child| child.is_element() && child.tag_name().name() == "bold"),
        italic: node
            .children()
            .any(|child| child.is_element() && child.tag_name().name() == "italic"),
        underline,
    }
}

fn parse_paragraph_style_definition(node: Node<'_, '_>) -> ParagraphStyle {
    let align = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "align")
        .and_then(|align| align.attribute("horizontal"))
        .map(|value| value.to_string());

    let margin = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "margin");

    let line_spacing = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "lineSpacing");

    let heading = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "heading");

    ParagraphStyle {
        align,
        indent: margin.and_then(|node| parse_hwp_value_child(node, "intent")),
        margin_left: margin.and_then(|node| parse_hwp_value_child(node, "left")),
        margin_right: margin.and_then(|node| parse_hwp_value_child(node, "right")),
        margin_prev: margin.and_then(|node| parse_hwp_value_child(node, "prev")),
        margin_next: margin.and_then(|node| parse_hwp_value_child(node, "next")),
        line_spacing_type: line_spacing
            .and_then(|node| node.attribute("type"))
            .map(|value| value.to_string()),
        line_spacing: line_spacing.and_then(|node| parse_i32_attribute(node, "value")),
        heading_type: heading
            .and_then(|node| node.attribute("type"))
            .map(|value| value.to_string()),
        heading_id_ref: heading.and_then(|node| parse_u32_attribute(node, "idRef")),
        heading_level: heading.and_then(|node| parse_u32_attribute(node, "level")),
        marker_align: None,
        marker_width_adjust: None,
        marker_text_offset_type: None,
        marker_text_offset: None,
    }
}

fn parse_hwp_value_child(node: Node<'_, '_>, child_name: &str) -> Option<i32> {
    node.children()
        .find(|child| child.is_element() && child.tag_name().name() == child_name)
        .and_then(|child| parse_i32_attribute(child, "value"))
}

fn parse_language_i32_attribute(node: Node<'_, '_>) -> Option<i32> {
    [
        "hangul",
        "latin",
        "other",
        "hanja",
        "japanese",
        "symbol",
        "user",
    ]
    .into_iter()
    .find_map(|attribute_name| parse_i32_attribute(node, attribute_name))
}

fn parse_u32_attribute(node: Node<'_, '_>, attribute_name: &str) -> Option<u32> {
    node.attribute(attribute_name)?.trim().parse().ok()
}

fn parse_i32_attribute(node: Node<'_, '_>, attribute_name: &str) -> Option<i32> {
    node.attribute(attribute_name)?.trim().parse().ok()
}

fn parse_bool_attribute(node: Node<'_, '_>, attribute_name: &str) -> bool {
    matches!(
        node.attribute(attribute_name).map(|value| value.trim()),
        Some("1" | "true" | "TRUE" | "True")
    )
}

fn normalize_color_value(raw: &str) -> Option<String> {
    let trimmed = raw.trim();
    if trimmed.is_empty() {
        return None;
    }

    if trimmed.starts_with('#') {
        return Some(trimmed.to_string());
    }

    if let Some(hex) = trimmed.strip_prefix("0x").or_else(|| trimmed.strip_prefix("0X")) {
        let parsed = u32::from_str_radix(hex, 16).ok()?;
        return Some(format!("#{:06X}", parsed & 0x00ff_ffff));
    }

    if let Ok(parsed) = trimmed.parse::<u32>() {
        return Some(format!("#{:06X}", parsed & 0x00ff_ffff));
    }

    None
}

fn parse_table(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
) -> TableBlock {
    let mut rows = Vec::new();
    collect_rows(node, style_resolver, asset_lookup, state, &mut rows);
    let (width, height, _, _) = parse_size_attributes(node);

    TableBlock {
        rows,
        width,
        height,
        cell_spacing: parse_i32_attribute(node, "cellSpacing"),
        repeat_header: parse_bool_attribute(node, "repeatHeader"),
    }
}

fn collect_rows(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
    rows: &mut Vec<TableRow>,
) {
    for child in node.children().filter(|node| node.is_element()) {
        if child.tag_name().name() == "tr" {
            rows.push(parse_row(child, style_resolver, asset_lookup, state));
        } else {
            collect_rows(child, style_resolver, asset_lookup, state, rows);
        }
    }
}

fn parse_row(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
) -> TableRow {
    let mut cells = Vec::new();
    collect_cells(node, style_resolver, asset_lookup, state, &mut cells);
    TableRow { cells }
}

fn collect_cells(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
    cells: &mut Vec<TableCell>,
) {
    for child in node.children().filter(|node| node.is_element()) {
        if child.tag_name().name() == "tc" {
            cells.push(parse_cell(child, style_resolver, asset_lookup, state));
        } else {
            collect_cells(child, style_resolver, asset_lookup, state, cells);
        }
    }
}

fn parse_cell(
    node: Node<'_, '_>,
    style_resolver: &HwpxStyleResolver,
    asset_lookup: &BTreeMap<String, String>,
    state: &mut SectionParseState,
) -> TableCell {
    let mut blocks = Vec::new();
    if let Some(sub_list) = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "subList")
    {
        collect_blocks(sub_list, style_resolver, asset_lookup, state, &mut blocks);
    } else {
        collect_blocks(node, style_resolver, asset_lookup, state, &mut blocks);
    }

    let cell_span = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "cellSpan");
    let cell_size = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "cellSz");
    let cell_margin = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "cellMargin");

    TableCell {
        text: blocks_to_plain_text(&blocks).unwrap_or_else(|| collect_text(node)),
        blocks,
        col_span: cell_span.and_then(|span| parse_u32_attribute(span, "colSpan")),
        row_span: cell_span.and_then(|span| parse_u32_attribute(span, "rowSpan")),
        width: cell_size.and_then(|size| parse_i32_attribute(size, "width")),
        height: cell_size.and_then(|size| parse_i32_attribute(size, "height")),
        padding_left: cell_margin.and_then(|margin| parse_i32_attribute(margin, "left")),
        padding_right: cell_margin.and_then(|margin| parse_i32_attribute(margin, "right")),
        padding_top: cell_margin.and_then(|margin| parse_i32_attribute(margin, "top")),
        padding_bottom: cell_margin.and_then(|margin| parse_i32_attribute(margin, "bottom")),
        is_header: parse_bool_attribute(node, "header"),
    }
}

fn parse_image(
    node: Node<'_, '_>,
    asset_lookup: &BTreeMap<String, String>,
) -> Option<ImageBlock> {
    let raw_asset_id = extract_asset_reference(node);

    let asset_id = raw_asset_id
        .as_deref()
        .and_then(|raw| resolve_asset_reference(raw, asset_lookup))
        .unwrap_or_else(|| synthetic_object_id(node));
    let (width, height, width_rel_to, height_rel_to) = parse_size_attributes(node);
    let position = parse_position_attributes(node);
    let caption = parse_caption_text(node);
    let kind = node.tag_name().name().to_string();

    Some(ImageBlock {
        kind: kind.clone(),
        asset_id,
        alt_text: caption
            .clone()
            .or_else(|| Some(format!("{} object", kind.to_uppercase()))),
        width,
        height,
        width_rel_to,
        height_rel_to,
        treat_as_char: position
            .as_ref()
            .map(|position| position.treat_as_char)
            .unwrap_or(false),
        text_wrap: node.attribute("textWrap").map(|value| value.to_string()),
        z_order: parse_i32_attribute(node, "zOrder"),
        vert_rel_to: position.as_ref().and_then(|position| position.vert_rel_to.clone()),
        horz_rel_to: position.as_ref().and_then(|position| position.horz_rel_to.clone()),
        vert_align: position.as_ref().and_then(|position| position.vert_align.clone()),
        horz_align: position.as_ref().and_then(|position| position.horz_align.clone()),
        vert_offset: position.as_ref().and_then(|position| position.vert_offset),
        horz_offset: position.as_ref().and_then(|position| position.horz_offset),
        caption,
    })
}

#[derive(Debug, Clone, Default)]
struct ParsedPosition {
    treat_as_char: bool,
    vert_rel_to: Option<String>,
    horz_rel_to: Option<String>,
    vert_align: Option<String>,
    horz_align: Option<String>,
    vert_offset: Option<i32>,
    horz_offset: Option<i32>,
}

fn parse_size_attributes(
    node: Node<'_, '_>,
) -> (Option<i32>, Option<i32>, Option<String>, Option<String>) {
    let size_node = node
        .children()
        .find(|child| {
            child.is_element() && matches!(child.tag_name().name(), "sz" | "curSz" | "orgSz")
        });

    (
        size_node.and_then(|size| parse_i32_attribute(size, "width")),
        size_node.and_then(|size| parse_i32_attribute(size, "height")),
        size_node.and_then(|size| size.attribute("widthRelTo").map(|value| value.to_string())),
        size_node.and_then(|size| size.attribute("heightRelTo").map(|value| value.to_string())),
    )
}

fn parse_position_attributes(node: Node<'_, '_>) -> Option<ParsedPosition> {
    let position_node = node
        .children()
        .find(|child| child.is_element() && child.tag_name().name() == "pos")?;

    Some(ParsedPosition {
        treat_as_char: parse_bool_attribute(position_node, "treatAsChar"),
        vert_rel_to: position_node.attribute("vertRelTo").map(|value| value.to_string()),
        horz_rel_to: position_node.attribute("horzRelTo").map(|value| value.to_string()),
        vert_align: position_node.attribute("vertAlign").map(|value| value.to_string()),
        horz_align: position_node.attribute("horzAlign").map(|value| value.to_string()),
        vert_offset: parse_i32_attribute(position_node, "vertOffset"),
        horz_offset: parse_i32_attribute(position_node, "horzOffset"),
    })
}

fn parse_caption_text(node: Node<'_, '_>) -> Option<String> {
    node.children()
        .find(|child| child.is_element() && child.tag_name().name() == "caption")
        .map(collect_text)
        .map(|text| text.trim().to_string())
        .filter(|text| !text.is_empty())
}

fn resolve_asset_reference(raw_asset_id: &str, asset_lookup: &BTreeMap<String, String>) -> Option<String> {
    let trimmed = raw_asset_id.trim();
    if trimmed.is_empty() {
        return None;
    }

    asset_lookup
        .get(trimmed)
        .cloned()
        .or_else(|| asset_lookup.get(&normalize_entry_path(trimmed)).cloned())
        .or_else(|| {
            Path::new(trimmed)
                .file_name()
                .and_then(|value| value.to_str())
                .and_then(|value| asset_lookup.get(value).cloned())
        })
        .or_else(|| {
            Path::new(trimmed)
                .file_stem()
                .and_then(|value| value.to_str())
                .and_then(|value| asset_lookup.get(value).cloned())
        })
}

fn synthetic_object_id(node: Node<'_, '_>) -> String {
    let object_id = node.attribute("id").unwrap_or("0");
    format!("{}:{object_id}", node.tag_name().name())
}

fn extract_asset_reference(node: Node<'_, '_>) -> Option<String> {
    const STRONG_PRIORITY: [&str; 4] = ["binaryItemIDRef", "href", "src", "idref"];

    for attribute_name in STRONG_PRIORITY {
        if let Some(value) = node.attribute(attribute_name).map(str::trim).filter(|value| !value.is_empty()) {
            return Some(value.to_string());
        }
    }

    if let Some(value) = node
        .descendants()
        .filter(|child| child.is_element())
        .find_map(|child| {
            STRONG_PRIORITY.into_iter().find_map(|attribute_name| {
                child.attribute(attribute_name)
                    .map(str::trim)
                    .filter(|value| !value.is_empty())
                    .map(str::to_string)
            })
        })
    {
        return Some(value);
    }

    node.attribute("id")
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .map(str::to_string)
        .or_else(|| {
            node.descendants()
                .filter(|child| child.is_element())
                .find_map(|child| {
                    child.attribute("id")
                        .map(str::trim)
                        .filter(|value| !value.is_empty())
                        .map(str::to_string)
                })
        })
}

fn blocks_to_plain_text(blocks: &[Block]) -> Option<String> {
    let mut output = String::new();

    for block in blocks {
        match block {
            Block::Paragraph(paragraph) => {
                if let Some(marker) = &paragraph.marker {
                    output.push_str(&marker.text);
                }
                for run in &paragraph.runs {
                    output.push_str(&run.text);
                }
                output.push('\n');
            }
            Block::Table(table) => {
                for row in &table.rows {
                    let row_text = row
                        .cells
                        .iter()
                        .map(|cell| {
                            blocks_to_plain_text(&cell.blocks)
                                .unwrap_or_else(|| cell.text.clone())
                                .replace('\n', " ")
                                .trim()
                                .to_string()
                        })
                        .collect::<Vec<_>>()
                        .join(" | ");
                    output.push_str(row_text.trim_end());
                    output.push('\n');
                }
            }
            Block::Image(image) => {
                if let Some(caption) = &image.caption {
                    output.push_str(caption);
                } else if let Some(alt_text) = &image.alt_text {
                    output.push_str(alt_text);
                } else {
                    output.push_str("[object]");
                }
                output.push('\n');
            }
            Block::Unsupported(unsupported) => {
                output.push_str(&unsupported.kind);
                output.push('\n');
            }
        }
    }

    let normalized = output.trim().to_string();
    (!normalized.is_empty()).then_some(normalized)
}

fn collect_text(node: Node<'_, '_>) -> String {
    let mut output = String::new();
    collect_text_parts(node, &mut output);
    output
}

fn collect_text_parts(node: Node<'_, '_>, output: &mut String) {
    for child in node.children() {
        if child.is_text() {
            if matches!(node.tag_name().name(), "t" | "text") {
                if let Some(text) = child.text() {
                    output.push_str(text);
                }
            }
            continue;
        }

        if !child.is_element() {
            continue;
        }

        match child.tag_name().name() {
            "tab" => output.push('\t'),
            "lineBreak" | "br" => output.push('\n'),
            "pageNum" => output.push_str(&page_placeholder_for(child)),
            "autoNum" => output.push_str(&auto_num_placeholder_for(child)),
            _ => collect_text_parts(child, output),
        }
    }
}

fn page_placeholder_for(node: Node<'_, '_>) -> String {
    let format = node.attribute("formatType").unwrap_or("DIGIT");
    let side_char = node.attribute("sideChar").unwrap_or("");
    format!("{{{{PAGE:{format}:{side_char}}}}}")
}

fn auto_num_placeholder_for(node: Node<'_, '_>) -> String {
    match node.attribute("numType").unwrap_or("") {
        "PAGE" => {
            let format = node
                .children()
                .find(|child| child.is_element() && child.tag_name().name() == "autoNumFormat")
                .and_then(|child| child.attribute("type"))
                .unwrap_or("DIGIT");
            format!("{{{{PAGE:{format}:}}}}")
        }
        "TOTAL_PAGE" => {
            let format = node
                .children()
                .find(|child| child.is_element() && child.tag_name().name() == "autoNumFormat")
                .and_then(|child| child.attribute("type"))
                .unwrap_or("DIGIT");
            format!("{{{{TOTAL_PAGES:{format}:}}}}")
        }
        _ => String::new(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::{Cursor, Write};
    use zip::write::SimpleFileOptions;
    use zip::ZipWriter;

    fn fixture_hwpx_bytes_with_section(mimetype: &str, section_xml: &[u8]) -> Vec<u8> {
        fixture_hwpx_bytes_with_entries(
            mimetype,
            br#"<opf:package xmlns:opf="http://www.idpf.org/2007/opf">
<opf:manifest>
  <opf:item id="sec0" href="section0.xml" />
</opf:manifest>
<opf:spine>
  <opf:itemref idref="sec0" />
</opf:spine>
</opf:package>"#,
            section_xml,
            &[],
        )
    }

    fn fixture_hwpx_bytes_with_entries(
        mimetype: &str,
        content_hpf: &[u8],
        section_xml: &[u8],
        extra_entries: &[(&str, &[u8])],
    ) -> Vec<u8> {
        let mut buffer = Cursor::new(Vec::new());
        let mut writer = ZipWriter::new(&mut buffer);
        let options = SimpleFileOptions::default();

        writer.start_file("mimetype", options).unwrap();
        writer.write_all(mimetype.as_bytes()).unwrap();

        writer.start_file("Contents/content.hpf", options).unwrap();
        writer.write_all(content_hpf).unwrap();

        writer.start_file("Contents/header.xml", options).unwrap();
        writer
            .write_all(
                r##"<hh:header xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
<hh:title>Fixture Document</hh:title>
<hh:creator>MAX Viewer Test</hh:creator>
<hh:refList>
  <hh:beginNum page="3" />
  <hh:fontfaces>
    <hh:fontface lang="HANGUL" fontCnt="1">
      <hh:font id="0" face="Malgun Gothic" type="TTF" isEmbedded="0" />
    </hh:fontface>
  </hh:fontfaces>
  <hh:charProperties>
    <hh:charPr id="7" height="1200" textColor="#6182D6">
      <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0" />
      <hh:bold />
    </hh:charPr>
    <hh:charPr id="8" height="1000" textColor="#111827">
      <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0" />
      <hh:italic />
    </hh:charPr>
  </hh:charProperties>
  <hh:paraProperties>
    <hh:paraPr id="20">
      <hh:align horizontal="CENTER" vertical="BASELINE" />
      <hh:margin>
        <hc:intent value="0" unit="HWPUNIT" />
        <hc:left value="0" unit="HWPUNIT" />
        <hc:right value="0" unit="HWPUNIT" />
        <hc:prev value="600" unit="HWPUNIT" />
        <hc:next value="600" unit="HWPUNIT" />
      </hh:margin>
      <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT" />
    </hh:paraPr>
    <hh:paraPr id="21">
      <hh:align horizontal="LEFT" vertical="BASELINE" />
      <hh:heading type="NUMBER" idRef="1" level="0" />
    </hh:paraPr>
    <hh:paraPr id="22">
      <hh:align horizontal="LEFT" vertical="BASELINE" />
      <hh:heading type="BULLET" idRef="2" level="0" />
    </hh:paraPr>
  </hh:paraProperties>
  <hh:numberings itemCnt="1">
    <hh:numbering id="1" start="1">
      <hh:paraHead level="0" align="LEFT" start="1" numFormat="DIGIT" charPrIDRef="7">^1.</hh:paraHead>
    </hh:numbering>
  </hh:numberings>
  <hh:bullets itemCnt="1">
    <hh:bullet id="2" char="•">
      <hh:paraHead align="LEFT" charPrIDRef="8">•</hh:paraHead>
    </hh:bullet>
  </hh:bullets>
</hh:refList>
</hh:header>"##
                    .as_bytes(),
            )
            .unwrap();

        writer.start_file("Contents/section0.xml", options).unwrap();
        writer.write_all(section_xml).unwrap();

        for (entry_name, entry_bytes) in extra_entries {
            writer.start_file(entry_name, options).unwrap();
            writer.write_all(entry_bytes).unwrap();
        }

        writer.finish().unwrap();
        buffer.into_inner()
    }

    fn fixture_hwpx_bytes(mimetype: &str) -> Vec<u8> {
        fixture_hwpx_bytes_with_section(
            mimetype,
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:p paraPrIDRef="20">
  <hp:run charPrIDRef="7">
    <hp:secPr>
      <hp:pagePr width="59528" height="84188" landscape="NARROWLY">
        <hp:margin left="8504" right="8504" top="5668" bottom="4252" header="4252" footer="4252" gutter="0" />
      </hp:pagePr>
    </hp:secPr>
    <hp:t>Hello </hp:t>
  </hp:run>
  <hp:run charPrIDRef="8"><hp:t>MAX-Viewer</hp:t></hp:run>
</hp:p>
<hp:tbl>
  <hp:tr>
    <hp:tc><hp:p><hp:run><hp:t>A1</hp:t></hp:run></hp:p></hp:tc>
    <hp:tc><hp:p><hp:run><hp:t>B1</hp:t></hp:run></hp:p></hp:tc>
  </hp:tr>
</hp:tbl>
</hp:section>"#,
        )
    }

    #[test]
    fn parses_basic_hwpx_document() {
        let bytes = fixture_hwpx_bytes("application/hwpx+zip");

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        assert_eq!(parsed.document.sections.len(), 1);
        assert_eq!(parsed.document.metadata.title.as_deref(), Some("fixture.hwpx"));
        assert_eq!(parsed.document.sections[0].page_layout.as_ref().and_then(|layout| layout.width), Some(59528));
        assert!(matches!(
            parsed.document.sections[0].blocks[0],
            Block::Paragraph(_)
        ));
        assert!(matches!(
            parsed.document.sections[0].blocks[1],
            Block::Table(_)
        ));
        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(paragraph.style.as_ref().and_then(|style| style.align.as_deref()), Some("CENTER"));
                assert_eq!(paragraph.runs.len(), 2);
                assert_eq!(
                    paragraph
                        .runs
                        .iter()
                        .map(|run| run.text.as_str())
                        .collect::<String>(),
                    "Hello MAX-Viewer"
                );
                assert_eq!(paragraph.runs[0].style.as_ref().and_then(|style| style.font_family.as_deref()), Some("Malgun Gothic"));
                assert_eq!(paragraph.runs[0].style.as_ref().and_then(|style| style.text_color.as_deref()), Some("#6182D6"));
                assert!(paragraph.runs[0].style.as_ref().map(|style| style.bold).unwrap_or(false));
                assert!(paragraph.runs[1].style.as_ref().map(|style| style.italic).unwrap_or(false));
            }
            _ => panic!("expected paragraph block"),
        }
    }

    #[test]
    fn preserves_empty_paragraphs_and_whitespace_runs() {
        let bytes = fixture_hwpx_bytes_with_section(
            "application/hwpx+zip",
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:p paraPrIDRef="20">
  <hp:run charPrIDRef="7">
    <hp:secPr>
      <hp:pagePr width="59528" height="84188" landscape="NARROWLY">
        <hp:margin left="8504" right="8504" top="5668" bottom="4252" header="4252" footer="4252" gutter="0" />
      </hp:pagePr>
    </hp:secPr>
    <hp:t>Before </hp:t>
  </hp:run>
  <hp:run charPrIDRef="8"><hp:t>after</hp:t></hp:run>
</hp:p>
<hp:p paraPrIDRef="20"></hp:p>
</hp:section>"#,
        );

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(
                    paragraph.runs.iter().map(|run| run.text.as_str()).collect::<String>(),
                    "Before after"
                );
            }
            _ => panic!("expected first block paragraph"),
        }

        match &parsed.document.sections[0].blocks[1] {
            Block::Paragraph(paragraph) => {
                assert!(paragraph.runs.is_empty());
            }
            _ => panic!("expected second block blank paragraph"),
        }
    }

    #[test]
    fn accepts_application_hwp_zip_mimetype() {
        let bytes = fixture_hwpx_bytes("application/hwp+zip");

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("application/hwp+zip should be accepted");

        assert_eq!(
            parsed.diagnostics.version_hint.as_deref(),
            Some("application/hwp+zip")
        );
    }

    #[test]
    fn parses_numbering_and_header_footer_controls() {
        let bytes = fixture_hwpx_bytes_with_section(
            "application/hwpx+zip",
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:ctrl>
  <hp:header id="10" applyPageType="BOTH">
    <hp:subList>
      <hp:p paraPrIDRef="20">
        <hp:run><hp:t>Header </hp:t></hp:run>
        <hp:run><hp:pageNum pos="TOP_CENTER" formatType="DIGIT" sideChar="-" /></hp:run>
      </hp:p>
    </hp:subList>
  </hp:header>
</hp:ctrl>
<hp:p paraPrIDRef="21"><hp:run charPrIDRef="7"><hp:t>First item</hp:t></hp:run></hp:p>
<hp:p paraPrIDRef="22"><hp:run charPrIDRef="8"><hp:t>Bullet item</hp:t></hp:run></hp:p>
<hp:ctrl>
  <hp:footer id="11" applyPageType="BOTH">
    <hp:subList>
      <hp:p paraPrIDRef="20">
        <hp:run><hp:t>Total </hp:t><hp:autoNum numType="TOTAL_PAGE"><hp:autoNumFormat type="DIGIT" /></hp:autoNum></hp:run>
      </hp:p>
    </hp:subList>
  </hp:footer>
</hp:ctrl>
</hp:section>"#,
        );

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        assert_eq!(parsed.document.sections[0].page_start_number, Some(3));
        assert_eq!(parsed.document.sections[0].headers.len(), 1);
        assert_eq!(parsed.document.sections[0].footers.len(), 1);

        match &parsed.document.sections[0].headers[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(
                    paragraph.runs.iter().map(|run| run.text.as_str()).collect::<String>(),
                    "Header {{PAGE:DIGIT:-}}"
                );
            }
            _ => panic!("expected header paragraph"),
        }

        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(paragraph.marker.as_ref().map(|marker| marker.text.as_str()), Some("1. "));
            }
            _ => panic!("expected first body paragraph"),
        }

        match &parsed.document.sections[0].blocks[1] {
            Block::Paragraph(paragraph) => {
                assert_eq!(paragraph.marker.as_ref().map(|marker| marker.text.as_str()), Some("• "));
            }
            _ => panic!("expected second body paragraph"),
        }
    }

    #[test]
    fn applies_linesegarray_breaks_to_runs() {
        let bytes = fixture_hwpx_bytes_with_section(
            "application/hwpx+zip",
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:p paraPrIDRef="20">
  <hp:run charPrIDRef="7"><hp:t>ABCDE</hp:t></hp:run>
  <hp:run charPrIDRef="8"><hp:t>FGHIJ</hp:t></hp:run>
  <hp:linesegarray>
    <hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="1000" flags="0" />
    <hp:lineseg textpos="5" vertpos="1000" vertsize="1000" textheight="1000" baseline="850" spacing="0" horzpos="0" horzsize="1000" flags="0" />
  </hp:linesegarray>
</hp:p>
</hp:section>"#,
        );

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        match &parsed.document.sections[0].blocks[0] {
            Block::Paragraph(paragraph) => {
                assert_eq!(
                    paragraph.runs.iter().map(|run| run.text.as_str()).collect::<String>(),
                    "ABCDE\nFGHIJ"
                );
            }
            _ => panic!("expected body paragraph"),
        }
    }

    #[test]
    fn parses_table_cells_and_embedded_assets() {
        let bytes = fixture_hwpx_bytes_with_entries(
            "application/hwpx+zip",
            br#"<opf:package xmlns:opf="http://www.idpf.org/2007/opf">
<opf:manifest>
  <opf:item id="sec0" href="section0.xml" />
  <opf:item id="img1" href="../BinData/sample.png" />
</opf:manifest>
<opf:spine>
  <opf:itemref idref="sec0" />
</opf:spine>
</opf:package>"#,
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:tbl repeatHeader="1" cellSpacing="120">
  <hp:sz width="22000" height="8000" widthRelTo="PAGE" heightRelTo="ABSOLUTE" />
  <hp:tr>
    <hp:tc header="1">
      <hp:cellSpan colSpan="2" rowSpan="1" />
      <hp:cellSz width="12000" height="3600" />
      <hp:cellMargin left="300" right="400" top="200" bottom="250" />
      <hp:subList>
        <hp:p paraPrIDRef="20"><hp:run><hp:t>Cell text</hp:t></hp:run></hp:p>
        <hp:pic id="77" zOrder="3" textWrap="SQUARE">
          <hp:sz width="7200" height="5400" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" />
          <hp:pos treatAsChar="0" horzAlign="RIGHT" vertAlign="TOP" horzOffset="1200" vertOffset="600" />
          <hp:img binaryItemIDRef="img1" />
        </hp:pic>
      </hp:subList>
    </hp:tc>
  </hp:tr>
</hp:tbl>
</hp:section>"#,
            &[("BinData/sample.png", b"fake-png")],
        );

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        assert_eq!(parsed.document.assets.len(), 1);
        assert_eq!(parsed.document.assets[0].id, "BinData/sample.png");
        assert!(parsed.document.assets[0].data_uri.as_ref().is_some());

        match &parsed.document.sections[0].blocks[0] {
            Block::Table(table) => {
                assert_eq!(table.width, Some(22000));
                assert!(table.repeat_header);
                assert_eq!(table.cell_spacing, Some(120));

                let cell = &table.rows[0].cells[0];
                assert_eq!(cell.col_span, Some(2));
                assert_eq!(cell.width, Some(12000));
                assert_eq!(cell.padding_left, Some(300));
                assert!(cell.is_header);
                assert_eq!(cell.blocks.len(), 2);

                match &cell.blocks[1] {
                    Block::Image(image) => {
                        assert_eq!(image.kind, "pic");
                        assert_eq!(image.asset_id, "BinData/sample.png");
                        assert_eq!(image.width, Some(7200));
                        assert!(!image.treat_as_char);
                        assert_eq!(image.horz_align.as_deref(), Some("RIGHT"));
                        assert_eq!(image.horz_offset, Some(1200));
                    }
                    _ => panic!("expected embedded image block"),
                }
            }
            _ => panic!("expected table block"),
        }
    }

    #[test]
    fn parses_ole_positioning_hints() {
        let bytes = fixture_hwpx_bytes_with_section(
            "application/hwpx+zip",
            br#"<hp:section xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
<hp:ole id="5" binaryItemIDRef="ole1" zOrder="7" textWrap="TOP_AND_BOTTOM">
  <hp:sz width="6400" height="4800" widthRelTo="ABSOLUTE" heightRelTo="ABSOLUTE" />
  <hp:pos treatAsChar="1" vertRelTo="PARA" horzRelTo="PAGE" vertAlign="CENTER" horzAlign="CENTER" vertOffset="300" horzOffset="450" />
</hp:ole>
</hp:section>"#,
        );

        let parsed = HwpxInspector
            .parse_bytes(&bytes, Some("fixture.hwpx"))
            .expect("fixture should parse");

        match &parsed.document.sections[0].blocks[0] {
            Block::Image(image) => {
                assert_eq!(image.kind, "ole");
                assert_eq!(image.asset_id, "ole:5");
                assert!(image.treat_as_char);
                assert_eq!(image.text_wrap.as_deref(), Some("TOP_AND_BOTTOM"));
                assert_eq!(image.vert_rel_to.as_deref(), Some("PARA"));
                assert_eq!(image.horz_rel_to.as_deref(), Some("PAGE"));
                assert_eq!(image.vert_align.as_deref(), Some("CENTER"));
                assert_eq!(image.horz_align.as_deref(), Some("CENTER"));
            }
            _ => panic!("expected ole image block"),
        }
    }
}

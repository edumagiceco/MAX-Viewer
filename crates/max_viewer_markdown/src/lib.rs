use std::{
    fs,
    path::{Path, PathBuf},
};

use base64::{Engine as _, engine::general_purpose::STANDARD as BASE64_STANDARD};
use max_viewer_core::{
    AssetRef, Block, Document, DocumentDiagnostics, DocumentFormat, DocumentMetadata,
    FormatInspector, FormatSupport, ImageBlock, PageLayout, Paragraph, ParagraphStyle, ParseError,
    Section, TableBlock, TableBorder, TableCell, TableCellStyle, TableRow, TextRun, TextStyle,
};
use pulldown_cmark::{
    Alignment, CodeBlockKind, Event, HeadingLevel, Options, Parser, Tag, TagEnd,
};

const A4_WIDTH: i32 = 59_528;
const A4_HEIGHT: i32 = 84_188;
const A4_MARGIN_LEFT_RIGHT: i32 = 8_504;
const A4_MARGIN_TOP: i32 = 5_668;
const A4_MARGIN_BOTTOM: i32 = 4_252;
const A4_MARGIN_HEADER_FOOTER: i32 = 4_252;

const FONT_SIZE_BASE: i32 = 1_200;
const FONT_SIZE_H1: i32 = 2_600;
const FONT_SIZE_H2: i32 = 2_100;
const FONT_SIZE_H3: i32 = 1_700;
const FONT_SIZE_H4: i32 = 1_500;
const FONT_SIZE_H5: i32 = 1_300;
const FONT_SIZE_H6: i32 = 1_200;

#[derive(Debug, Default)]
pub struct MarkdownInspector;

#[derive(Debug, Clone)]
pub struct MarkdownParseResult {
    pub document: Document,
    pub diagnostics: DocumentDiagnostics,
}

#[derive(Debug, Clone, Default)]
struct ParseState {
    assets: Vec<AssetRef>,
    next_asset_index: usize,
    base_dir: Option<PathBuf>,
    discovered_title: Option<String>,
}

#[derive(Debug, Clone, Default)]
struct InlineStyleState {
    bold: bool,
    italic: bool,
    underline: bool,
    code: bool,
    text_color: Option<String>,
}

#[derive(Debug, Clone)]
enum InlineElement {
    Run(TextRun),
    Image {
        alt: String,
        dest_url: String,
        title: Option<String>,
    },
}

impl MarkdownInspector {
    pub fn scaffold_support() -> FormatSupport {
        FormatSupport {
            format: DocumentFormat::Markdown,
            status: "active".to_string(),
            implemented: vec![
                "UTF-8 Markdown parsing".to_string(),
                "heading, paragraph, list, table, blockquote, and code block mapping".to_string(),
                "plain text export through the shared document model".to_string(),
                "basic local image asset loading for path-based opens".to_string(),
            ],
            planned: vec![
                "remote image fetching".to_string(),
                "front matter metadata extraction".to_string(),
                "HTML block rendering".to_string(),
            ],
        }
    }

    pub fn parse_bytes(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
    ) -> Result<MarkdownParseResult, ParseError> {
        self.parse_bytes_with_base_dir(bytes, fallback_title, None)
    }

    pub fn parse_bytes_with_base_dir(
        &self,
        bytes: &[u8],
        fallback_title: Option<&str>,
        base_dir: Option<&Path>,
    ) -> Result<MarkdownParseResult, ParseError> {
        let diagnostics = self.inspect_bytes(bytes)?;
        let markdown =
            std::str::from_utf8(bytes).map_err(|error| ParseError::InvalidData(error.to_string()))?;

        let mut state = ParseState {
            assets: Vec::new(),
            next_asset_index: 1,
            base_dir: base_dir.map(Path::to_path_buf),
            discovered_title: None,
        };
        let mut events = Parser::new_ext(markdown, markdown_options()).peekable();
        let blocks = parse_blocks(&mut events, &mut state, None, 0)?;

        let metadata_title = state
            .discovered_title
            .clone()
            .or_else(|| fallback_title.map(ToOwned::to_owned));

        Ok(MarkdownParseResult {
            document: Document {
                format: Some(DocumentFormat::Markdown),
                metadata: DocumentMetadata {
                    title: metadata_title,
                    language: Some("markdown".to_string()),
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
                assets: state.assets,
            },
            diagnostics,
        })
    }
}

impl FormatInspector for MarkdownInspector {
    fn format(&self) -> DocumentFormat {
        DocumentFormat::Markdown
    }

    fn inspect_bytes(&self, bytes: &[u8]) -> Result<DocumentDiagnostics, ParseError> {
        let text =
            std::str::from_utf8(bytes).map_err(|error| ParseError::InvalidData(error.to_string()))?;
        let line_count = text.lines().count();
        let notes = vec![
            "UTF-8 Markdown text stream was parsed directly.".to_string(),
            format!("Detected {line_count} source lines."),
        ];

        Ok(DocumentDiagnostics {
            format: DocumentFormat::Markdown,
            entry_count: 1,
            section_count: 1,
            asset_count: 0,
            is_encrypted: false,
            version_hint: Some("CommonMark".to_string()),
            notes,
        })
    }
}

fn markdown_options() -> Options {
    let mut options = Options::empty();
    options.insert(Options::ENABLE_TABLES);
    options.insert(Options::ENABLE_STRIKETHROUGH);
    options.insert(Options::ENABLE_TASKLISTS);
    options
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

fn default_paragraph_style(blockquote_depth: usize) -> ParagraphStyle {
    ParagraphStyle {
        align: Some("LEFT".to_string()),
        margin_left: (blockquote_depth > 0).then_some((blockquote_depth as i32) * 1_400),
        margin_right: None,
        margin_prev: Some(500),
        margin_next: Some(500),
        line_spacing_type: Some("PERCENT".to_string()),
        line_spacing: Some(170),
        heading_type: None,
        heading_id_ref: None,
        heading_level: None,
        marker_align: None,
        marker_width_adjust: None,
        marker_text_offset_type: None,
        marker_text_offset: None,
        keep_with_next: false,
        keep_lines: false,
        indent: None,
    }
}

fn heading_style(level: HeadingLevel) -> ParagraphStyle {
    let mut style = default_paragraph_style(0);
    style.margin_prev = Some(match level {
        HeadingLevel::H1 => 0,
        _ => 900,
    });
    style.margin_next = Some(match level {
        HeadingLevel::H1 => 1_200,
        HeadingLevel::H2 => 1_000,
        _ => 800,
    });
    style.heading_type = Some("OUTLINE".to_string());
    style.heading_level = Some(match level {
        HeadingLevel::H1 => 1,
        HeadingLevel::H2 => 2,
        HeadingLevel::H3 => 3,
        HeadingLevel::H4 => 4,
        HeadingLevel::H5 => 5,
        HeadingLevel::H6 => 6,
    });
    style.keep_with_next = true;
    style
}

fn list_paragraph_style(depth: usize, blockquote_depth: usize) -> ParagraphStyle {
    let mut style = default_paragraph_style(blockquote_depth);
    style.margin_left = Some(((depth + 1) as i32) * 1_600);
    style.marker_align = Some("LEFT".to_string());
    style.marker_width_adjust = Some(1_100);
    style.marker_text_offset_type = Some("ABSOLUTE".to_string());
    style.marker_text_offset = Some(300);
    style
}

fn base_text_style(font_size: i32) -> TextStyle {
    TextStyle {
        font_family: None,
        font_size: Some(font_size),
        text_color: Some("#0F172A".to_string()),
        background_color: None,
        underline_color: None,
        width_ratio: None,
        letter_spacing: None,
        relative_size: None,
        baseline_offset: None,
        use_font_space: false,
        use_kerning: true,
        bold: false,
        italic: false,
        underline: false,
    }
}

fn border(style: &str, width: &str, color: &str) -> TableBorder {
    TableBorder {
        style: Some(style.to_string()),
        width: Some(width.to_string()),
        color: Some(color.to_string()),
    }
}

fn parse_blocks<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    until: Option<TagEnd>,
    blockquote_depth: usize,
) -> Result<Vec<Block>, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut blocks = Vec::new();

    while let Some(event) = events.next() {
        match event {
            Event::Start(tag) => match tag {
                Tag::Paragraph => {
                    let block = parse_paragraph_block(
                        events,
                        state,
                        default_paragraph_style(blockquote_depth),
                        None,
                        FONT_SIZE_BASE,
                        TagEnd::Paragraph,
                    )?;
                    blocks.push(block);
                }
                Tag::Heading { level, .. } => {
                    let block = parse_paragraph_block(
                        events,
                        state,
                        heading_style(level),
                        None,
                        heading_font_size(level),
                        TagEnd::Heading(level),
                    )?;
                    maybe_capture_title(&block, state);
                    blocks.push(block);
                }
                Tag::BlockQuote(kind) => {
                    blocks.extend(parse_blocks(
                        events,
                        state,
                        Some(TagEnd::BlockQuote(kind)),
                        blockquote_depth + 1,
                    )?);
                }
                Tag::List(start) => {
                    blocks.extend(parse_list(events, state, start, blockquote_depth)?);
                }
                Tag::CodeBlock(kind) => {
                    blocks.push(parse_code_block(events, kind)?);
                }
                Tag::Table(alignments) => {
                    blocks.push(parse_table(events, state, alignments)?);
                }
                Tag::HtmlBlock => {
                    consume_until(events, TagEnd::HtmlBlock);
                    blocks.push(Block::Unsupported(max_viewer_core::UnsupportedBlock {
                        kind: "htmlBlock".to_string(),
                        reason: Some("HTML blocks are not rendered in Markdown mode yet.".to_string()),
                        page_break_before: false,
                    }));
                }
                _ => {}
            },
            Event::Rule => blocks.push(make_rule_block()),
            Event::End(end) => {
                if until.as_ref().is_some_and(|expected| *expected == end) {
                    break;
                }
            }
            _ => {}
        }
    }

    Ok(blocks)
}

fn parse_list<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    start: Option<u64>,
    blockquote_depth: usize,
) -> Result<Vec<Block>, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut blocks = Vec::new();
    let mut next_index = start.unwrap_or(1);
    let ordered = start.is_some();

    while let Some(event) = events.next() {
        match event {
            Event::Start(Tag::Item) => {
                let marker_text = if ordered {
                    let text = format!("{next_index}. ");
                    next_index += 1;
                    text
                } else {
                    "• ".to_string()
                };
                blocks.extend(parse_list_item(
                    events,
                    state,
                    marker_text,
                    ordered,
                    blockquote_depth,
                )?);
            }
            Event::End(TagEnd::List(_)) => break,
            _ => {}
        }
    }

    Ok(blocks)
}

fn parse_list_item<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    marker_text: String,
    ordered: bool,
    blockquote_depth: usize,
) -> Result<Vec<Block>, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut blocks = Vec::new();
    let mut first_content_block = true;

    while let Some(event) = events.next() {
        match event {
            Event::Start(tag) => match tag {
                Tag::Paragraph => {
                    let marker = first_content_block.then(|| TextRun {
                        text: marker_text.clone(),
                        style: Some(TextStyle {
                            bold: ordered,
                            ..base_text_style(FONT_SIZE_BASE)
                        }),
                    });
                    let block = parse_paragraph_block(
                        events,
                        state,
                        list_paragraph_style(0, blockquote_depth),
                        marker,
                        FONT_SIZE_BASE,
                        TagEnd::Paragraph,
                    )?;
                    blocks.push(block);
                    first_content_block = false;
                }
                Tag::List(start) => {
                    blocks.extend(parse_list(events, state, start, blockquote_depth + 1)?);
                    first_content_block = false;
                }
                Tag::CodeBlock(kind) => {
                    blocks.push(parse_code_block(events, kind)?);
                    first_content_block = false;
                }
                Tag::Table(alignments) => {
                    blocks.push(parse_table(events, state, alignments)?);
                    first_content_block = false;
                }
                Tag::Heading { level, .. } => {
                    let marker = first_content_block.then(|| TextRun {
                        text: marker_text.clone(),
                        style: Some(base_text_style(heading_font_size(level))),
                    });
                    let block = parse_paragraph_block(
                        events,
                        state,
                        heading_style(level),
                        marker,
                        heading_font_size(level),
                        TagEnd::Heading(level),
                    )?;
                    maybe_capture_title(&block, state);
                    blocks.push(block);
                    first_content_block = false;
                }
                _ => {}
            },
            Event::Text(text) => {
                let block = Block::Paragraph(Paragraph {
                    marker: first_content_block.then(|| TextRun {
                        text: marker_text.clone(),
                        style: Some(base_text_style(FONT_SIZE_BASE)),
                    }),
                    runs: vec![TextRun {
                        text: text.into_string(),
                        style: Some(base_text_style(FONT_SIZE_BASE)),
                    }],
                    style: Some(list_paragraph_style(0, blockquote_depth)),
                    line_segment_count: None,
                    layout_height_hint: None,
                    page_break_before: false,
                });
                blocks.push(block);
                first_content_block = false;
            }
            Event::End(TagEnd::Item) => break,
            _ => {}
        }
    }

    Ok(blocks)
}

fn parse_paragraph_block<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    style: ParagraphStyle,
    marker: Option<TextRun>,
    font_size: i32,
    end: TagEnd,
) -> Result<Block, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let elements = collect_inline_elements_until(events, state, end, InlineStyleState::default(), font_size)?;
    if elements.len() == 1 {
        if let InlineElement::Image {
            alt,
            dest_url,
            title,
        } = &elements[0]
        {
            if let Some(image_block) = resolve_markdown_image(state, dest_url, title.as_deref(), alt) {
                return Ok(Block::Image(image_block));
            }
        }
    }

    Ok(Block::Paragraph(Paragraph {
        marker,
        runs: flatten_inline_elements(elements, font_size),
        style: Some(style),
        line_segment_count: None,
        layout_height_hint: None,
        page_break_before: false,
    }))
}

fn parse_code_block<'a, I>(
    events: &mut std::iter::Peekable<I>,
    kind: CodeBlockKind<'a>,
) -> Result<Block, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut code = String::new();

    while let Some(event) = events.next() {
        match event {
            Event::Text(text) | Event::Code(text) | Event::Html(text) | Event::InlineHtml(text) => {
                code.push_str(text.as_ref());
            }
            Event::SoftBreak | Event::HardBreak => code.push('\n'),
            Event::End(TagEnd::CodeBlock) => break,
            _ => {}
        }
    }

    let _info = match kind {
        CodeBlockKind::Indented => None,
        CodeBlockKind::Fenced(info) => {
            let trimmed = info.trim();
            (!trimmed.is_empty()).then(|| trimmed.to_string())
        }
    };

    Ok(Block::Table(TableBlock {
        rows: vec![TableRow {
            cells: vec![TableCell {
                text: String::new(),
                blocks: vec![Block::Paragraph(Paragraph {
                    marker: None,
                    runs: vec![TextRun {
                        text: code.trim_end_matches('\n').to_string(),
                        style: Some(TextStyle {
                            font_family: Some("SF Mono".to_string()),
                            background_color: Some("#F8FAFC".to_string()),
                            ..base_text_style(1_050)
                        }),
                    }],
                    style: Some(ParagraphStyle {
                        margin_prev: Some(0),
                        margin_next: Some(0),
                        line_spacing_type: Some("PERCENT".to_string()),
                        line_spacing: Some(150),
                        ..ParagraphStyle::default()
                    }),
                    line_segment_count: None,
                    layout_height_hint: None,
                    page_break_before: false,
                })],
                col_span: Some(1),
                row_span: Some(1),
                width: None,
                height: None,
                padding_left: Some(720),
                padding_right: Some(720),
                padding_top: Some(540),
                padding_bottom: Some(540),
                style: Some(TableCellStyle {
                    background_color: Some("#F8FAFC".to_string()),
                    background_image: None,
                    border_left: Some(border("SOLID", "0.2mm", "#CBD5E1")),
                    border_right: Some(border("SOLID", "0.2mm", "#CBD5E1")),
                    border_top: Some(border("SOLID", "0.2mm", "#CBD5E1")),
                    border_bottom: Some(border("SOLID", "0.2mm", "#CBD5E1")),
                    diagonal: None,
                }),
                is_header: false,
            }],
        }],
        width: None,
        height: None,
        cell_spacing: Some(0),
        style: None,
        no_adjust: false,
        repeat_header: false,
        header_row_count: None,
        page_break_before: false,
    }))
}

fn parse_table<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    alignments: Vec<Alignment>,
) -> Result<Block, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut rows = Vec::new();

    while let Some(event) = events.next() {
        match event {
            Event::Start(Tag::TableHead) => {
                rows.push(parse_table_head(events, state, &alignments)?);
            }
            Event::Start(Tag::TableRow) => {
                rows.push(parse_table_row(events, state, &alignments, false)?);
            }
            Event::End(TagEnd::Table) => break,
            _ => {}
        }
    }

    Ok(Block::Table(TableBlock {
        rows,
        width: None,
        height: None,
        cell_spacing: Some(0),
        style: Some(TableCellStyle {
            background_color: None,
            background_image: None,
            border_left: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_right: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_top: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_bottom: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            diagonal: None,
        }),
        no_adjust: false,
        repeat_header: true,
        header_row_count: Some(1),
        page_break_before: false,
    }))
}

fn parse_table_head<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    alignments: &[Alignment],
) -> Result<TableRow, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut cells = Vec::new();

    while let Some(event) = events.next() {
        match event {
            Event::Start(Tag::TableCell) => {
                let cell_index = cells.len();
                cells.push(parse_table_cell(
                    events,
                    state,
                    alignments.get(cell_index).copied().unwrap_or(Alignment::None),
                    true,
                )?);
            }
            Event::End(TagEnd::TableHead) => break,
            _ => {}
        }
    }

    Ok(TableRow { cells })
}

fn parse_table_row<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    alignments: &[Alignment],
    is_header: bool,
) -> Result<TableRow, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut cells = Vec::new();

    while let Some(event) = events.next() {
        match event {
            Event::Start(Tag::TableCell) => {
                let cell_index = cells.len();
                cells.push(parse_table_cell(
                    events,
                    state,
                    alignments.get(cell_index).copied().unwrap_or(Alignment::None),
                    is_header,
                )?);
            }
            Event::End(TagEnd::TableRow) => break,
            _ => {}
        }
    }

    Ok(TableRow { cells })
}

fn parse_table_cell<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    alignment: Alignment,
    is_header: bool,
) -> Result<TableCell, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let elements = collect_inline_elements_until(
        events,
        state,
        TagEnd::TableCell,
        InlineStyleState {
            bold: is_header,
            ..InlineStyleState::default()
        },
        FONT_SIZE_BASE,
    )?;
    let runs = flatten_inline_elements(elements, FONT_SIZE_BASE);
    let text = runs.iter().map(|run| run.text.as_str()).collect::<String>();

    Ok(TableCell {
        text,
        blocks: vec![Block::Paragraph(Paragraph {
            marker: None,
            runs,
            style: Some(ParagraphStyle {
                align: Some(match alignment {
                    Alignment::Left => "LEFT",
                    Alignment::Center => "CENTER",
                    Alignment::Right => "RIGHT",
                    Alignment::None => "LEFT",
                }
                .to_string()),
                margin_prev: Some(0),
                margin_next: Some(0),
                line_spacing_type: Some("PERCENT".to_string()),
                line_spacing: Some(160),
                ..ParagraphStyle::default()
            }),
            line_segment_count: None,
            layout_height_hint: None,
            page_break_before: false,
        })],
        col_span: Some(1),
        row_span: Some(1),
        width: None,
        height: None,
        padding_left: Some(420),
        padding_right: Some(420),
        padding_top: Some(260),
        padding_bottom: Some(260),
        style: Some(TableCellStyle {
            background_color: is_header.then_some("#F8FAFC".to_string()),
            background_image: None,
            border_left: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_right: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_top: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            border_bottom: Some(border("SOLID", "0.2mm", "#CBD5E1")),
            diagonal: None,
        }),
        is_header,
    })
}

fn collect_inline_elements_until<'a, I>(
    events: &mut std::iter::Peekable<I>,
    state: &mut ParseState,
    end: TagEnd,
    inline_style: InlineStyleState,
    font_size: i32,
) -> Result<Vec<InlineElement>, ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut elements = Vec::new();

    while let Some(event) = events.next() {
        match event {
            Event::End(tag_end) if tag_end == end => break,
            Event::Text(text) => push_inline_run(
                &mut elements,
                text.as_ref(),
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::Code(text) => {
                let mut code_style = inline_style.clone();
                code_style.code = true;
                push_inline_run(
                    &mut elements,
                    text.as_ref(),
                    inline_style_to_text_style(&code_style, font_size),
                );
            }
            Event::SoftBreak => push_inline_run(
                &mut elements,
                " ",
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::HardBreak => push_inline_run(
                &mut elements,
                "\n",
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::InlineHtml(html) | Event::Html(html) => push_inline_run(
                &mut elements,
                html.as_ref(),
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::FootnoteReference(label) => push_inline_run(
                &mut elements,
                &format!("[^{label}]"),
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::TaskListMarker(checked) => push_inline_run(
                &mut elements,
                if checked { "[x] " } else { "[ ] " },
                inline_style_to_text_style(&inline_style, font_size),
            ),
            Event::Start(tag) => match tag {
                Tag::Emphasis => {
                    let mut nested = inline_style.clone();
                    nested.italic = true;
                    elements.extend(collect_inline_elements_until(
                        events,
                        state,
                        TagEnd::Emphasis,
                        nested,
                        font_size,
                    )?);
                }
                Tag::Strong => {
                    let mut nested = inline_style.clone();
                    nested.bold = true;
                    elements.extend(collect_inline_elements_until(
                        events,
                        state,
                        TagEnd::Strong,
                        nested,
                        font_size,
                    )?);
                }
                Tag::Strikethrough => {
                    elements.extend(collect_inline_elements_until(
                        events,
                        state,
                        TagEnd::Strikethrough,
                        inline_style.clone(),
                        font_size,
                    )?);
                }
                Tag::Link { .. } => {
                    let mut nested = inline_style.clone();
                    nested.underline = true;
                    nested.text_color = Some("#2563EB".to_string());
                    elements.extend(collect_inline_elements_until(
                        events,
                        state,
                        TagEnd::Link,
                        nested,
                        font_size,
                    )?);
                }
                Tag::Image {
                    dest_url, title, ..
                } => {
                    let alt = inline_elements_to_plain_text(collect_inline_elements_until(
                        events,
                        state,
                        TagEnd::Image,
                        inline_style.clone(),
                        font_size,
                    )?);
                    elements.push(InlineElement::Image {
                        alt,
                        dest_url: dest_url.to_string(),
                        title: (!title.is_empty()).then(|| title.to_string()),
                    });
                }
                _ => {
                    consume_nested_tag(events, tag)?;
                }
            },
            _ => {}
        }
    }

    Ok(elements)
}

fn consume_nested_tag<'a, I>(
    events: &mut std::iter::Peekable<I>,
    tag: Tag<'a>,
) -> Result<(), ParseError>
where
    I: Iterator<Item = Event<'a>>,
{
    let mut depth = 1usize;
    let end = tag.to_end();
    while let Some(event) = events.next() {
        match event {
            Event::Start(start) if start.to_end() == end => depth += 1,
            Event::End(found) if found == end => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            _ => {}
        }
    }
    Ok(())
}

fn consume_until<'a, I>(events: &mut std::iter::Peekable<I>, end: TagEnd)
where
    I: Iterator<Item = Event<'a>>,
{
    while let Some(event) = events.next() {
        if let Event::End(found) = event {
            if found == end {
                break;
            }
        }
    }
}

fn push_inline_run(elements: &mut Vec<InlineElement>, text: &str, style: TextStyle) {
    if text.is_empty() {
        return;
    }

    if let Some(InlineElement::Run(last)) = elements.last_mut() {
        if last.style.as_ref() == Some(&style) {
            last.text.push_str(text);
            return;
        }
    }

    elements.push(InlineElement::Run(TextRun {
        text: text.to_string(),
        style: Some(style),
    }));
}

fn inline_style_to_text_style(inline: &InlineStyleState, font_size: i32) -> TextStyle {
    TextStyle {
        font_family: inline.code.then_some("SF Mono".to_string()),
        font_size: Some(font_size),
        text_color: inline
            .text_color
            .clone()
            .or_else(|| Some("#0F172A".to_string())),
        background_color: inline.code.then_some("#F8FAFC".to_string()),
        underline_color: inline.underline.then_some("#2563EB".to_string()),
        width_ratio: None,
        letter_spacing: None,
        relative_size: None,
        baseline_offset: None,
        use_font_space: false,
        use_kerning: true,
        bold: inline.bold,
        italic: inline.italic,
        underline: inline.underline,
    }
}

fn flatten_inline_elements(elements: Vec<InlineElement>, font_size: i32) -> Vec<TextRun> {
    let mut runs = Vec::new();

    for element in elements {
        match element {
            InlineElement::Run(run) => runs.push(run),
            InlineElement::Image { alt, dest_url, .. } => runs.push(TextRun {
                text: if alt.trim().is_empty() {
                    format!("[image: {dest_url}]")
                } else {
                    alt
                },
                style: Some(TextStyle {
                    italic: true,
                    text_color: Some("#475569".to_string()),
                    ..base_text_style(font_size)
                }),
            }),
        }
    }

    if runs.is_empty() {
        runs.push(TextRun {
            text: String::new(),
            style: Some(base_text_style(font_size)),
        });
    }

    runs
}

fn inline_elements_to_plain_text(elements: Vec<InlineElement>) -> String {
    let mut output = String::new();
    for element in elements {
        match element {
            InlineElement::Run(run) => output.push_str(&run.text),
            InlineElement::Image { alt, .. } => output.push_str(&alt),
        }
    }
    output
}

fn heading_font_size(level: HeadingLevel) -> i32 {
    match level {
        HeadingLevel::H1 => FONT_SIZE_H1,
        HeadingLevel::H2 => FONT_SIZE_H2,
        HeadingLevel::H3 => FONT_SIZE_H3,
        HeadingLevel::H4 => FONT_SIZE_H4,
        HeadingLevel::H5 => FONT_SIZE_H5,
        HeadingLevel::H6 => FONT_SIZE_H6,
    }
}

fn maybe_capture_title(block: &Block, state: &mut ParseState) {
    if state.discovered_title.is_some() {
        return;
    }
    if let Block::Paragraph(paragraph) = block {
        let title = paragraph
            .runs
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>()
            .trim()
            .to_string();
        if !title.is_empty() {
            state.discovered_title = Some(title);
        }
    }
}

fn make_rule_block() -> Block {
    Block::Table(TableBlock {
        rows: vec![TableRow {
            cells: vec![TableCell {
                text: String::new(),
                blocks: Vec::new(),
                col_span: Some(1),
                row_span: Some(1),
                width: None,
                height: Some(120),
                padding_left: Some(0),
                padding_right: Some(0),
                padding_top: Some(0),
                padding_bottom: Some(0),
                style: Some(TableCellStyle {
                    background_color: None,
                    background_image: None,
                    border_left: None,
                    border_right: None,
                    border_top: Some(border("SOLID", "0.3mm", "#CBD5E1")),
                    border_bottom: None,
                    diagonal: None,
                }),
                is_header: false,
            }],
        }],
        width: None,
        height: None,
        cell_spacing: Some(0),
        style: None,
        no_adjust: false,
        repeat_header: false,
        header_row_count: None,
        page_break_before: false,
    })
}

fn resolve_markdown_image(
    state: &mut ParseState,
    dest_url: &str,
    title: Option<&str>,
    alt: &str,
) -> Option<ImageBlock> {
    let resolved = resolve_image_asset(state, dest_url)?;
    Some(ImageBlock {
        kind: "image".to_string(),
        asset_id: resolved,
        alt_text: (!alt.trim().is_empty()).then(|| alt.trim().to_string()),
        width: None,
        height: None,
        width_rel_to: None,
        height_rel_to: None,
        treat_as_char: true,
        text_wrap: None,
        z_order: None,
        vert_rel_to: None,
        horz_rel_to: None,
        vert_align: None,
        horz_align: None,
        vert_offset: None,
        horz_offset: None,
        distance_left: None,
        distance_right: None,
        distance_top: None,
        distance_bottom: None,
        rotation: None,
        caption: title.map(ToOwned::to_owned),
        page_break_before: false,
    })
}

fn resolve_image_asset(state: &mut ParseState, dest_url: &str) -> Option<String> {
    if let Some((media_type, data_uri)) = parse_data_uri(dest_url) {
        let asset_id = next_asset_id(state);
        state.assets.push(AssetRef {
            id: asset_id.clone(),
            media_type,
            source_path: None,
            data_uri: Some(data_uri),
        });
        return Some(asset_id);
    }

    let base_dir = state.base_dir.as_ref()?;
    if dest_url.starts_with("http://") || dest_url.starts_with("https://") {
        return None;
    }

    let path = if Path::new(dest_url).is_absolute() {
        PathBuf::from(dest_url)
    } else {
        base_dir.join(dest_url)
    };

    let bytes = fs::read(&path).ok()?;
    let media_type = detect_media_type(&path)?;
    let asset_id = next_asset_id(state);
    state.assets.push(AssetRef {
        id: asset_id.clone(),
        media_type: media_type.to_string(),
        source_path: Some(path.to_string_lossy().into_owned()),
        data_uri: Some(format!(
            "data:{media_type};base64,{}",
            BASE64_STANDARD.encode(bytes)
        )),
    });
    Some(asset_id)
}

fn parse_data_uri(uri: &str) -> Option<(String, String)> {
    let (prefix, data) = uri.split_once(',')?;
    let media_type = prefix
        .strip_prefix("data:")?
        .split(';')
        .next()
        .filter(|value| !value.is_empty())
        .unwrap_or("text/plain")
        .to_string();
    Some((media_type, format!("{prefix},{data}")))
}

fn detect_media_type(path: &Path) -> Option<&'static str> {
    match path.extension()?.to_str()?.to_ascii_lowercase().as_str() {
        "png" => Some("image/png"),
        "jpg" | "jpeg" => Some("image/jpeg"),
        "gif" => Some("image/gif"),
        "svg" => Some("image/svg+xml"),
        "webp" => Some("image/webp"),
        "bmp" => Some("image/bmp"),
        _ => None,
    }
}

fn next_asset_id(state: &mut ParseState) -> String {
    let id = format!("md-asset-{}", state.next_asset_index);
    state.next_asset_index += 1;
    id
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parses_basic_markdown_document() {
        let bytes = br#"# Sample Title

Paragraph with **bold** text.

- first
- second

| Name | Value |
| ---- | ----: |
| a | 1 |
"#;

        let inspector = MarkdownInspector;
        let parsed = inspector
            .parse_bytes(bytes, Some("sample.md"))
            .expect("markdown should parse");

        assert_eq!(parsed.diagnostics.format, DocumentFormat::Markdown);
        assert_eq!(parsed.document.metadata.title.as_deref(), Some("Sample Title"));
        assert_eq!(parsed.document.sections.len(), 1);
        assert!(parsed.document.sections[0].blocks.len() >= 4);
        assert!(parsed
            .document
            .sections[0]
            .blocks
            .iter()
            .any(|block| matches!(block, Block::Table(_))));
    }

    #[test]
    fn resolves_data_uri_images_into_assets() {
        let bytes = br#"![logo](data:image/png;base64,AAAA)"#;
        let inspector = MarkdownInspector;
        let parsed = inspector
            .parse_bytes(bytes, Some("image.md"))
            .expect("markdown should parse");

        assert_eq!(parsed.document.assets.len(), 1);
        assert!(matches!(
            parsed.document.sections[0].blocks.first(),
            Some(Block::Image(_))
        ));
    }
}

use max_viewer_core::{Block, Document};

pub fn to_plain_text(document: &Document) -> String {
    let mut output = String::new();

    for (section_index, section) in document.sections.iter().enumerate() {
        if section_index > 0 {
            output.push_str("\n\n");
        }

        for block in &section.blocks {
            write_block_text(block, &mut output);
        }
    }

    output.trim().to_string()
}

fn write_block_text(block: &Block, output: &mut String) {
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
                let line = row
                    .cells
                    .iter()
                    .map(cell_to_text)
                    .collect::<Vec<_>>()
                    .join(" | ");
                output.push_str(line.trim_end());
                output.push('\n');
            }
        }
        Block::Image(image) => {
            output.push_str("[");
            output.push_str(&image.kind);
            output.push(']');
            if let Some(caption) = &image.caption {
                output.push(' ');
                output.push_str(caption);
            } else if let Some(alt_text) = &image.alt_text {
                output.push(' ');
                output.push_str(alt_text);
            }
            output.push('\n');
        }
        Block::Unsupported(unsupported) => {
            output.push_str("[unsupported] ");
            output.push_str(&unsupported.kind);
            output.push('\n');
        }
    }
}

fn cell_to_text(cell: &max_viewer_core::TableCell) -> String {
    if !cell.blocks.is_empty() {
        let mut output = String::new();
        for block in &cell.blocks {
            write_block_text(block, &mut output);
        }
        let normalized = output.replace('\n', " ").trim().to_string();
        if !normalized.is_empty() {
            return normalized;
        }
    }

    cell.text.trim().to_string()
}

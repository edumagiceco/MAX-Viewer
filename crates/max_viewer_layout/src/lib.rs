use max_viewer_core::{Block, Document};
use serde::Serialize;

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Default)]
#[serde(rename_all = "camelCase")]
pub struct LayoutSummary {
    pub section_count: usize,
    pub paragraph_count: usize,
    pub table_count: usize,
    pub image_count: usize,
    pub unsupported_count: usize,
}

pub fn summarize(document: &Document) -> LayoutSummary {
    let mut summary = LayoutSummary {
        section_count: document.sections.len(),
        ..LayoutSummary::default()
    };

    for section in &document.sections {
        for block in &section.blocks {
            accumulate_block_summary(block, &mut summary);
        }
    }

    summary
}

fn accumulate_block_summary(block: &Block, summary: &mut LayoutSummary) {
    match block {
        Block::Paragraph(_) => summary.paragraph_count += 1,
        Block::Table(table) => {
            summary.table_count += 1;
            for row in &table.rows {
                for cell in &row.cells {
                    for nested_block in &cell.blocks {
                        accumulate_block_summary(nested_block, summary);
                    }
                }
            }
        }
        Block::Image(_) => summary.image_count += 1,
        Block::Unsupported(_) => summary.unsupported_count += 1,
    }
}

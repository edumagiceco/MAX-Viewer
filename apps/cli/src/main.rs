use std::fs;
use std::path::PathBuf;
use std::process;

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        eprintln!("MAX Viewer CLI — HWP/HWPX/Markdown/PDF 문서 변환 도구");
        eprintln!();
        eprintln!("사용법:");
        eprintln!("  max-viewer <파일경로>              문서 정보 출력");
        eprintln!("  max-viewer <파일경로> --text        텍스트 추출");
        eprintln!("  max-viewer <파일경로> --json        JSON 문서 모델 출력");
        eprintln!("  max-viewer <파일경로> --inspect     진단 정보 출력");
        process::exit(1);
    }

    let path = PathBuf::from(&args[1]);
    let mode = args.get(2).map(|s| s.as_str()).unwrap_or("--info");

    let bytes = match fs::read(&path) {
        Ok(b) => b,
        Err(e) => {
            eprintln!("파일 읽기 실패: {e}");
            process::exit(1);
        }
    };

    let file_name = path
        .file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_default();
    let ext = path
        .extension()
        .map(|e| e.to_string_lossy().to_lowercase())
        .unwrap_or_default();

    let (document, diagnostics) = match ext.as_str() {
        "hwpx" => {
            let inspector = max_viewer_hwpx::HwpxInspector;
            match inspector.parse_bytes(&bytes, Some(&file_name)) {
                Ok(result) => (result.document, result.diagnostics),
                Err(e) => {
                    eprintln!("HWPX 파싱 실패: {e}");
                    process::exit(1);
                }
            }
        }
        "hwp" => {
            let inspector = max_viewer_hwp::HwpInspector;
            match inspector.parse_bytes(&bytes, Some(&file_name)) {
                Ok(result) => (result.document, result.diagnostics),
                Err(e) => {
                    eprintln!("HWP 파싱 실패: {e}");
                    process::exit(1);
                }
            }
        }
        "md" | "markdown" => {
            let inspector = max_viewer_markdown::MarkdownInspector;
            match inspector.parse_bytes_with_base_dir(&bytes, Some(&file_name), path.parent()) {
                Ok(result) => (result.document, result.diagnostics),
                Err(e) => {
                    eprintln!("Markdown 파싱 실패: {e}");
                    process::exit(1);
                }
            }
        }
        "pdf" => {
            let inspector = max_viewer_pdf::PdfInspector;
            match inspector.parse_bytes(&bytes, Some(&file_name)) {
                Ok(result) => (result.document, result.diagnostics),
                Err(e) => {
                    eprintln!("PDF 파싱 실패: {e}");
                    process::exit(1);
                }
            }
        }
        _ => {
            eprintln!("지원하지 않는 파일 형식: .{ext}");
            process::exit(1);
        }
    };

    match mode {
        "--text" => {
            let text = max_viewer_export::to_plain_text(&document);
            println!("{text}");
        }
        "--json" => {
            let json = serde_json::to_string_pretty(&document).unwrap_or_default();
            println!("{json}");
        }
        "--inspect" => {
            let json = serde_json::to_string_pretty(&diagnostics).unwrap_or_default();
            println!("{json}");
        }
        _ => {
            println!("파일: {file_name}");
            println!("형식: {:?}", diagnostics.format);
            if let Some(ref v) = diagnostics.version_hint {
                println!("버전: {v}");
            }
            println!("구역: {}", diagnostics.section_count);
            println!("자산: {}", diagnostics.asset_count);
            if diagnostics.is_encrypted {
                println!("암호화: 예");
            }
            let text = max_viewer_export::to_plain_text(&document);
            let char_count = text.chars().count();
            println!("글자 수: {char_count}");
        }
    }
}

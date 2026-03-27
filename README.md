# MAX-Viewer

MAX-Viewer는 `.hwp`와 `.hwpx` 문서를 읽기 전용으로 열람하기 위한 멀티플랫폼 뷰어 프로젝트입니다.

현재 스캐폴드는 다음 구조를 기준으로 잡혀 있습니다.

- Rust workspace 기반 문서 코어
- `Tauri 2` 데스크톱 앱 셸
- `React + Vite` 프론트엔드

## 디렉토리 구조

```text
MAX-Viewer/
├── apps/
│   └── desktop-tauri/
├── crates/
│   ├── max_viewer_core/
│   ├── max_viewer_hwpx/
│   ├── max_viewer_hwp/
│   ├── max_viewer_layout/
│   └── max_viewer_export/
├── docs/
├── fixtures/
└── packages/
    └── viewer-ui/
```

## 시작 방법

1. `pnpm install`
2. `pnpm dev`

웹 UI만 확인하려면 `pnpm dev:web`를 사용할 수 있습니다.

## 현재 상태

- `max_viewer_core`: 공통 문서 모델과 파서 진단 타입
- `max_viewer_hwpx`: HWPX ZIP, `content.hpf`, `header.xml`, `section*.xml` 파서와 스타일/번호/머리말/표/개체 배치 해석
- `max_viewer_hwp`: HWP `FileHeader`, 최소 `BodyText` 문단 복원, `PrvText` fallback preview 파서
- `max_viewer_layout`: 문서 블록 수 기반 레이아웃 요약
- `max_viewer_export`: 내부 문서 모델의 텍스트 추출 스캐폴드
- `apps/desktop-tauri`: Tauri 명령, 네이티브 파일 열기 다이얼로그, 실제 문서 로드 셸
- `packages/viewer-ui`: 네이티브 문서 열기, 페이지 분할, 실제 이미지 렌더링, 표 셀 내부 조판, floating 개체 배치가 연결된 뷰어 UI

세부 설계는 [docs/README.md](/Users/magic/work/MAX-Viewer/docs/README.md)에 정리되어 있습니다.

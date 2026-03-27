# MAX-Viewer 사용 설명서

작성일: 2026-03-27

## 개요

MAX-Viewer는 `.hwp`, `.hwpx` 문서를 읽기 전용으로 여는 문서 뷰어입니다.

현재 기준으로 가장 안정적인 경로는 다음 두 가지입니다.

1. 개발 모드에서 실행해 바로 테스트하기
2. 빌드된 macOS `.app` 번들을 실행해 문서 열기

## 지원 형식

- `.hwpx`
  - 페이지 윤곽 표시
  - 문단/글자 스타일 일부 반영
  - 번호 매기기/글머리표 1차 반영
  - 머리말/꼬리말, 페이지 번호 placeholder 반영
  - 표 셀 내부 문단/개체 렌더링
  - 그림, 도형, OLE의 inline/floating 배치 1차 반영
- `.hwp`
  - 최소 본문/preview 기반 열람 지원
  - HWPX보다 재현도는 낮음

## 개발 환경에서 실행

루트 디렉터리에서 아래 순서로 실행합니다.

```bash
pnpm install
pnpm dev
```

실행되면 Tauri 데스크톱 앱이 열립니다.

웹 UI만 확인하려면 다음 명령을 사용할 수 있습니다.

```bash
pnpm dev:web
```

다만 실제 문서 열기는 Tauri 데스크톱 셸에서 사용하는 것이 맞습니다.

## macOS 앱 실행

이미 빌드된 앱 번들이 있으면 다음 경로의 앱을 실행하면 됩니다.

`/Users/magic/work/MAX-Viewer/target/release/bundle/macos/MAX-Viewer.app`

터미널에서 직접 열려면:

```bash
open /Users/magic/work/MAX-Viewer/target/release/bundle/macos/MAX-Viewer.app
```

앱 번들을 새로 만들려면 루트 디렉터리에서:

```bash
pnpm --dir apps/desktop-tauri tauri build --bundles app
```

## 문서 열기

앱 상단의 `파일 열기` 버튼을 누릅니다.

열 수 있는 형식:

- `.hwp`
- `.hwpx`

문서를 열면 다음 정보가 화면에 표시됩니다.

- 문서 이름
- 형식
- 현재 렌더링된 페이지 수
- 문단 수

이후 본문은 회색 작업 영역 위의 흰 종이 페이지 형태로 표시됩니다.

## 화면 사용 방법

상단 배율 버튼으로 보기 배율을 바꿀 수 있습니다.

- `폭 맞춤`: 현재 창 너비 기준으로 페이지 폭을 자동 맞춤
- `100%`
- `125%`
- `150%`

문서는 페이지 단위로 표시되며, 각 페이지에는 머리말/꼬리말과 쪽 번호가 반복 렌더링될 수 있습니다.

## 현재 동작 방식

### HWPX

- `content.hpf`에서 문서 순서를 읽습니다.
- `header.xml`에서 스타일, 번호, 시작 페이지 번호를 읽습니다.
- `section*.xml`에서 문단, 표, 그림/도형/OLE, 페이지 레이아웃을 읽습니다.
- `BinData`의 이미지 자산은 가능한 경우 실제 화면에 렌더링합니다.

### HWP

- `FileHeader`를 읽어 문서 여부를 검사합니다.
- 최소 `BodyText`와 `PrvText` 기반으로 본문 preview를 복원합니다.

## 제한 사항

현재 버전은 읽기 전용 뷰어이며, 아래 항목은 아직 완전하지 않습니다.

- HWP 문서의 정밀 조판 재현
- 한컴과 100% 동일한 페이지 수 계산
- 도형 회전, 클리핑, 효과의 완전 복원
- 표 자동 맞춤과 고급 편집 기능
- 각주/미주/주석의 완전 복원

즉, 현재 목표는 "안정적으로 열고 읽을 수 있는 뷰어"이지 "한글 편집기와 완전히 동일한 렌더러"는 아닙니다.

## 문제 해결

### 앱이 열리지 않을 때

먼저 의존성이 설치되어 있는지 확인합니다.

```bash
pnpm install
cargo check --workspace
```

### 문서가 열리지 않을 때

- 파일 확장자가 `.hwp` 또는 `.hwpx`인지 확인합니다.
- 암호화된 문서는 일부 내용만 표시될 수 있습니다.
- 손상된 문서는 파서 단계에서 오류가 날 수 있습니다.

### macOS에서 앱 번들이 최신 상태가 아닐 때

다시 빌드합니다.

```bash
pnpm --dir apps/desktop-tauri tauri build --bundles app
```

## 관련 문서

- [문서 인덱스](./README.md)
- [포맷 조사](./hwp-hwpx-format-research.md)
- [구현 계획](./viewer-implementation-plan.md)
- [HWPX 출력 품질 계획](./hwpx-hancom-viewer-fidelity-plan.md)

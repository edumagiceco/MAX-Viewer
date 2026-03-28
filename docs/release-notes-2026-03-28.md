# MAX-Viewer 업데이트 노트

작성일: 2026-03-28

## 핵심 변경

### 1. HWPX 렌더링 정합도 개선

- `keepWithNext`, `keepLines`, `lineSegArray` 높이 힌트를 이용한 pagination 보정
- 표 내부 블록 분할, rowSpan 보정, 헤더 행 반복 처리 강화
- floating 개체 footprint와 exclusion zone을 이용한 본문 감싸기 보정
- embedded table wrapper 문단 제거로 잘못 늘어나던 페이지 수 보정
- `noAdjust="0"` wrapper table의 자동 맞춤 처리로 macOS/WebKit 가로 스크롤과 외곽선 잘림 해결
- 표 셀 `borderFill`의 `slash`/`backSlash`/`diagonal` 파싱 및 SVG 오버레이 렌더링 추가

### 2. HWP 파서 확장

- `DocInfo`의 글꼴, 글자 모양, 문단 모양, 스타일, `BorderFill`, `BinData` 파싱 추가
- `BodyText`에서 표, 그림/도형, 페이지 설정, 머리말/꼬리말 복원 강화
- `PrvText` fallback 유지와 자산 base64 로드 연결

### 3. 데스크톱/웹 뷰어 기능 추가

- macOS Tauri 앱 번들 빌드 정리
- 파일 드래그 앤 드롭 열기
- 최근 문서 목록
- Ctrl+F 검색과 매칭 하이라이트
- Ctrl+G 페이지 이동, PageUp/PageDown/Home/End/Space 키보드 탐색
- 목차/썸네일 사이드바
- 다크 모드와 인쇄 스타일
- 10쪽 이상 문서 lazy rendering

### 4. 개발 환경 및 배포 보강

- CLI 앱 `apps/cli` 추가
- WASM 크레이트 `apps/wasm` 추가
- GitHub Actions CI 워크플로우 추가
- 구현 계획/사용 가이드 문서 보강

## 이번 릴리스에서 직접 해결한 대표 이슈

- HWPX 문서가 아래한글 11페이지인데 MAX Viewer에서 과도하게 많은 페이지로 보이던 문제
- macOS 앱에서 첫 페이지 안내 박스 테두리가 좌우로 잘리던 문제
- 제목 장식 표의 오른쪽 사선이 표시되지 않던 문제

## 남은 주요 과제

- 암호/DRM 문서 지원
- 복합 도형, 수식, 차트의 고충실도 렌더링
- 한컴 뷰어 수준의 쪽수/조판 완전 일치
- 모바일(Tauri Mobile) 확장

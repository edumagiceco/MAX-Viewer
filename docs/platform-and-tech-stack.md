# MAX-Viewer 플랫폼 및 기술 스택 추천

작성일: 2026-03-26

## 1. 결론

2026-03-26 기준으로 MAX-Viewer에는 아래 조합을 추천합니다.

- 코어: `Rust`
- 데스크톱 앱 셸: `Tauri 2`
- 프론트엔드 UI: `React + Vite`
- 장기 확장:
  - 웹: `Rust -> WASM` 재사용
  - 모바일: `Tauri 2 mobile`은 2차 단계에서 검토

핵심 판단은 단순합니다.

- HWP/HWPX 뷰어의 본질적인 난이도는 UI보다 `파일 파싱`, `스타일 해석`, `레이아웃 복원`에 있습니다.
- 이 핵심 로직은 Rust가 가장 잘 맞습니다.
- Tauri 2는 현재 `Windows`, `macOS`, `Linux`, `iOS`, `Android` 확장을 공식적으로 겨냥하고 있고, 프론트엔드는 웹 스택을 그대로 사용할 수 있습니다.
- 따라서 `Rust 문서 엔진 + Tauri 데스크톱 셸`이 개발 속도, 성능, 배포 크기, 확장성의 균형이 가장 좋습니다.

## 2. MAX-Viewer가 플랫폼 선택에서 중요하게 봐야 할 조건

이 프로젝트는 일반 CRUD 앱이 아니라 문서 뷰어입니다. 그래서 다음 조건이 중요합니다.

- 로컬 파일을 안전하게 읽어야 함
- 큰 문서를 빠르게 파싱해야 함
- 표, 이미지, 문단 스타일, 페이지 구분을 무리 없이 렌더링해야 함
- Windows/macOS/Linux에서 일관되게 동작해야 함
- 장기적으로 웹 또는 모바일 확장이 가능해야 함
- 설치 크기와 메모리 사용량이 과도하지 않아야 함
- 외부 파일을 여는 앱이므로 보안 경계가 분명해야 함

즉, "UI를 어디에 그릴까"보다 "문서 엔진을 어디에 둘까"가 더 중요합니다.

## 3. 후보별 비교

### 3.1 Tauri 2 + Rust

장점:

- Rust 코어와 자연스럽게 맞물림
- 프론트엔드는 React/Vue/Svelte 등 익숙한 웹 스택 사용 가능
- 데스크톱 앱 크기가 Electron보다 대체로 가벼운 편
- 공식적으로 모바일 확장 경로가 열려 있음
- 파일 시스템, 윈도우, 메뉴, 다이얼로그, 업데이트 등 데스크톱 기능과 결합하기 좋음

주의점:

- 웹뷰 기반이라 OS별 렌더링 엔진이 다름
- Linux에서는 `webkit2gtk` 의존성을 고려해야 함
- 모바일은 가능하지만, 아직 데스크톱만큼의 성숙도를 기대하면 안 됨

MAX-Viewer 적합도:

- 매우 높음

### 3.2 Flutter + Rust FFI

장점:

- Android/iOS/Windows/macOS/Linux/web를 폭넓게 지원
- 하나의 UI 코드베이스를 유지하기 쉬움
- 직접 그리는 UI라 OS별 뷰 차이가 비교적 적음

주의점:

- 문서 뷰어처럼 텍스트가 많은 읽기형 화면에서는 HTML/CSS 기반보다 유연성이 떨어질 수 있음
- 웹은 지원되지만, Flutter 공식 문서도 텍스트 중심 문서형 웹에는 최적이 아닐 수 있다고 설명함
- Rust 코어와의 FFI, 플랫폼별 빌드 체인 관리가 추가 비용

MAX-Viewer 적합도:

- 높음, 하지만 웹/문서형 UI 관점에서는 Tauri보다 한 단계 아래

### 3.3 Electron + Rust/Node 네이티브 모듈

장점:

- 웹 프론트엔드 개발 생산성이 높음
- 생태계가 매우 큼
- 데스크톱 앱 사례가 많음

주의점:

- 패키지 크기와 메모리 사용량이 상대적으로 큼
- 보안과 권한 경계 설계를 더 엄격히 해야 함
- 이 프로젝트의 핵심 가치가 "문서 파싱 엔진"인데, Electron 자체는 그 부분의 장점을 주지 않음

MAX-Viewer 적합도:

- 보통

### 3.4 Qt + C++/Rust 연동

장점:

- 데스크톱 앱 제어력과 렌더링 커스터마이징이 강함
- 성숙한 전통적 데스크톱 프레임워크
- 문서 뷰어 같은 앱에 잘 맞는 타입의 UI 제어 가능

주의점:

- 빌드와 배포 복잡도가 큼
- 팀이 Qt/C++에 익숙하지 않다면 생산성이 크게 떨어질 수 있음
- 라이선스 검토 비용이 생길 수 있음

MAX-Viewer 적합도:

- 기술적으로는 강하지만, 초기 제품 개발 속도 기준으로는 과함

### 3.5 .NET MAUI

장점:

- Microsoft 생태계와 궁합이 좋음
- Windows/macOS/iOS/Android 지원

주의점:

- Linux를 공식 지원하지 않음
- MAX-Viewer처럼 Linux 데스크톱까지 포함하는 범용 뷰어에는 맞지 않음

MAX-Viewer 적합도:

- 낮음

## 4. 왜 Tauri 2를 기본 추천으로 보는가

### 4.1 현재 공식 지원 방향과 구조가 맞다

Tauri 2.0 stable 발표 기준으로 Tauri는 단일 UI 코드베이스를 유지하면서 `Windows`, `macOS`, `Linux`, `Android`, `iOS`까지 겨냥하는 방향을 분명히 제시하고 있습니다. 또한 프론트엔드는 웹 기술로 작성하고, 애플리케이션 코어는 Rust로 구성하는 구조를 기본 모델로 둡니다.

이 구조는 MAX-Viewer와 잘 맞습니다.

- 문서 엔진은 Rust
- UI는 웹 기술
- 시스템 연동은 Tauri 플러그인/명령 인터페이스

### 4.2 데스크톱 우선 전략에 적합하다

MAX-Viewer의 실사용 가치는 우선 데스크톱에서 큽니다. 공공문서, 회사 문서, 대용량 문서 열람은 모바일보다 데스크톱 수요가 높습니다.

Tauri는 이 데스크톱 우선 전략에 잘 맞습니다.

- 파일 열기
- 최근 문서
- 드래그 앤 드롭
- 로컬 캐시
- 네이티브 메뉴
- 다중 창

이런 기능을 구현하기 쉽습니다.

### 4.3 코어를 고정해 두면 나중에 웹과 모바일이 쉬워진다

처음부터 모든 플랫폼 UI를 한 번에 맞추려 하면 일정이 커집니다. 대신 아래 순서가 더 현실적입니다.

1. `Rust 코어`를 만든다.
2. `Tauri 데스크톱`으로 제품 가치를 먼저 검증한다.
3. 그 뒤 같은 코어를 `WASM` 또는 `Tauri mobile`로 재사용한다.

즉, 프레임워크보다 `코어 재사용 구조`가 중요하고, Tauri는 그 첫 번째 셸로 적합합니다.

## 5. Tauri 2를 쓸 때 알아야 할 실제 제약

### 5.1 웹뷰 엔진 차이

Tauri 공식 문서 기준으로:

- Windows: `WebView2`
- macOS/iOS: `WKWebView`
- Linux: `WebKitGTK`

이 의미는 OS마다 HTML/CSS 렌더링 특성이 조금씩 다를 수 있다는 뜻입니다.

대응 원칙:

- 문서 레이아웃 계산은 Rust 코어에서 최대한 공통화
- UI는 결과 표시와 상호작용에 집중
- 페이지 정확도가 중요한 부분은 CSS에만 의존하지 말고 측정/배치 로직을 분리

### 5.2 모바일은 "가능하지만 1차 목표로 잡지 않는 편"이 낫다

Tauri 2.0 stable 글은 모바일 지원을 공식 기능으로 소개하면서도, 개발 경험을 데스크톱 수준으로 계속 끌어올리는 중이고 일부 공식 플러그인은 아직 모바일을 지원하지 않는다고 설명합니다.

실무적으로는 이렇게 해석하는 것이 타당합니다.

- 모바일 확장 가능성은 충분함
- 하지만 MAX-Viewer 1차 제품은 데스크톱 우선이 맞음
- 모바일은 2차 단계에서 읽기 모드 중심으로 재사용 검토

## 6. Flutter를 바로 선택하지 않는 이유

Flutter는 나쁜 선택이 아닙니다. 오히려 아래 조건이면 유력합니다.

- 모바일이 1차 시장이다
- UI를 거의 직접 그려야 한다
- 플랫폼별 동일한 시각 결과가 아주 중요하다

하지만 MAX-Viewer는 우선 문서 뷰어입니다. Flutter 공식 웹 문서는 텍스트가 많고 흐름형인 문서성 콘텐츠는 웹의 문서 중심 모델이 더 자연스러운 경우가 있다고 명시합니다. MAX-Viewer는 검색, 복사, 텍스트 선택, 링크, 긴 본문, 표가 중요하기 때문에 HTML/CSS 기반 읽기 모드가 더 유리합니다.

그래서 현재 시점에서는:

- 데스크톱 우선: `Tauri`
- 모바일 우선: `Flutter`

로 보는 것이 더 현실적입니다.

## 7. 권장 기술 스택

### 7.1 최종 추천 스택

- 언어: `Rust`, `TypeScript`
- 앱 셸: `Tauri 2`
- 프론트엔드: `React + Vite`
- 상태 관리: 가볍게 시작해서 `Zustand` 정도
- 문서 엔진: Rust workspace
- 렌더링:
  - 읽기 모드: `HTML/CSS`
  - 페이지 모드: `Canvas/SVG + HTML 혼합` 또는 점진적 고도화
- 테스트:
  - Rust 단위 테스트
  - fixture 기반 골든 테스트
  - Playwright 기반 UI 스냅샷 테스트

### 7.2 저장소 구조 권장안

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
├── packages/
│   └── viewer-ui/
├── fixtures/
└── docs/
```

### 7.3 역할 분리

- Rust:
  - 파일 파서
  - 내부 문서 모델
  - 스타일 해석
  - 텍스트 추출
  - 레이아웃 계산
- React/Tauri:
  - 파일 선택 UI
  - 문서 리스트
  - 뷰어 화면
  - 검색 패널
  - 썸네일/네비게이션
  - 설정 UI

## 8. 단계별 플랫폼 전략

### Phase 1. 데스크톱 MVP

대상:

- Windows
- macOS
- Linux

구성:

- Rust 코어
- Tauri 2 데스크톱 앱
- HWPX 우선 지원

목표:

- 공공문서와 일반 문서를 안정적으로 열람

### Phase 2. 웹 미리보기/임베디드 뷰어

대상:

- 브라우저

구성:

- Rust 코어 일부를 WASM으로 빌드
- 읽기 모드 중심의 웹 뷰어 제공

목표:

- 링크 공유, 사내 포털 임베드, 웹 미리보기

주의:

- HWP 전체 파서를 브라우저로 바로 옮길지는 성능과 파일 크기를 보고 결정

### Phase 3. 모바일 뷰어

대상:

- iOS
- Android

구성:

- Tauri mobile 또는 별도 모바일 셸
- 읽기 모드 중심 UI

목표:

- 외부에서 문서 열람
- 검토/확인용 간편 뷰어

## 9. 최종 추천

이 프로젝트는 아래 선택이 가장 안정적입니다.

1. `Rust`로 문서 코어를 만든다.
2. `Tauri 2`로 데스크톱 앱을 먼저 만든다.
3. UI는 `React + TypeScript`로 빠르게 구축한다.
4. 웹과 모바일은 같은 코어를 재사용하는 2차 확장으로 둔다.

한 줄로 줄이면:

`MAX-Viewer는 "멀티플랫폼 UI 프레임워크"보다 "Rust 문서 엔진을 중심으로 셸을 늘려 가는 구조"가 맞고, 그 첫 셸로는 Tauri 2가 가장 적절합니다.`

## 10. 참고 자료

- Tauri 2.0 Stable Release  
  <https://v2.tauri.app/blog/tauri-20/>
- Tauri Webview Versions  
  <https://v2.tauri.app/reference/webview-versions/>
- Flutter Supported deployment platforms  
  <https://docs.flutter.dev/reference/supported-platforms>
- Flutter Web support  
  <https://docs.flutter.dev/platform-integration/web>
- Flutter WebAssembly support  
  <https://docs.flutter.dev/platform-integration/web/wasm>
- Electron official site  
  <https://www.electronjs.org/>
- Electron Packager supported platforms  
  <https://packages.electronjs.org/packager/>
- Qt supported platforms  
  <https://doc.qt.io/qt-6.8/supported-platforms.html>
- .NET MAUI supported platforms  
  <https://learn.microsoft.com/en-us/dotnet/maui/supported-platforms?view=net-maui-10.0>

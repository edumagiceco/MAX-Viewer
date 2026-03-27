# MAX-Viewer 문서

작성일: 2026-03-27

## 프로젝트 목적

MAX-Viewer의 목표는 `.hwp`와 `.hwpx` 문서를 안정적으로 읽고 표시할 수 있는 뷰어 프로그램을 만드는 것입니다. 이 프로젝트는 편집기보다 먼저 "열기", "보기", "검색", "텍스트 추출", "표/이미지 표시"에 집중하는 읽기 전용 제품을 우선 목표로 둡니다.

조사 결과를 기준으로 보면, 구현 우선순위는 다음이 가장 현실적입니다.

1. `HWPX`를 먼저 안정적으로 지원한다.
2. `HWP`는 별도 바이너리 어댑터로 지원한다.
3. 두 포맷을 바로 화면에 그리지 말고, 공통 내부 문서 모델로 변환한 뒤 렌더링한다.
4. 초기 버전은 "완전한 편집 호환"이 아니라 "신뢰할 수 있는 열람 호환"에 집중한다.

한컴 공개 자료에 따르면 한컴은 2010년부터 HWP 구조 공개와 HWPX(OWPML) 지원을 시작했고, 2021년에 HWPX를 기본 저장 포맷으로 전환했습니다. 따라서 장기적으로는 두 포맷을 모두 읽어야 하지만, 실제 개발 순서는 `HWPX 우선`이 맞습니다.

## 문서 구성

- [hwp-hwpx-format-research.md](./hwp-hwpx-format-research.md): HWP/HWPX 포맷 조사, 공식 자료 요약, 구현 시사점
- [viewer-implementation-plan.md](./viewer-implementation-plan.md): 실제 구현 방향, 권장 아키텍처, 단계별 개발 계획
- [platform-and-tech-stack.md](./platform-and-tech-stack.md): 멀티플랫폼 개발 프레임워크 비교와 최종 권장 스택
- [hwpx-hancom-viewer-fidelity-plan.md](./hwpx-hancom-viewer-fidelity-plan.md): 한컴뷰어와 유사한 HWPX 출력 품질을 목표로 한 세부 개선 계획

## 지금 바로 착수할 개발 순서

1. `HWPX -> 내부 문서 모델 -> 페이지 렌더링`에서 한컴뷰어와의 시각 차이를 줄인다.
2. `header.xml` 스타일 테이블과 `section.xml`의 스타일 참조를 해석한다.
3. 페이지 크기, 여백, 문단 모양, 글자 모양을 반영한다.
4. 이후 `HWP -> 내부 문서 모델` 어댑터를 같은 렌더러에 연결한다.
5. 마지막으로 번호, 각주, 도형, OLE, 페이지 나눔 정확도를 끌어올린다.

## 현재 구현 상태

2026-03-27 기준 현재 저장소는 다음 상태다.

- HWPX: ZIP/`content.hpf`/`header.xml`/`section*.xml`를 읽고 `charPr`/`paraPr`/`style` 참조를 1차 적용해 페이지 윤곽형 화면으로 렌더링
- HWPX 렌더링 보정: 스타일 경계 공백, 빈 문단, 표 셀 줄바꿈 보존
- HWPX 레이아웃: 번호 매기기/글머리표 1차 표시, 머리말/꼬리말 반복 렌더링, 화면 측정 기반 실제 페이지 분할
- HWPX 조판 보정: `lineSegArray` 줄바꿈 반영, 번호 폭/오프셋 기반 hanging indent 근사, `charPr` 메트릭 기반 글자 폭/자간/베이스라인 1차 보정
- HWPX 표/개체: 표 셀 내부 블록 조판, `cellSpan`/`cellSz`/`cellMargin` 해석, 그림/도형/OLE의 inline/floating 배치 1차 지원
- HWP: `FileHeader`, 최소 `BodyText`, `PrvText` fallback preview 지원
- Desktop shell: Tauri 2 기반 macOS 앱 번들 실행 가능
- UI: 파일 열기 중심의 단순 뷰어 화면, 구역별 용지/여백 반영, 기본 배율 제어 완료
- 현재 한계: 조판 엔진은 아직 근사치이며, 도형 회전/효과/클리핑, 표 자동 맞춤, 한컴 수준의 쪽수 완전 일치까지는 미도달
- 다음 우선순위: 페이지 분할 정확도와 표/개체의 페이지 경계 처리, HWP DocInfo/control 복원도를 높이기

## 현재 권장 플랫폼 전략

2026-03-26 기준으로는 `Rust 코어 + Tauri 2 데스크톱 앱`을 기본 선택으로 두는 것이 가장 현실적입니다.

- 1차 제품: `Windows`, `macOS`, `Linux`
- 2차 확장: 같은 Rust 코어를 `Tauri Mobile(iOS/Android)` 또는 `WASM 기반 웹 뷰어`로 재사용
- UI 전략: 프론트엔드는 웹 기술, 파서/문서 모델/핵심 로직은 Rust

## 참고한 핵심 자료

- 한컴 FAQ: HWP 공개 및 HWPX 기본 저장 포맷 전환  
  <https://recruit.hancom.co.kr/support/faqCenter/faq/detail/3129>
- 한컴테크: HWP 포맷 구조  
  <https://tech.hancom.com/%ED%95%9C-%EA%B8%80-%EB%AC%B8%EC%84%9C-%ED%8C%8C%EC%9D%BC-%ED%98%95%EC%8B%9D-hwp-%ED%8F%AC%EB%A7%B7-%EA%B5%AC%EC%A1%B0-%EC%82%B4%ED%8E%B4%EB%B3%B4%EA%B8%B0/>
- 한컴테크: HWPX 포맷 구조  
  <https://tech.hancom.com/hwpxformat/>
- 한컴 공식 HWP 5.0 문서 형식 PDF  
  <https://cdn.hancom.com/link/docs/%ED%95%9C%EA%B8%80%EB%AC%B8%EC%84%9C%ED%8C%8C%EC%9D%BC%ED%98%95%EC%8B%9D_5.0_revision1.2.pdf>
- 한컴 공개 OWPML 모델 저장소  
  <https://github.com/hancom-io/hwpx-owpml-model>

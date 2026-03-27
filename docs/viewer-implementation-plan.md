# MAX-Viewer 구현 방향

작성일: 2026-03-27

## 0. 현재 상태

2026-03-27 기준 현재 저장소에는 다음 구현이 들어가 있다.

- `max_viewer_hwpx`: ZIP 컨테이너, `content.hpf`, `header.xml`, `section.xml` 파싱과 `charPr`/`paraPr`/`style` 기반 1차 스타일 복원
- `max_viewer_hwp`: `FileHeader`, 최소 `BodyText` 문단 복원, `PrvText` fallback
- `packages/viewer-ui`: 파일 열기, 페이지 윤곽형 뷰, 기본 배율 제어가 있는 읽기 전용 UI
- `apps/desktop-tauri`: 네이티브 파일 다이얼로그와 macOS `.app` 번들 실행

다만 현재 HWPX 렌더링은 `구역별 용지/여백 기반 페이지 윤곽`과 `기본 문단/글자 스타일`까지만 반영한다. 실제 쪽 나눔, 머리말/꼬리말, 번호 매기기, 도형 배치는 아직 미구현이므로 다음 구현 중심은 이 격차를 줄이는 데 있다.

## 1. 제품 목표

MAX-Viewer는 `.hwp`, `.hwpx` 문서를 읽기 전용으로 안정적으로 열어 보여주는 프로그램을 목표로 합니다.

초기 제품 목표:

- 파일 열기
- 페이지/읽기 모드 보기
- 텍스트 검색
- 복사
- 표, 이미지, 기본 스타일 표시
- 메타데이터 표시

초기 비목표:

- 문서 편집
- 고급 수식/차트/OLE 완전 재현
- 암호/DRM 문서 전체 지원
- 한컴 편집기 수준의 100% 레이아웃 일치

## 2. 권장 아키텍처

가장 현실적인 구조는 `Rust 기반 문서 코어 + UI 셸`입니다.

이 방향을 추천하는 이유:

- HWP는 바이너리 파싱 비중이 크므로 메모리 안정성과 성능이 중요합니다.
- HWPX는 ZIP/XML 기반이라 Rust로 처리해도 단순합니다.
- 같은 코어를 데스크톱 앱, CLI, 서버, WASM에 재사용할 수 있습니다.
- 실제 공개 구현체 중 `hwpjs`도 Rust 코어 + Web 바인딩 구조를 사용합니다.

### 추천 구성

- `core`
  - 문서 모델
  - 스타일 모델
  - 공통 에러/로깅
- `parser-hwpx`
  - ZIP/XML 기반 HWPX 파서
- `parser-hwp`
  - CFB + 레코드 기반 HWP 파서
- `renderer-html`
  - 내부 문서 모델을 HTML/CSS로 렌더링
- `renderer-page`
  - 페이지 단위 배치 렌더링
- `app-desktop`
  - Tauri 또는 다른 데스크톱 셸
- `app-web`
  - 브라우저 뷰어 또는 개발용 샌드박스

## 3. 저장소 구조 제안

저장소가 아직 비어 있으므로 처음부터 아래 구조로 가는 것을 추천합니다.

```text
MAX-Viewer/
├── docs/
├── crates/
│   ├── max_viewer_core/
│   ├── max_viewer_hwpx/
│   ├── max_viewer_hwp/
│   ├── max_viewer_render_html/
│   └── max_viewer_render_page/
├── apps/
│   ├── desktop/
│   └── web/
├── fixtures/
│   ├── hwpx/
│   ├── hwp/
│   ├── expected-json/
│   └── expected-html/
└── tools/
```

핵심 원칙은 `포맷 파서`, `문서 모델`, `렌더러`, `UI`를 분리하는 것입니다.

## 4. 내부 문서 모델을 먼저 정의해야 하는 이유

포맷마다 표현 방식은 다르지만, 뷰어에 필요한 정보는 대부분 비슷합니다. 따라서 먼저 공통 모델을 정의해야 합니다.

### 권장 문서 모델

```text
Document
├── metadata
├── sections[]
│   ├── page_settings
│   ├── header/footer
│   └── blocks[]
│       ├── Paragraph
│       ├── Table
│       ├── Image
│       ├── Shape
│       ├── Control
│       └── PageBreak
├── styles
│   ├── char_styles[]
│   ├── para_styles[]
│   ├── table_styles[]
│   └── numbering[]
└── assets
    ├── images
    └── embedded_objects
```

### Paragraph 구조 예시

```text
Paragraph
├── para_style_id
├── runs[]
│   ├── TextRun
│   ├── LineBreak
│   ├── Tab
│   ├── FieldStart/End
│   └── InlineObjectRef
├── align
├── indent
├── spacing
└── controls[]
```

이 모델을 두 포맷이 공유하면:

- 렌더러를 한 번만 만들 수 있고
- 검색, 복사, 미리보기, 목차, 내보내기 기능도 공통으로 구현할 수 있습니다.

## 5. HWPX 구현 계획

HWPX는 먼저 완성해야 할 1순위 대상입니다.

### 5.1 파싱 순서

1. ZIP 컨테이너 열기
2. `mimetype` 검사
3. `version.xml`, `settings.xml` 읽기
4. `Contents/content.hpf` 파싱
5. `manifest`와 `spine` 인덱스 생성
6. `Contents/header.xml` 파싱
7. `Contents/sectionN.xml`들을 `spine` 순서대로 파싱
8. `BinData` 로드
9. 내부 문서 모델로 변환

### 5.2 구현 시 포인트

- XML 네임스페이스를 엄격하게 처리해야 합니다.
- `header.xml`의 스타일 매핑이 먼저 메모리에 올라와야 `sectionN.xml`의 참조를 해석할 수 있습니다.
- `<hp:t>`만 추출하는 단순 텍스트 모드와, 스타일/레이아웃까지 반영하는 정식 파서를 분리하면 개발 속도가 빨라집니다.
- `content.hpf`의 `spine` 순서를 무시하면 문서 순서가 틀어질 수 있습니다.
- `BinData`는 lazy-loading으로 두는 편이 메모리 효율상 유리합니다.

### 5.3 HWPX 1차 지원 범위

- 문단
- 글자 스타일
- 문단 스타일
- 줄바꿈/탭
- 표
- 이미지
- 페이지 구분
- 머리글/바닥글의 기본 텍스트
- 메타데이터

### 5.4 HWPX 2차 지원 범위

- 번호 매기기/글머리표
- 각주/미주
- 도형 일부
- 링크/필드
- 변경 추적/주석 일부 표시

## 6. HWP 구현 계획

HWP는 난이도가 높기 때문에 파서를 더 작은 단계로 쪼개야 합니다.

### 6.1 파싱 파이프라인

1. CFB 컨테이너 열기
2. 디렉터리 엔트리 인덱스 구성
3. `FileHeader` 읽기
4. 버전, 압축, 암호, 배포용 문서 여부 확인
5. `DocInfo` 스트림 읽기
6. 압축 여부에 따라 해제
7. 레코드 반복 파싱
8. 스타일/자산 매핑 테이블 구축
9. `BodyText/SectionN` 스트림 반복 파싱
10. 문단/표/컨트롤 계층 복원
11. `BinData` 연결
12. 내부 문서 모델로 변환

### 6.2 HWP 파서에서 꼭 필요한 서브모듈

- `cfb_reader`
  - OLE/CFB 엔트리 읽기
- `compression`
  - 스트림 압축 해제
- `record_reader`
  - 32비트 레코드 헤더 해석
- `docinfo_parser`
  - 글꼴, 문단 모양, 스타일, BinData 참조 복원
- `body_parser`
  - 문단, 텍스트, 컨트롤, 표, 페이지 설정 파싱
- `asset_loader`
  - `BinData` 추출

### 6.3 HWP 파싱 시 주의점

- 버전별 레코드 길이 차이를 허용해야 합니다.
- 모르는 필드는 에러로 중단하지 말고 남은 바이트를 스킵해야 합니다.
- 압축 문서와 비압축 문서를 모두 지원해야 합니다.
- 암호 문서와 DRM 문서는 처음부터 완전 지원을 목표로 잡지 않는 편이 맞습니다.
- `PrvText`, `PrvImage`는 본문 파싱 실패 시 fallback preview로 활용할 수 있습니다.

### 6.4 초기 HWP 지원 범위

- 일반 텍스트 문단
- 기본 문자/문단 스타일
- 표
- 이미지
- 페이지/구역 구분

나중으로 미룰 항목:

- 배포용 문서 특수 처리
- 복합 도형
- 스크립트
- OLE의 실제 렌더링

## 7. 렌더링 전략

뷰어는 처음부터 한 가지 렌더러만 두지 않는 편이 좋습니다. `읽기 모드`와 `페이지 모드`를 나누는 방향을 추천합니다.

### 7.1 읽기 모드

특징:

- HTML/CSS 기반
- 빠름
- 접근성 좋음
- 검색/복사/반응형 UI에 유리

적합한 용도:

- 문서 내용 확인
- 텍스트 검색
- 모바일/웹 뷰

### 7.2 페이지 모드

특징:

- 페이지 크기, 여백, 머리글/바닥글, 줄 배치를 더 엄격하게 반영
- 정확도는 높지만 구현 난이도와 비용이 큼

적합한 용도:

- 인쇄 유사 보기
- 공문서 레이아웃 확인

### 7.3 현실적인 구현 순서

1. 내부 문서 모델 -> 읽기 모드 HTML 렌더링
2. 페이지 구분만 반영한 기본 페이지 모드
3. 줄 배치/표 높이/머리글/바닥글 정밀도 개선
4. 이후 정확도 테스트를 기반으로 보정

## 8. 폰트 전략

문서 렌더링 품질은 폰트에서 크게 갈립니다.

### 권장 원칙

- 문서에 지정된 폰트명을 그대로 존중하되, 없는 경우 대체 폰트 매핑 테이블 사용
- 기본 대체 폰트는 한국어 가독성이 높은 계열로 선정
- 한컴 제공 글꼴 사용 여부는 라이선스를 먼저 확인
- 동일 문서를 플랫폼별로 비슷하게 보이게 하려면 자체 폰트 번들을 검토

### 구현 포인트

- 폰트 매핑 테이블을 설정 파일로 분리
- 문단/글자 스타일 적용 시 폰트 fallback 체인 지원
- 텍스트 폭 측정은 페이지 모드 품질에 직접 영향

## 9. 자산과 복합 객체 처리

### 이미지

- `BinData`와 문서 내 참조를 연결
- 원본 바이너리를 캐시하고 렌더러에서 URL 또는 메모리 핸들로 사용

### OLE

- 초기에는 실제 실행/렌더링 대신 placeholder로 표시
- 아이콘, 파일명, 대체 텍스트, 크기 정보만 먼저 보여주는 방식이 현실적

### 도형/차트/수식

- 초기 버전에서는 완전 재현보다 "깨지지 않게 대체 표시"가 더 중요
- 지원하지 않는 타입은 warning을 남기고 bounding box placeholder를 그리면 됩니다

## 10. 에러 처리와 보안

문서 뷰어는 외부 파일을 직접 읽기 때문에 입력 안전성이 중요합니다.

### 최소 요구사항

- ZIP bomb 방지
- CFB stream size 상한
- XML entity/DTD 확장 차단
- 비정상 레코드 길이 검증
- 이미지 디코딩 실패 격리
- 악성 스크립트 미실행

### 권장 정책

- "파일을 못 열면 앱이 죽는 구조"를 피하고, 섹션 단위/자산 단위로 오류를 격리
- 파서 에러를 사용자 메시지와 내부 로그로 분리
- fallback preview를 제공할 수 있으면 제공

## 11. 테스트 전략

이 프로젝트는 테스트 코퍼스가 품질을 결정합니다.

### 꼭 필요한 샘플 문서 종류

- 순수 텍스트
- 다양한 글꼴/스타일
- 표가 많은 문서
- 이미지 포함 문서
- 머리글/바닥글 포함 문서
- 각주/번호 매기기 포함 문서
- 도형/OLE 포함 문서
- 압축 HWP
- 비압축 HWP
- 손상 파일
- 암호/배포용 문서 샘플

### 테스트 종류

- 단위 테스트
  - 레코드 헤더 파싱
  - XML 노드 매핑
  - 스타일 참조 복원
- 골든 테스트
  - 문서 -> 내부 JSON 비교
  - 문서 -> HTML 결과 비교
- 스냅샷 테스트
  - 대표 문서 화면 캡처 비교
- 회귀 테스트
  - 실패했던 문서를 fixtures에 계속 추가

### 실무 팁

- 가능하면 같은 문서를 HWP와 HWPX로 각각 저장해 비교 샘플로 사용
- 한컴 뷰어 출력과 MAX-Viewer 출력을 눈으로 비교할 기준 스냅샷을 남겨야 합니다.

## 12. 단계별 로드맵

### Phase 0. 문서/코퍼스 준비

- 공식 문서 정리
- 샘플 문서 수집
- fixture 구조 생성

### Phase 1. HWPX MVP

- ZIP/XML 로더
- `content.hpf`, `header.xml`, `sectionN.xml` 파서
- 내부 문서 모델
- 읽기 모드 HTML 렌더러

완료 기준:

- 일반적인 HWPX 문서에서 텍스트/표/이미지 열람 가능

### Phase 2. HWP MVP

- CFB 리더
- `FileHeader`, `DocInfo`, `BodyText` 파서
- 공통 모델 변환

완료 기준:

- 일반적인 HWP 문서에서 텍스트/표/이미지 열람 가능

### Phase 3. 페이지 품질 개선

- 페이지 모드
- 머리글/바닥글
- 번호 매기기
- 각주/미주

### Phase 4. 호환성 확대

- 도형 일부
- OLE placeholder
- 에러 복원력 강화
- 폰트 대체 고도화

## 13. 최종 권장 방향

MAX-Viewer는 아래 원칙으로 시작하는 것이 가장 안전합니다.

1. `HWPX 우선`
2. `HWP는 별도 파서`
3. `공통 내부 문서 모델`
4. `읽기 모드 먼저, 페이지 모드는 나중`
5. `완전 재현보다 열람 안정성 우선`
6. `샘플 문서와 회귀 테스트 중심 개발`

이 방향이면 개발 범위를 통제하면서도, 실제로 사용 가능한 HWP/HWPX 뷰어로 빠르게 진입할 수 있습니다.

## 14. 참고 자료

- 한컴 FAQ  
  <https://recruit.hancom.co.kr/support/faqCenter/faq/detail/3129>
- 한컴테크 HWP 포맷 구조  
  <https://tech.hancom.com/%ED%95%9C-%EA%B8%80-%EB%AC%B8%EC%84%9C-%ED%8C%8C%EC%9D%BC-%ED%98%95%EC%8B%9D-hwp-%ED%8F%AC%EB%A7%B7-%EA%B5%AC%EC%A1%B0-%EC%82%B4%ED%8E%B4%EB%B3%B4%EA%B8%B0/>
- 한컴테크 HWPX 포맷 구조  
  <https://tech.hancom.com/hwpxformat/>
- 한컴 공식 HWP 5.0 문서 형식 PDF  
  <https://cdn.hancom.com/link/docs/%ED%95%9C%EA%B8%80%EB%AC%B8%EC%84%9C%ED%8C%8C%EC%9D%BC%ED%98%95%EC%8B%9D_5.0_revision1.2.pdf>
- Hancom OWPML model  
  <https://github.com/hancom-io/hwpx-owpml-model>
- pyhwp  
  <https://github.com/mete0r/pyhwp>
- hwpjs  
  <https://github.com/ohah/hwpjs>
- LibreOffice hwpfilter  
  <https://docs.libreoffice.org/hwpfilter.html>

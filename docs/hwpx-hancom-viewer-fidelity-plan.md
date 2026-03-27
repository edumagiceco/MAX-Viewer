# HWPX 한컴뷰어 유사 출력 개선 계획

작성일: 2026-03-27

## 목표

MAX-Viewer의 HWPX 출력이 "텍스트를 읽을 수 있는 수준"을 넘어서, 한컴뷰어가 보여주는 기본 문서 인상에 최대한 가깝게 보이도록 개선한다.

이번 단계의 목표는 다음과 같다.

1. 쪽 윤곽이 보이는 페이지 기반 화면으로 전환
2. `header.xml` 스타일 테이블을 읽어 문단/글자 모양을 반영
3. `section.xml`의 구역별 페이지 크기와 여백을 반영
4. 한컴뷰어와 비교했을 때 가장 눈에 띄는 차이인 정렬, 들여쓰기, 줄 간격, 글자 크기, 굵게/기울임, 글자색을 우선 맞춤

이번 단계의 비목표는 다음과 같다.

- 각주/미주/주석 완전 재현
- 한컴 내부 조판 엔진과 바이트 단위로 동일한 계산
- 편집기 수준의 객체 편집, 드래그, 재배치

## 현재 반영 상태

2026-03-27 현재 저장소에는 다음 1차 구현이 반영되어 있다.

- `header.xml`의 `charPr`, `paraPr`, `style` 참조를 읽어 기본 문단/글자 모양을 공통 모델로 변환
- `section.xml`의 `secPr/pagePr/margin`을 읽어 구역별 용지 크기와 여백을 반영
- 프론트엔드를 회색 작업 배경 위 흰 종이 페이지 형태로 바꾸고, 글자 크기/정렬/간격/색/굵게/기울임을 1차 적용
- 기본 배율 제어를 제공해 문서 폭 기준으로 보는 흐름을 맞춤
- 스타일 경계 공백, 빈 문단, 표 셀 내부 줄바꿈을 보존해 서식 붕괴를 줄임
- `numberings`/`bullets`를 읽어 번호/글머리표를 1차 표시
- `header`/`footer`와 `pageNum` placeholder를 읽어 각 페이지에 반복 렌더링
- 화면에서 측정한 블록 높이를 기준으로 실제 페이지를 여러 장으로 분할
- `lineSegArray`의 줄 시작 위치를 읽어 브라우저 줄바꿈을 한컴 줄 나눔에 더 가깝게 보정
- 번호 `widthAdjust`와 `textOffset`을 읽어 번호 폭과 hanging indent를 근사
- `charPr`의 `ratio`, `spacing`, `relSz`, `offset`, `useKerning`, `useFontSpace`를 읽어 글자 폭/자간/베이스라인 보정을 화면에 반영
- 표 셀을 단순 문자열이 아니라 내부 블록 목록으로 복원하고 `cellSpan`, `cellSz`, `cellMargin`을 반영
- 그림, 도형, OLE의 `sz`, `pos`, `textWrap`, `zOrder`를 읽어 inline/floating 배치를 구분
- `BinData`를 data URI로 전달해 HWPX 그림을 실제 이미지로 렌더링
- 렌더링된 실제 페이지 수를 기준으로 머리말/꼬리말 placeholder와 총 쪽 수를 계산

아직 구현되지 않은 부분은 다음과 같다.

- 한컴 편집기와 정확히 같은 쪽 수 계산
- 번호 폭, hanging indent, 다단계 번호 패턴의 세밀한 재현
- 표 자동 맞춤, 셀 병합/분할 편집, 캡션 줄바꿈 등 고급 표 기능
- 그림, 도형, 수식, OLE의 회전/클리핑/효과까지 포함한 정밀 배치

## 이번 단계 구현 계획

이번 단계는 아래 네 축을 함께 올린다.

1. `header.xml`
   - `beginNum`
   - `refList/numberings`
   - `refList/bullets`
   를 읽어 문단 번호와 페이지 시작 번호를 복원한다.
2. `section.xml`
   - `ctrl/header`
   - `ctrl/footer`
   - `pageNum`, `autoNum`
   를 읽어 머리말/꼬리말과 페이지 번호 placeholder를 복원한다.
3. `section.xml` 개체/표 구조
   - `tbl/tr/tc/subList`
   - `cellSpan`, `cellSz`, `cellMargin`
   - `pic`, `ole`, 기본 도형 계열의 `sz`, `pos`
   를 읽어 셀 내부 블록과 개체 배치 힌트를 복원한다.
4. 프론트엔드
   - 구역당 1쪽 표시를 제거
   - 화면에서 측정한 블록 높이로 실제 페이지를 여러 장으로 분할
   - 각 페이지에 머리말/꼬리말을 반복 렌더링
   - floating 객체는 페이지 본문 위 절대 위치 레이어에 렌더링
   - 표 셀 안에서도 문단/표/개체를 다시 렌더링

이번 단계의 구현 원칙:

- 조판 엔진은 한컴과 100% 동일하게 재현하지 않는다.
- 대신 HWPX가 제공하는 글자/줄/개체 힌트와 현재 화면에서 측정된 실제 렌더링 높이를 함께 사용한다.
- 문단 번호는 `heading(type/idRef/level)`와 `numberings/bullets`를 연결해 계산한다.
- 페이지 번호는 문서의 실제 렌더링 결과 기준 쪽 번호를 사용한다.
- 총 쪽 수 placeholder는 렌더링이 끝난 뒤 실제 페이지 개수로 치환한다.

## 조사 결과 요약

### 1. 한컴뷰어/한글의 문서 보기 특징

한컴 공식 도움말 기준으로 한글의 문서 보기는 단순 스크롤 영역이 아니라 `쪽 윤곽(page outline)` 중심이다.

- `쪽 윤곽`을 켜면 인쇄될 실제 용지 여백, 머리말/꼬리말, 쪽 테두리 등 페이지에 들어갈 요소를 화면에서 바로 볼 수 있다.
- `쪽 맞춤`, `폭 맞춤`, `한 쪽`, `두 쪽`, `맞쪽`, `여러 쪽` 같은 배율/배치 모드가 존재한다.
- `여러 쪽 보기`를 선택하면 자동으로 `쪽 윤곽`이 켜진다.
- 머리말/꼬리말도 쪽 윤곽 상태에서 페이지에 배치된 형태로 보인다.

즉, MAX-Viewer도 단순한 "본문 블록 목록"이 아니라 최소한 다음을 제공해야 한다.

- 회색 작업 배경 위의 흰 종이 페이지
- 문서 폭에 맞춘 페이지 폭
- 페이지 내부 여백을 반영한 본문 영역
- 페이지 단위로 구역을 시각적으로 분리하는 레이아웃

### 2. HWPX에서 한컴뷰어 유사 출력에 직접 필요한 구조

한컴 공개 `hwpx-owpml-model`과 한컴테크 HWPX 포맷 설명을 기준으로 보면, 한컴뷰어 유사 출력에 필요한 핵심 구조는 이미 HWPX에 분리되어 있다.

- `Contents/header.xml`
  - `refList/fontfaces`
  - `refList/charProperties/charPr`
  - `refList/paraProperties/paraPr`
  - `refList/styles/style`
- `Contents/sectionN.xml`
  - 문단 `p`는 `paraPrIDRef`, `styleIDRef`를 가진다.
  - 런 `run`은 `charPrIDRef`를 가진다.
  - 구역 정의 `secPr` 아래 `pagePr`가 있고, 여기서 페이지 너비/높이/방향/여백을 읽을 수 있다.

공개 모델에서 확인한 속성:

- 문단 `p`: `paraPrIDRef`, `styleIDRef`, `pageBreak`, `columnBreak`
- 런 `run`: `charPrIDRef`
- 스타일 `style`: `paraPrIDRef`, `charPrIDRef`
- 글자 모양 `charPr`: `height`, `textColor`, `fontRef`, `bold`, `italic`, `underline`
- 문단 모양 `paraPr`: `align`, `margin`, `lineSpacing`
- 페이지 모양 `pagePr`: `width`, `height`, `landscape`, `margin`

### 3. 현재 MAX-Viewer와의 격차

현재 구현 상태:

- `content.hpf` 기준으로 섹션 순서를 읽음
- `header.xml`에서 메타데이터와 기본 스타일 테이블을 읽음
- `section.xml`에서 문단/표/이미지와 `secPr/pagePr` 기반 구역 레이아웃을 읽음
- 화면은 구역별 흰 종이 페이지와 기본 배율 제어를 제공

한컴뷰어와의 주요 차이:

1. 실제 쪽 나눔이 없어 구역을 페이지 캔버스로 근사하고 있다.
2. 머리말/꼬리말이 페이지 내부에 렌더링되지 않는다.
3. 번호 매기기/글머리표와 문단 번호 체계가 없다.
4. 그림, 도형, 수식, OLE 배치가 한컴뷰어와 다르다.
5. 표 셀이 문자열 수준이라 셀 내부 문단/개체 조판이 부족하다.
6. 글자 폭, 자간, 베이스라인 보정이 약해 페이지 경계에서 줄 수 차이가 남는다.

## 개선 전략

### 단계 1. 내부 문서 모델 확장

`max_viewer_core`에 아래 구조를 추가한다.

- `PageLayout`
  - `width`
  - `height`
  - `landscape`
  - `margin_left`
  - `margin_right`
  - `margin_top`
  - `margin_bottom`
  - `margin_header`
  - `margin_footer`
- `ParagraphStyle`
  - `align`
  - `indent`
  - `margin_left`
  - `margin_right`
  - `margin_prev`
  - `margin_next`
  - `line_spacing_type`
  - `line_spacing`
- `TextStyle`
  - `font_family`
  - `font_size`
  - `color`
  - `background_color`
  - `bold`
  - `italic`
  - `underline`

추가로 다음 구조를 넣는다.

- `TextStyle`
  - `width_ratio`
  - `letter_spacing`
  - `relative_size`
  - `baseline_offset`
  - `use_font_space`
  - `use_kerning`
- `TableCell`
  - `blocks`
  - `col_span`
  - `row_span`
  - `width`
  - `height`
  - `padding_left/right/top/bottom`
- `ImageBlock`
  - `kind`
  - `width/height`
  - `width_rel_to/height_rel_to`
  - `treat_as_char`
  - `text_wrap`
  - `horz/vert align`
  - `horz/vert offset`
  - `z_order`
  - `caption`

그리고 다음 연결을 만든다.

- `Section.page_layout`
- `Paragraph.style`
- `TextRun.style`

### 단계 2. HWPX 스타일/메트릭 해석

`max_viewer_hwpx`에서 다음 순서로 해석한다.

1. `header.xml`의 `refList`를 읽어 font/char/para/style 맵 구성
2. `styleIDRef`가 있으면 그 스타일이 참조하는 `paraPrIDRef`, `charPrIDRef`를 따라감
3. 문단의 직접 `paraPrIDRef`가 있으면 스타일보다 우선 적용
4. 런의 직접 `charPrIDRef`가 있으면 상위 스타일보다 우선 적용
5. `secPr/pagePr/margin`을 읽어 구역별 페이지 레이아웃 구성

이번 단계에서 실제 적용할 속성:

- paragraph: `horizontal align`, `indent`, `left/right`, `prev/next`, `line spacing`
- text: `height`, `textColor`, `fontRef`, `bold`, `italic`, `underline`, `ratio`, `spacing`, `relSz`, `offset`, `useKerning`
- section: `pagePr.width`, `pagePr.height`, `margin.left/right/top/bottom/header/footer`

### 단계 3. 표/개체 구조 해석

`max_viewer_hwpx`에서 다음을 추가 해석한다.

1. `tbl/tr/tc/subList`를 읽어 셀 내부 블록 목록 복원
2. `cellSpan`, `cellSz`, `cellMargin`으로 셀 병합/크기/안쪽 여백 복원
3. `pic`, `ole`, 기본 도형 계열의 `sz`, `pos`, `textWrap`, `zOrder`를 읽어 개체 배치 힌트 복원
4. `content.hpf`와 `BinData`를 연결해 실제 그림 자산을 data URI로 전달

### 단계 4. 페이지 기반 렌더러

프론트엔드를 다음과 같이 바꾼다.

1. 회색 작업 배경
2. 가운데 정렬된 흰 종이 페이지
3. 각 섹션을 하나의 페이지 카드처럼 표시
4. 페이지 안쪽 패딩을 HWPX 여백값으로 반영
5. 문단마다 스타일을 CSS로 변환
6. 텍스트 런을 `<span>` 단위로 렌더링해 글자 모양을 적용

CSS 변환 규칙:

- HWP 단위는 `1/7200 inch` 기준으로 보고, 화면 렌더링은 `px = hwp_unit / 75`로 변환
- 글자 높이는 `px = height / 100`으로 먼저 근사 적용
- 정렬은 `left/center/right/justify`로 매핑
- 줄 간격은 타입별로 별도 변환 함수를 둔다

### 단계 5. 비교/검증

비교용 fixture 문서를 만든다.

- 본문 기본 문서
- 정렬/들여쓰기 문서
- 글자 크기/굵기/기울임/색 문서
- 표가 섞인 문서
- 여백과 페이지 방향이 다른 구역 문서

검증 방식:

1. 한컴뷰어에서 동일 문서를 열어 기준 스크린샷 확보
2. MAX-Viewer 렌더링과 나란히 비교
3. 시각 차이를 체크리스트로 관리

## 이번 턴 구현 범위

이번 턴에서 실제로 구현할 항목은 다음이다.

1. `max_viewer_core` 문서 모델에 글자 메트릭, 표 셀 블록, 개체 배치 필드 추가
2. `max_viewer_hwpx`에서 `charPr` 메트릭과 `tbl/tc/subList`, `pic/ole/shape` 구조 파싱 추가
3. `BinData`를 실제 이미지로 렌더링할 수 있게 자산 data URI 전달 추가
4. 프론트엔드에서 표 셀 내부 문단/개체 렌더링과 floating 객체 절대 배치 추가
5. 페이지 수 계산과 placeholder 치환을 실제 렌더링 결과 기준으로 보정

이번 턴에서 미루는 항목:

- 각주/미주/주석
- 도형/OLE/수식의 회전/효과/클리핑 정밀 복원
- 한컴 내부 엔진과 1:1 동일한 페이지 재조판

## 문서 업데이트 원칙

이 문서는 구현이 진행될 때마다 함께 갱신한다.

- 완료된 항목은 `현재 상태` 문서에 반영
- 범위가 변경되면 이 계획 문서를 먼저 수정
- 구현과 문서가 어긋나지 않도록 `docs/README.md` 인덱스를 같이 갱신

## 참고 자료

- 한컴테크: HWPX 포맷 구조  
  <https://tech.hancom.com/hwpxformat/>
- 한컴 공개 OWPML 모델 저장소  
  <https://github.com/hancom-io/hwpx-owpml-model>
- 한컴 도움말: 화면 확대/축소  
  <https://help.hancom.com/hoffice_mac/ko_kr/hwp/view/zooming/zoom.htm>
- 한컴 도움말: 여러 쪽  
  <https://help.hancom.com/hoffice_mac/ko_kr/hwp/view/zooming/multiple_pages.htm>
- 한컴 도움말: 쪽 윤곽  
  <https://help.hancom.com/hoffice_mac/ko_kr/hwp/view/page_outline.htm>
- 한컴 도움말: 머리말/꼬리말  
  <https://help.hancom.com/hoffice_mac/ko_kr/hwp/format/header/header.htm>

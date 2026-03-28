# Floating 개체 본문 감싸기(Text Wrap) 구현 계획

작성일: 2026-03-28

## 목표

현재 MAX-Viewer는 floating 그림/도형의 위치를 절대 좌표 또는 CSS float로 배치하지만, 본문 텍스트가 개체를 감싸며 흐르는 text wrap을 제대로 구현하지 않는다. 한컴뷰어에서는 floating 개체 주변으로 본문이 자연스럽게 reflow되는데, MAX-Viewer에서는 본문과 개체가 겹치거나 개체 아래로 빈 공간이 생긴다.

이번 과제의 목표는 다음과 같다.

1. `textWrap` 속성에 따라 본문이 floating 개체를 좌우로 감싸며 흐르도록 구현
2. `TOP_AND_BOTTOM`, `SQUARE`, `TIGHT`, `BEHIND_TEXT`, `IN_FRONT_OF_TEXT` 모드를 구분 처리
3. 같은 줄 높이 안에 여러 개체가 있을 때 본문 영역을 올바르게 축소
4. 페이지 분할에서 floating 개체의 footprint를 정확히 반영

이번 과제의 비목표는 다음과 같다.

- 도형 윤곽을 따라 흐르는 pixel-perfect tight wrap
- 개체 회전/클리핑/효과 렌더링
- anchor 이동이나 편집 기능

## 현재 구현 상태 (2026-03-28 반영)

### Rust 코어 (`max_viewer_core/src/lib.rs`)

`ImageBlock` 구조체에 위치, wrap, distance 필드가 모두 존재한다.

- `treat_as_char: bool` - inline vs floating 구분
- `text_wrap: Option<String>` - 감싸기 모드
- `horz_align/vert_align` - 정렬
- `horz_offset/vert_offset` - 위치 오프셋
- `horz_rel_to/vert_rel_to` - 기준 좌표계
- `z_order: Option<i32>` - 레이어 순서
- `distance_left/right/top/bottom: Option<i32>` - 개체와 본문 사이 간격 (단계 6 완료)

### HWPX 파서 (`max_viewer_hwpx/src/lib.rs`)

- `parse_position_attributes`에서 `<pos>` 요소 파싱
- `parse_image`에서 `<outMargin>` 요소의 left/right/top/bottom 간격 파싱 (단계 6 완료)

### 프론트엔드 (`packages/viewer-ui/src/App.tsx`)

- `resolveTextWrapMode`: textWrap 값을 정규화하여 INLINE/TOP_AND_BOTTOM/SQUARE/TIGHT/BEHIND_TEXT/IN_FRONT_OF_TEXT 구분 (단계 2 완료)
- `objectPlacementStyle`: wrap 모드별 z-index, pointer-events 분기 (단계 2 완료)
  - BEHIND_TEXT: z-index -1, pointer-events none
  - IN_FRONT_OF_TEXT: 높은 z-index, pointer-events none
- `supportsSideTextWrap`: SQUARE/TIGHT만 true (단계 2 완료)
- `floatingObjectReserveStyle`: distance 반영, shape-outside: margin-box, TOP_AND_BOTTOM clear 처리 (단계 2 완료)
- `floatingObjectFootprintHeight/Width`: wrap 모드별 footprint 계산, BEHIND/IN_FRONT는 0 (단계 5 완료)
- `calculateExclusionZone`: floating 개체의 점유 영역을 페이지 본문 좌표로 계산 (단계 1 완료)
- `collectExclusionZones`: 페이지 내 모든 floating 개체의 exclusion zone 수집 (단계 1 완료)
- `calculateParagraphInsets`: exclusion zone과 겹치는 문단에 좌우 인셋 계산, 다중 개체 병합 및 최소 폭 보장 (단계 3, 4 완료)
- 페이지 본문 렌더링에서 블록별 누적 높이 추적 → exclusion 기반 인셋 적용 (단계 3 완료)
- `imageDistancePx`: outMargin 값을 px로 변환, 기본값 적용 (단계 6 완료)
- `pageContentWidth`: 페이지 본문 가용 폭 계산 (단계 1 기반)

### 남은 과제

- `horzRelTo`가 page/paragraph일 때 좌표 변환 보정 (현재 column 기준 통일)
- 실제 DOM 높이 측정 기반 인셋 보정 (현재 layoutHeightHint 추정치 사용)
- pixel-perfect tight wrap (현재 TIGHT는 SQUARE와 동일)

## 핵심 개념

### textWrap 모드 분류

HWPX/OWPML의 `textWrap` 값과 렌더링 동작을 다음과 같이 매핑한다.

| textWrap 값 | 동작 | 본문 reflow |
|---|---|---|
| `TOP_AND_BOTTOM` | 개체 위아래로만 본문 배치, 좌우에 본문 없음 | 개체 높이만큼 본문 영역을 수직으로 비움 |
| `SQUARE` | 개체의 bounding box 기준으로 좌우 감싸기 | 좌우 본문 영역 축소 |
| `TIGHT` | 개체 윤곽 기준 감싸기 (1차는 SQUARE와 동일) | 좌우 본문 영역 축소 |
| `BEHIND_TEXT` | 개체가 본문 뒤 (배경) | 본문 reflow 없음, z-index만 조정 |
| `IN_FRONT_OF_TEXT` | 개체가 본문 앞 (오버레이) | 본문 reflow 없음, z-index만 조정 |
| 없음 / 기본값 | `treatAsChar`면 inline, 아니면 SQUARE 근사 | 상황에 따라 |

### 본문 감싸기 레이아웃 모델

CSS의 `shape-outside` 또는 `float`만으로는 한컴의 감싸기를 정확히 재현하기 어렵다. 현실적인 접근은 다음과 같다.

1. **floating 개체의 점유 영역(exclusion zone)을 계산**
   - 개체의 위치(offset + align)와 크기(width, height)로 bounding box 결정
   - 좌우 여백(distanceFromText)이 있으면 bounding box 확장

2. **본문 문단에 CSS margin/padding으로 감싸기 근사**
   - 개체가 왼쪽에 있으면 본문 `margin-left` 추가
   - 개체가 오른쪽에 있으면 본문 `margin-right` 추가
   - 개체 높이 범위에 해당하는 문단/줄에만 적용

3. **페이지 좌표 기반 영역 할당**
   - 페이지 본문 영역에서 개체가 차지하는 좌표 범위를 계산
   - 그 범위와 겹치는 문단에 동적으로 좌우 인셋을 적용

## 구현 계획

### 단계 1. exclusion zone 계산 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

floating 개체의 점유 영역을 페이지 본문 좌표계로 계산하는 함수를 추가한다.

구현 항목:

1. `calculateExclusionZone(image: ImageBlock, pageLayout: PageLayout)` 함수 신규 작성
   - 입력: 개체 정보, 페이지 레이아웃
   - 출력: `{ left, top, right, bottom, wrapMode }` (페이지 본문 영역 내 좌표)
   - `horzRelTo` 기준 변환:
     - `page`: 페이지 왼쪽 가장자리 기준 → 본문 영역 기준으로 변환
     - `column` / `paragraph`: 본문 영역 기준
   - `horzAlign` 처리:
     - `LEFT`: left = margin_left + horzOffset
     - `CENTER`: left = (contentWidth - objWidth) / 2 + horzOffset
     - `RIGHT`: left = contentWidth - objWidth + horzOffset
   - `vertRelTo` + `vertAlign` 유사하게 처리

2. `collectExclusionZones(blocks: Block[], pageLayout: PageLayout)` 함수
   - 페이지 내 모든 floating 개체의 exclusion zone 목록 반환
   - `BEHIND_TEXT`, `IN_FRONT_OF_TEXT`는 exclusion 대상에서 제외

예상 변경량: 함수 2개 (각 40-50줄)

### 단계 2. textWrap 모드별 렌더링 분기 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

기존 `objectPlacementStyle`과 `floatingObjectReserveStyle`을 개선한다.

구현 항목:

1. `objectPlacementStyle` 개선
   - `BEHIND_TEXT`: `z-index: -1` 또는 낮은 값, `pointer-events: none`
   - `IN_FRONT_OF_TEXT`: `z-index` 높은 값
   - `TOP_AND_BOTTOM`: 기존 absolute 유지, 본문 clearance만 확보
   - `SQUARE` / `TIGHT`: float 기반 배치 + 본문 인셋

2. `supportsSideTextWrap` 로직 정교화
   - `SQUARE`, `TIGHT`: true (좌우 감싸기)
   - `TOP_AND_BOTTOM`: false (위아래만)
   - `BEHIND_TEXT`, `IN_FRONT_OF_TEXT`: false (본문에 영향 없음)

3. `wrapModeStyle(image: ImageBlock)` 함수 신규 작성
   - wrap 모드에 따른 CSS 속성 세트 반환
   - SQUARE/TIGHT → CSS `float` + `shape-outside: margin-box` + margin
   - TOP_AND_BOTTOM → `display: block` + `clear: both` + margin-top/bottom

예상 변경량: 기존 함수 수정 30줄, 신규 함수 1개 (30줄)

### 단계 3. 본문 인셋 적용 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

floating 개체와 같은 수직 범위에 있는 본문 문단에 좌우 인셋을 적용한다.

구현 항목:

1. `calculateParagraphInsets(paragraphTop: number, paragraphBottom: number, exclusions: ExclusionZone[])` 함수 신규 작성
   - 문단의 수직 범위와 겹치는 exclusion zone을 찾음
   - 왼쪽 exclusion이 있으면 `marginLeft` 반환
   - 오른쪽 exclusion이 있으면 `marginRight` 반환
   - 여러 exclusion이 겹치면 가장 넓은 인셋 적용

2. 문단 렌더링 시 인셋 적용
   - `renderParagraph`에서 해당 문단의 누적 높이를 기준으로 인셋 계산
   - 계산된 인셋을 문단의 inline style에 추가

3. 높이 추적
   - 페이지 내 블록을 렌더링하면서 누적 높이를 추적
   - 각 문단의 수직 위치를 exclusion zone과 대조

예상 변경량: 함수 1개 (30줄), 렌더링 수정 20줄, 높이 추적 15줄

### 단계 4. 다중 개체 처리 (프론트엔드)

같은 수직 범위에 여러 floating 개체가 있을 때의 처리.

구현 항목:

1. exclusion zone 병합
   - 같은 쪽(왼쪽/오른쪽)에 있는 개체들의 exclusion을 병합
   - 양쪽에 개체가 있으면 본문 영역이 양쪽에서 축소

2. 본문 영역 최소 폭 보장
   - 본문 영역이 너무 좁아지면(예: 전체 폭의 20% 미만) 해당 구간은 TOP_AND_BOTTOM 방식으로 fallback
   - 개체 아래로 본문을 밀어냄

3. z-order 기반 레이어링
   - 개체 간 겹침이 있을 때 `z_order` 순서대로 렌더링

예상 변경량: 병합 로직 30줄, fallback 20줄

### 단계 5. 페이지 분할 연동 (프론트엔드)

floating 개체의 footprint를 페이지 분할 계산에 정확히 반영한다.

구현 항목:

1. `floatingObjectFootprintHeight` 개선
   - `TOP_AND_BOTTOM`: 개체 높이 + 위아래 간격이 본문 높이에 추가
   - `SQUARE` / `TIGHT`: 개체가 본문 옆에 있으므로 max(개체 높이, 감싸진 본문 높이)
   - `BEHIND_TEXT` / `IN_FRONT_OF_TEXT`: 본문 높이에 영향 없음

2. 페이지 경계에서 개체 위치 결정
   - 개체가 페이지 경계를 넘으면: 개체를 다음 페이지로 이동
   - 개체와 anchor 문단의 관계 유지: anchor 문단이 다음 페이지로 넘어가면 개체도 함께 이동

예상 변경량: footprint 함수 수정 30줄, 경계 처리 20줄

### 단계 6. Rust 코어 모델 보강 (선택)

향후 서버사이드 레이아웃을 위한 필드 추가.

**파일**: `crates/max_viewer_core/src/lib.rs`

```rust
pub struct ImageBlock {
    // ... 기존 필드
    pub distance_left: Option<i32>,    // 개체와 본문 사이 간격
    pub distance_right: Option<i32>,
    pub distance_top: Option<i32>,
    pub distance_bottom: Option<i32>,
}
```

**파일**: `crates/max_viewer_hwpx/src/lib.rs`

- `parse_image`에서 `<outMargin>` 또는 `<offset>` 요소의 상하좌우 간격 속성 파싱

예상 변경량: 코어 필드 4개 추가, 파서 15줄 수정

## 구현 순서

```
단계 1 (exclusion zone) ← 기반 인프라
  ↓
단계 2 (모드별 렌더링) ← 단계 1 완료 후
  ↓
단계 3 (본문 인셋) ← 단계 1, 2 완료 후
  ↓
단계 4 (다중 개체) ← 단계 3 완료 후
  ↓
단계 5 (페이지 분할) ← 단계 1 완료 후, 단계 3과 병행 가능
  ↓
단계 6 (코어 모델) ← 독립, 언제든 가능
```

## 검증 방법

### 테스트 문서

다음 유형의 HWPX fixture를 준비한다.

1. **좌측 floating + 우측 본문 감싸기**: 그림이 왼쪽, 본문이 오른쪽으로 흐르는 문서
2. **우측 floating + 좌측 본문 감싸기**: 반대 방향
3. **중앙 floating + TOP_AND_BOTTOM**: 개체 위아래로만 본문
4. **BEHIND_TEXT**: 배경 이미지 위에 본문
5. **다중 개체**: 같은 단락 근처에 2-3개 floating 개체
6. **페이지 경계 floating**: 개체가 페이지 아래쪽에 있어 경계를 넘는 경우

### 검증 기준

- 본문 텍스트가 개체와 겹치지 않음
- 개체 주변 간격이 자연스러움
- TOP_AND_BOTTOM 모드에서 개체 좌우에 본문이 없음
- BEHIND_TEXT에서 본문이 개체 위에 정상 표시
- 한컴뷰어와 나란히 비교했을 때 감싸기 방향과 간격이 유사

## 위험 요소

1. **CSS float 한계**: CSS float만으로는 한컴의 감싸기를 완벽히 재현하기 어려움 → 인셋 margin 방식으로 보완
2. **높이 추적 정확도**: 문단의 실제 렌더링 높이와 계산된 높이의 차이 → 기존 line measurement 로직 재활용
3. **성능**: 많은 floating 개체가 있는 문서에서 exclusion zone 계산 비용 → 페이지 단위로 캐시
4. **기준 좌표계 다양성**: `horzRelTo`가 page/column/paragraph 등 다양 → 1차는 column 기준으로 통일, 이후 보정

## 참고

- 현재 floating 처리: `packages/viewer-ui/src/App.tsx` `objectPlacementStyle`, `floatingObjectReserveStyle`
- HWPX 위치 속성: `<pos treatAsChar="" textWrap="" vertRelTo="" horzRelTo="" ...>`
- 한컴 OWPML 개체 배치: `<sz>`, `<pos>`, `<outMargin>` 요소
- CSS shape-outside 참고: 향후 TIGHT wrap 정밀도 향상 시 활용 가능

# 복잡한 rowSpan 표 페이지 분할 개선 계획

작성일: 2026-03-28

## 목표

현재 MAX-Viewer는 표를 페이지 경계에서 분할할 때 단순 행 단위 또는 1x1 박스형 테이블 내부 블록 단위로만 끊는다. `rowSpan`이 걸린 셀이 포함된 복잡한 표에서는 분할 지점을 잘못 잡거나, 병합 셀이 잘리면서 렌더링이 깨진다.

이번 과제의 목표는 다음과 같다.

1. `rowSpan > 1`인 셀이 있는 표를 페이지 경계에서 안전하게 분할한다.
2. 분할된 표의 각 조각이 독립적으로 올바른 HTML `<table>`로 렌더링된다.
3. `repeatHeader: true`인 표는 분할된 페이지마다 헤더 행을 반복한다.
4. 한컴뷰어와의 시각 차이를 줄인다.

이번 과제의 비목표는 다음과 같다.

- 표 셀 내부를 줄 단위로 다시 쪼개는 셀 내부 분할 (셀 단위 분할까지만)
- 한컴 조판 엔진과 1:1 동일한 행 높이 계산
- 표 자동 맞춤, 캡션 줄바꿈 등 고급 표 기능

## 현재 구현 상태 (2026-03-28 반영)

### Rust 코어 (`max_viewer_core/src/lib.rs`)

- `TableBlock`: `rows`, `repeat_header`, `header_row_count: Option<u32>` 필드 존재
- `TableCell`: `col_span: Option<u32>`, `row_span: Option<u32>` 필드 존재
- `TableRow`: `cells: Vec<TableCell>`

### HWPX 파서 (`max_viewer_hwpx/src/lib.rs`)

- `parse_cell`에서 `<cellSpan>` 요소의 `colSpan`, `rowSpan` 속성을 읽음
- `parse_table`에서 `repeatHeader` 속성을 읽고, `is_header` 셀 기반으로 `header_row_count` 계산
- 셀 내부 `<subList>` 블록을 재귀적으로 파싱

### 프론트엔드 (`packages/viewer-ui/src/App.tsx`)

- `computeSafeRowBreakStarts`: rowSpan을 추적하여 안전한 분할 행 판별 (단계 1 완료)
- `buildAtomicRowGroups` + `splitTableByRowsItem`: 안전 분할 행 기반 표 조각 생성 (단계 2 완료)
- `resolveHeaderRowCount` + `cloneTableRowsFragment`: Rust 모델의 `headerRowCount`를 우선 사용, 분할 시 헤더 행 반복 (단계 3 완료)
- `buildRowSpanGrid` + `findBestForcedSplitRow` + `forceSplitTableByRows`: 안전한 분할 행이 없을 때 rowSpan이 가장 적게 걸친 행에서 강제 분할 (단계 4 완료)
- `cloneTableRowsFragmentWithSpanFix`: 강제 분할 시 잘린 rowSpan 셀의 span 보정 및 continuation 셀 삽입 (단계 4 완료)
- Rust 코어 `header_row_count` 필드 추가 및 HWPX 파서 연동 (단계 5 완료)

### 남은 과제

- 실제 복잡한 rowSpan 문서로 시각 검증
- 강제 분할 시 셀 내부 블록이 잘리는 극단적 케이스 보정

## 핵심 개념: rowSpan 인지 분할

### 안전한 분할 행 판별

표를 행 `r`에서 분할하려면, 그 행 위쪽에서 시작해 `r` 이후까지 걸치는 `rowSpan` 셀이 없어야 한다.

```
안전한_분할_행(r) = ∀ 셀 c, c.시작행 + c.rowSpan - 1 < r
```

즉, 행 `r`이 어떤 `rowSpan` 셀의 중간에 있으면 그 행에서는 분할할 수 없다.

### rowSpan 그리드 맵

표 전체를 `rows × cols` 2차원 그리드로 매핑한다. 각 칸에는 해당 셀의 원점(시작 행, 시작 열)과 span 정보가 들어간다.

```
GridMap[row][col] = {
  ownerRow: number,    // 셀이 시작하는 행
  ownerCol: number,    // 셀이 시작하는 열
  rowSpan: number,
  colSpan: number,
}
```

이 맵을 기준으로:
- 분할 가능 행을 O(rows) 스캔으로 판별
- 분할 후 아래쪽 조각에서 잘린 `rowSpan` 셀의 잔여 높이를 계산

## 구현 계획

### 단계 1. rowSpan 그리드 맵 생성 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

현재 pagination 로직에서 표 블록을 만나면 행 단위로 분할 가능 여부를 판단하는데, 여기에 그리드 맵 기반 판별을 추가한다.

구현 항목:

1. `buildRowSpanGrid(table: TableBlock)` 함수 신규 작성
   - 입력: `TableBlock`의 `rows` 배열
   - 출력: `GridMap[row][col]` 2차원 배열
   - 각 행을 순회하면서 `colSpan`과 `rowSpan`을 그리드에 기록
   - 이미 점유된 칸(위쪽 행의 `rowSpan`이 차지)은 건너뜀

2. `findSafeSplitRows(grid: GridMap, rowCount: number)` 함수 신규 작성
   - 입력: 그리드 맵, 행 수
   - 출력: 안전하게 분할 가능한 행 인덱스 목록
   - 행 `r`이 안전하려면: 그리드의 모든 열에서 `ownerRow + rowSpan - 1 < r`

3. 기존 표 분할 로직에서 `findSafeSplitRows` 결과를 참조하도록 수정

예상 변경량: 함수 2개 신규 (각 30-40줄), 기존 분할 로직 수정 20줄 내외

### 단계 2. 표 조각 생성 로직 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

안전한 분할 행이 결정되면, 표를 여러 조각으로 나눈다.

구현 항목:

1. `splitTableAtRows(table: TableBlock, splitRows: number[])` 함수 신규 작성
   - 입력: 원본 표, 분할 행 인덱스 목록
   - 출력: `TableBlock[]` 조각 배열
   - 각 조각은 독립적인 `<table>`로 렌더링 가능해야 함

2. 분할 시 `rowSpan` 보정
   - 분할 경계를 넘는 `rowSpan` 셀이 없으므로 (안전한 행에서만 분할) 추가 보정 불필요
   - 다만, 분할 경계 바로 위 행에서 끝나는 `rowSpan` 셀의 `rowSpan` 값은 원래 값 유지

3. 안전한 분할 행이 없는 경우의 fallback
   - 표 전체를 한 페이지에 넣을 수 없고 안전한 분할 행도 없으면
   - 차선: `rowSpan`이 가장 적게 걸친 행에서 강제 분할
   - 강제 분할 시 잘린 `rowSpan` 셀은 아래쪽 조각에서 잔여 `rowSpan`으로 계속 표시

예상 변경량: 함수 1개 신규 (50-70줄), fallback 로직 30줄 내외

### 단계 3. 헤더 행 반복 (프론트엔드)

**파일**: `packages/viewer-ui/src/App.tsx`

`repeatHeader: true`인 표가 분할될 때, 첫 번째 조각 이후의 모든 조각에 헤더 행을 삽입한다.

구현 항목:

1. 헤더 행 식별
   - `TableCell.is_header == true`인 셀이 포함된 연속 행들을 헤더 영역으로 판별
   - 또는 `repeatHeader`가 설정되면 첫 번째 행을 헤더로 간주

2. 분할된 조각에 헤더 삽입
   - 두 번째 조각부터 헤더 행을 앞에 복제 삽입
   - 삽입된 헤더 행의 높이도 페이지 남은 공간 계산에 포함

3. 헤더 행은 분할 대상에서 제외
   - 헤더 영역 자체는 분할하지 않음

예상 변경량: 헤더 식별 15줄, 삽입 로직 20줄

### 단계 4. 강제 분할 시 rowSpan 보정 (프론트엔드)

안전한 분할 행이 없어 강제 분할해야 할 때, 잘린 `rowSpan` 셀을 처리한다.

구현 항목:

1. 위쪽 조각: 잘린 셀의 `rowSpan`을 분할 행까지의 남은 행 수로 축소
2. 아래쪽 조각: 잘린 셀을 첫 행에 계속 셀로 삽입, `rowSpan`은 잔여 행 수
3. 계속 셀은 원본 셀의 `colSpan`, 스타일, 내용을 복제하되 `(continued)` 표시 없이 자연스럽게 이어지도록 함

예상 변경량: 보정 함수 60-80줄

### 단계 5. Rust 코어 모델 보강 (선택)

현재 `TableBlock`에는 분할 관련 메타데이터가 없다. 프론트엔드에서 분할하는 현재 방식을 유지하되, 향후 서버사이드 분할을 위해 다음 필드를 예약할 수 있다.

**파일**: `crates/max_viewer_core/src/lib.rs`

```rust
pub struct TableBlock {
    // ... 기존 필드
    pub header_row_count: Option<u32>,  // 헤더 행 수 (repeatHeader와 연동)
}
```

**파일**: `crates/max_viewer_hwpx/src/lib.rs`

- `parse_table`에서 `<headerRows>` 또는 `is_header` 셀을 기준으로 `header_row_count` 계산

예상 변경량: 코어 모델 필드 1개 추가, 파서 10줄 수정

## 구현 순서

```
단계 1 (그리드 맵)
  ↓
단계 2 (조각 생성)
  ↓
단계 3 (헤더 반복) ← 단계 2 완료 후
  ↓
단계 4 (강제 분할 보정) ← 단계 1, 2 완료 후
  ↓
단계 5 (코어 모델) ← 독립, 언제든 가능
```

## 검증 방법

### 테스트 문서

다음 유형의 HWPX fixture를 준비한다.

1. **단순 rowSpan**: 2행 병합 셀이 있는 3x3 표
2. **복합 rowSpan**: 여러 열에 걸친 다단계 행 병합
3. **페이지 넘김 rowSpan**: 표가 페이지 경계를 넘고 rowSpan 셀이 경계에 걸치는 경우
4. **헤더 반복 표**: `repeatHeader: true`이고 2페이지 이상 걸치는 표
5. **안전한 분할 행이 없는 표**: 모든 행에 rowSpan이 걸린 극단적 케이스

### 검증 기준

- 분할된 각 표 조각이 독립적으로 올바른 HTML table 구조를 가짐
- `rowSpan` 셀이 페이지 경계에서 시각적으로 자연스럽게 이어짐
- 헤더 행이 올바르게 반복됨
- 한컴뷰어와 나란히 비교했을 때 행 구성이 동일하거나 근사함

## 위험 요소

1. **성능**: 대형 표(100행 이상)에서 그리드 맵 생성 비용 → O(rows × cols)이므로 문제 없을 것으로 예상
2. **복합 병합**: `colSpan`과 `rowSpan`이 동시에 걸린 셀의 그리드 매핑 정확도 → 단위 테스트로 커버
3. **강제 분할 품질**: 안전한 분할 행이 없을 때 차선 분할의 시각 품질 → fixture 기반 비교 테스트

## 참고

- 현재 표 분할 로직: `packages/viewer-ui/src/App.tsx` pagination 관련 코드
- 한컴 HWPX 표 구조: `<tbl>` → `<tr>` → `<tc>` → `<cellSpan colSpan="" rowSpan="">`
- 한컴 헤더 반복: `<tbl repeatHeader="true">` + `<tc>` 의 `header="true"` 속성

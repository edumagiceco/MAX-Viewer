# HWP DocInfo/Control 복원도 향상 계획

작성일: 2026-03-28

## 목표

현재 MAX-Viewer의 HWP 파서는 `FileHeader` 읽기, `BodyText`의 PARA_HEADER/PARA_TEXT 레코드 추출, `PrvText` fallback preview만 지원한다. `DocInfo` 스트림을 읽지 않으므로 글꼴, 문단 모양, 글자 모양, 스타일, BinData 참조 정보가 없고, `BodyText`의 control 레코드를 읽지 않으므로 표, 그림, 도형, 페이지 설정이 복원되지 않는다.

이번 과제의 목표는 다음과 같다.

1. `DocInfo` 스트림의 레코드를 파싱하여 글꼴/글자 모양/문단 모양/스타일/BinData 매핑 테이블을 구축
2. `BodyText`의 control 레코드를 파싱하여 표, 그림, 페이지 설정을 복원
3. 파싱 결과를 기존 `max_viewer_core` 문서 모델로 변환하여 HWPX와 동일한 렌더링 경로를 공유
4. HWP 문서를 열었을 때 텍스트뿐 아니라 기본 서식과 표/이미지가 표시되도록 함

이번 과제의 비목표는 다음과 같다.

- 암호화/DRM 문서 해독
- 배포용 문서의 특수 보호 해제
- OLE/수식/차트의 실제 렌더링
- 복합 도형의 정밀 복원
- 한컴 편집기와 1:1 동일한 레이아웃

## 현재 구현 상태 (2026-03-28 반영)

### HWP 파서 (`max_viewer_hwp/src/lib.rs`)

현재 ~1300줄 규모. 8단계 모두 구현 완료.

**단계 1 (DocInfo 기반)**: `iter_records` 범용 레코드 반복기, `DocInfoStore` 구조체, `read_doc_info` 함수 — 완료
**단계 2 (글꼴/글자/문단)**: `FaceName`, `CharShape`, `ParaShape`, `HwpStyle` 파서 — 완료
**단계 3 (BinData)**: `BinDataRef` 파서, `load_bin_data_assets` (base64 data URI 생성), `guess_media_type` — 완료
**단계 4 (문단 서식)**: `PARA_HEADER` 확장 (paraShapeId, styleId), `PARA_CHAR_SHAPE` 파싱, `char_shape_to_text_style`, `para_shape_to_paragraph_style` 변환 — 완료
**단계 5 (표)**: `CTRL_HEADER("tbl ")`, `HWPTAG_TABLE`, `HWPTAG_LIST_HEADER` 파싱, `assemble_table` (col_addr/row_addr 기반 2D 배치) — 완료
**단계 6 (그림/도형)**: `CTRL_HEADER("gso ")`, `HWPTAG_SHAPE_COMPONENT` 파싱, BinData ID → `ImageBlock` 변환, 미지원 도형 → `UnsupportedBlock` — 완료
**단계 7 (페이지 설정)**: `HWPTAG_PAGE_DEF` 파싱 → `PageLayout`, `CTRL_HEADER("secd")` → Section 분리 — 완료
**단계 8 (통합)**: `parse_bytes` 확장 (DocInfo → BodyText → assets 연동), fallback 유지, `scaffold_support` 업데이트 — 완료

### 남은 과제

- Border/fill 스타일 적용
- 각주/미주 복원
- 고급 도형 복원
- 암호화/DRM 문서 처리

### 레코드 구조

HWP 5.0 공식 문서 기준 레코드 헤더:

```
[TagID: 10비트] [Level: 10비트] [Size: 12비트]
Size == 0xFFF이면 다음 4바이트가 실제 크기
```

TagID는 `HWPTAG_BEGIN(0x010) + offset`으로 정의된다.

## HWP 5.0 레코드 태그 참조

한컴 공식 PDF 기준으로 이번 과제에 필요한 주요 TagID는 다음과 같다.

### DocInfo 레코드

| 태그 이름 | offset | TagID | 설명 |
|---|---|---|---|
| HWPTAG_DOCUMENT_PROPERTIES | 0 | 0x010 | 문서 속성 (구역/시작번호) |
| HWPTAG_ID_MAPPINGS | 1 | 0x011 | ID 매핑 개수 테이블 |
| HWPTAG_BIN_DATA | 2 | 0x012 | BinData 참조 정보 |
| HWPTAG_FACE_NAME | 3 | 0x013 | 글꼴 이름 |
| HWPTAG_BORDER_FILL | 4 | 0x014 | 테두리/채움 |
| HWPTAG_CHAR_SHAPE | 5 | 0x015 | 글자 모양 |
| HWPTAG_TAB_DEF | 6 | 0x016 | 탭 정의 |
| HWPTAG_NUMBERING | 7 | 0x017 | 번호 매기기 |
| HWPTAG_BULLET | 8 | 0x018 | 글머리표 |
| HWPTAG_PARA_SHAPE | 9 | 0x019 | 문단 모양 |
| HWPTAG_STYLE | 10 | 0x01A | 스타일 |
| HWPTAG_DOC_DATA | 11 | 0x01B | 문서 데이터 |

### BodyText 레코드

| 태그 이름 | offset | TagID | 설명 |
|---|---|---|---|
| HWPTAG_PARA_HEADER | 50 | 0x042 | 문단 헤더 |
| HWPTAG_PARA_TEXT | 51 | 0x043 | 문단 텍스트 |
| HWPTAG_PARA_CHAR_SHAPE | 52 | 0x044 | 문단 내 글자 모양 참조 |
| HWPTAG_PARA_LINE_SEG | 53 | 0x045 | 줄 배치 정보 |
| HWPTAG_CTRL_HEADER | 54 | 0x046 | 컨트롤 헤더 |
| HWPTAG_LIST_HEADER | 55 | 0x047 | 리스트(표 셀 등) 헤더 |
| HWPTAG_PAGE_DEF | 56 | 0x048 | 쪽 정의 |
| HWPTAG_FOOTNOTE_SHAPE | 57 | 0x049 | 각주/미주 |
| HWPTAG_PAGE_BORDER_FILL | 58 | 0x04A | 쪽 테두리/배경 |
| HWPTAG_TABLE | 59 | 0x04B | 표 정보 |
| HWPTAG_SHAPE_COMPONENT | 60 | 0x04C | 도형 정보 |

## 구현 계획

### 단계 1. DocInfo 레코드 디코딩 기반

**파일**: `crates/max_viewer_hwp/src/lib.rs` (확장)

DocInfo 스트림을 레코드 단위로 읽고, 태그별로 분류하는 기반을 만든다.

구현 항목:

1. `DocInfoStore` 구조체 신규 정의
   ```rust
   struct DocInfoStore {
       face_names: Vec<FaceName>,
       char_shapes: Vec<CharShape>,
       para_shapes: Vec<ParaShape>,
       styles: Vec<HwpStyle>,
       border_fills: Vec<BorderFill>,
       bin_data_refs: Vec<BinDataRef>,
       numberings: Vec<Numbering>,
   }
   ```

2. `read_doc_info(compound_file) -> Result<DocInfoStore>` 함수
   - `/DocInfo` 스트림을 읽음
   - 압축 여부에 따라 해제
   - 레코드 반복 순회하며 태그별 분기

3. 기존 `parse_body_text_records` 패턴을 재활용한 범용 레코드 반복기
   ```rust
   fn iter_records(bytes: &[u8]) -> impl Iterator<Item = Record>
   ```
   - `Record { tag_id, level, data: &[u8] }` 구조

예상 변경량: 구조체 정의 50줄, read_doc_info 40줄, iter_records 30줄

### 단계 2. 글꼴/글자 모양/문단 모양 파싱

**파일**: `crates/max_viewer_hwp/src/lib.rs`

DocInfo의 핵심 매핑 테이블을 파싱한다.

구현 항목:

1. **HWPTAG_FACE_NAME (0x013)** 파싱
   ```rust
   struct FaceName {
       name: String,
       font_type: u8,  // 한글, 영문, 한자, 일본어, 기타, 기호, 사용자
   }
   ```
   - 레코드 데이터: 속성(1바이트) + 이름 길이(2바이트) + UTF-16LE 이름
   - 대체 글꼴, 판속 정보 등 추가 필드는 건너뜀

2. **HWPTAG_CHAR_SHAPE (0x015)** 파싱
   ```rust
   struct CharShape {
       face_name_ids: [u16; 7],  // 언어별 글꼴 참조
       ratios: [u8; 7],          // 장평
       spacings: [i8; 7],        // 자간
       rel_sizes: [u8; 7],       // 상대크기
       offsets: [i8; 7],         // 위치
       height: u32,              // 글자 크기 (1/7200 inch)
       attributes: u32,          // 속성 (굵게, 기울임, 밑줄 등)
       text_color: u32,          // 글자색 (COLORREF)
       shade_color: u32,
       // ... 추가 필드는 건너뜀
   }
   ```
   - 속성 비트맵: bit 0 = italic, bit 1 = bold, bit 2-3 = underline type
   - 색상: `COLORREF` 형식 (0x00BBGGRR)

3. **HWPTAG_PARA_SHAPE (0x019)** 파싱
   ```rust
   struct ParaShape {
       attributes: u32,     // 정렬(0-2비트), 세로정렬 등
       margin_left: i32,
       margin_right: i32,
       indent: i32,
       margin_prev: i32,
       margin_next: i32,
       line_spacing_type: u32,
       line_spacing: i32,
       // ... 추가 필드는 건너뜀
   }
   ```
   - 정렬: 0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔

4. **HWPTAG_STYLE (0x01A)** 파싱
   ```rust
   struct HwpStyle {
       name: String,
       para_shape_id: u16,
       char_shape_id: u16,
   }
   ```

각 파서 함수에 내결함성 원칙 적용:
- 알려진 필드만 읽고 남은 바이트는 건너뜀
- 레코드가 예상보다 짧으면 기본값으로 채움
- 파싱 실패 시 해당 항목만 건너뛰고 계속 진행

예상 변경량: 구조체 + 파서 함수 약 200줄

### 단계 3. BinData 참조 및 자산 로드

**파일**: `crates/max_viewer_hwp/src/lib.rs`

BinData 매핑 테이블을 구축하고, 실제 바이너리 자산을 로드한다.

구현 항목:

1. **HWPTAG_BIN_DATA (0x012)** 파싱
   ```rust
   struct BinDataRef {
       storage_type: u16,  // LINK, EMBEDDING, STORAGE
       abs_path: Option<String>,
       rel_path: Option<String>,
       bin_data_id: u16,   // BinData 스트림 내 ID
       extension: String,
   }
   ```

2. `BinData` 스트림 로드
   - `/BinData/BIN0001.xxx` 등의 스트림을 읽음
   - `DocInfoStore.bin_data_refs`의 인덱스와 매칭
   - base64 인코딩하여 `AssetRef.data_uri`로 변환

3. 미디어 타입 추정
   - 확장자 기반: png, jpg, gif, bmp, wmf, emf
   - 매직바이트 기반 fallback

예상 변경량: BinData 파서 30줄, 자산 로드 50줄

### 단계 4. BodyText 문단 서식 복원

**파일**: `crates/max_viewer_hwp/src/lib.rs`

현재 PARA_HEADER와 PARA_TEXT만 처리하는 `parse_body_text_records`를 확장한다.

구현 항목:

1. **HWPTAG_PARA_HEADER 확장**
   - 기존: 글자 수(4바이트)만 읽음
   - 추가: `paraShapeId`(offset 8, 2바이트), `styleId`(offset 10, 1바이트) 읽기

2. **HWPTAG_PARA_CHAR_SHAPE (0x044)** 파싱
   - 문단 내 글자 모양 변경 위치 목록
   - `[(position: u32, charShapeId: u32), ...]` 배열
   - 이 정보로 문단 텍스트를 run으로 분할

3. 문단에 서식 적용
   - `paraShapeId` → `DocInfoStore.para_shapes` → `ParagraphStyle`
   - `styleId` → `DocInfoStore.styles` → 스타일의 para/char shape 참조
   - `charShapeId` → `DocInfoStore.char_shapes` → `TextStyle`
   - face_name_id → `DocInfoStore.face_names` → font_family

4. `CharShape` → `TextStyle` 변환 함수
   ```rust
   fn char_shape_to_text_style(cs: &CharShape, face_names: &[FaceName]) -> TextStyle {
       TextStyle {
           font_family: face_names.get(cs.face_name_ids[0]).map(|f| f.name.clone()),
           font_size: Some((cs.height as i32) / 100),  // 1/100 pt 단위
           bold: cs.attributes & (1 << 1) != 0,
           italic: cs.attributes & (1 << 0) != 0,
           text_color: Some(colorref_to_hex(cs.text_color)),
           // ...
       }
   }
   ```

5. `ParaShape` → `ParagraphStyle` 변환 함수

예상 변경량: PARA_HEADER 확장 20줄, PARA_CHAR_SHAPE 파싱 30줄, 변환 함수 60줄

### 단계 5. Control 레코드 파싱 (표)

**파일**: `crates/max_viewer_hwp/src/lib.rs`

BodyText 내 control 레코드를 파싱하여 표를 복원한다.

구현 항목:

1. **제어 문자 감지**
   - PARA_TEXT에서 특수 제어 문자 감지:
     - `0x0002`: 구역/열 정의
     - `0x000b`: 컨트롤 확장 (표, 그림, 도형 등)
     - `0x000d`: 문단 끝
   - 제어 문자 위치를 기록하여 후속 CTRL_HEADER와 매칭

2. **HWPTAG_CTRL_HEADER (0x046)** 파싱
   - 처음 4바이트: 컨트롤 타입 코드 (ASCII 4문자)
   - 주요 타입:
     - `"tbl "` (0x62 6C 74 20): 표
     - `"gso "`: 그리기 개체
     - `"secd"`: 구역 정의
     - `"cold"`: 단 정의

3. **HWPTAG_TABLE (0x04B)** 파싱
   ```rust
   struct HwpTableDef {
       attributes: u32,
       row_count: u16,
       col_count: u16,
       cell_spacing: u16,
       margin: [u16; 4],  // left, right, top, bottom
       row_sizes: Vec<u16>,
       border_fill_id: u16,
   }
   ```

4. **HWPTAG_LIST_HEADER (0x047)** 파싱 (표 셀)
   ```rust
   struct HwpListHeader {
       para_count: u16,
       attributes: u32,
       col_addr: u16,
       row_addr: u16,
       col_span: u16,
       row_span: u16,
       width: u32,
       height: u32,
       margin: [u16; 4],
   }
   ```

5. 레코드 계층 구조(level) 활용
   - `level`로 부모-자식 관계 복원
   - 표 → 셀(LIST_HEADER) → 문단(PARA_HEADER/TEXT) 계층

6. `TableBlock`으로 변환
   - `HwpTableDef` + `HwpListHeader` 목록 → `TableBlock { rows, ... }`
   - `col_addr`, `row_addr`로 셀 위치 매핑
   - `col_span`, `row_span` 그대로 전달

예상 변경량: 제어 문자 감지 30줄, CTRL_HEADER 40줄, TABLE/LIST_HEADER 80줄, 변환 60줄

### 단계 6. Control 레코드 파싱 (그림/도형)

**파일**: `crates/max_viewer_hwp/src/lib.rs`

그리기 개체(gso) 컨트롤에서 그림을 복원한다.

구현 항목:

1. **gso 컨트롤 헤더** 파싱
   - 공통 개체 속성: 위치, 크기, 회전, z-order
   - `treatAsChar` 여부
   - `textWrap` 모드

2. **HWPTAG_SHAPE_COMPONENT (0x04C)** 파싱
   - 도형 타입 판별: 그림, 사각형, 타원, 선 등
   - 그림인 경우: `BinDataRef` ID 추출

3. `ImageBlock`으로 변환
   - BinData ID → 자산 참조
   - 위치/크기 → `ImageBlock` 필드
   - `treatAsChar`, `textWrap` 매핑

4. 지원하지 않는 도형 → `UnsupportedBlock`
   - 복합 도형, 수식, OLE 등은 placeholder로 처리
   - 크기 정보가 있으면 bounding box만 표시

예상 변경량: gso 헤더 40줄, SHAPE_COMPONENT 50줄, 변환 30줄

### 단계 7. 페이지 설정 복원

**파일**: `crates/max_viewer_hwp/src/lib.rs`

구역 정의(secd) 컨트롤에서 페이지 레이아웃을 복원한다.

구현 항목:

1. **HWPTAG_PAGE_DEF (0x048)** 파싱
   ```rust
   struct HwpPageDef {
       width: u32,       // 용지 너비
       height: u32,      // 용지 높이
       margin_left: u32,
       margin_right: u32,
       margin_top: u32,
       margin_bottom: u32,
       margin_header: u32,
       margin_footer: u32,
       margin_gutter: u32,
       landscape: bool,  // attributes 비트
   }
   ```

2. `PageLayout`으로 변환
   - 단위 변환: HWP 단위(1/7200 inch) → 코어 모델

3. `Section`에 `page_layout` 설정
   - `secd` 컨트롤이 나오면 새 Section 시작
   - 해당 Section의 `page_layout`에 반영

예상 변경량: PAGE_DEF 파서 30줄, 변환 15줄

### 단계 8. 통합 및 문서 모델 변환

**파일**: `crates/max_viewer_hwp/src/lib.rs`

각 단계에서 파싱한 결과를 기존 `parse_bytes` 흐름에 통합한다.

구현 항목:

1. `parse_bytes` 확장
   - DocInfo 읽기 → DocInfoStore 구축
   - BodyText 읽기 시 DocInfoStore 참조
   - 표/그림/페이지 설정을 포함한 완전한 Block 목록 생성

2. assets 생성
   - BinData 스트림에서 실제 바이너리를 읽어 `AssetRef` 생성
   - `Document.assets`에 추가

3. fallback 유지
   - DocInfo 파싱 실패 시 기존 텍스트 전용 경로로 fallback
   - 개별 레코드 파싱 실패 시 해당 항목만 건너뛰고 계속 진행

4. `scaffold_support` 업데이트
   - `implemented`에 새 항목 추가
   - `planned`에서 완료 항목 제거

예상 변경량: parse_bytes 수정 40줄, assets 30줄, fallback 20줄

## 구현 순서

```
단계 1 (DocInfo 기반)
  ↓
단계 2 (글꼴/글자/문단) ← 단계 1 완료 후
  ↓
단계 3 (BinData) ← 단계 1 완료 후, 단계 2와 병행 가능
  ↓
단계 4 (문단 서식) ← 단계 2 완료 후
  ↓
단계 5 (표) ← 단계 1 완료 후
  ↓
단계 6 (그림/도형) ← 단계 3, 5 완료 후
  ↓
단계 7 (페이지 설정) ← 단계 1 완료 후, 독립 가능
  ↓
단계 8 (통합) ← 모든 단계 완료 후
```

병행 가능 그룹:
- 그룹 A: 단계 2 + 단계 3 (DocInfo 파싱, 서로 독립)
- 그룹 B: 단계 5 + 단계 7 (BodyText 컨트롤, 서로 독립)

## 내결함성 원칙

HWP 바이너리 파싱에서 가장 중요한 원칙은 **모르는 것은 건너뛰고 계속 진행**이다.

1. **레코드 길이 신뢰**: 레코드 내부 파싱에 실패해도 레코드 길이만큼 건너뛰면 다음 레코드를 읽을 수 있음
2. **버전 가변성**: 버전에 따라 레코드 길이가 다를 수 있으므로 "알려진 필드만 읽고 나머지 skip"
3. **누락 허용**: DocInfo에 특정 항목이 없으면 기본값 사용
4. **부분 성공**: 표 파싱에 실패해도 텍스트는 복원, 그림 파싱에 실패해도 표는 복원
5. **fallback 체인**: DocInfo 실패 → 서식 없는 텍스트, BodyText 실패 → PrvText

## 검증 방법

### 테스트 문서

다음 유형의 HWP fixture를 준비한다.

1. **서식 있는 텍스트**: 굵게, 기울임, 글자 크기, 글자색이 적용된 문서
2. **문단 스타일**: 정렬, 들여쓰기, 줄 간격이 다른 문단
3. **표 포함**: 기본 표, rowSpan/colSpan이 있는 표
4. **그림 포함**: 인라인 이미지, floating 이미지
5. **복합 문서**: 텍스트 + 표 + 그림이 혼합된 문서
6. **압축 문서**: 바디 스트림이 zlib 압축된 문서
7. **다양한 버전**: HWP 5.0.x ~ 5.1.x 버전 문서

### 검증 기준

- HWP 문서를 열었을 때 서식(굵기, 크기, 색상)이 반영됨
- 표가 올바른 행/열 구조로 표시됨
- 그림이 올바른 위치에 표시됨
- 페이지 크기와 여백이 반영됨
- 파싱 실패 시 앱이 죽지 않고 가능한 범위까지 표시
- 기존 PrvText fallback이 여전히 동작

### 단위 테스트 전략

기존 테스트 패턴을 따라 `CompoundFile::create`로 인메모리 CFB를 생성하고, 레코드를 수동 조립하여 테스트한다.

```rust
#[test]
fn parses_docinfo_char_shape() {
    // DocInfo 스트림에 HWPTAG_CHAR_SHAPE 레코드를 넣고
    // 파싱 결과가 올바른 TextStyle로 변환되는지 확인
}
```

## 예상 작업량

| 단계 | 핵심 산출물 | 예상 코드량 |
|---|---|---|
| 1. DocInfo 기반 | DocInfoStore, iter_records | ~120줄 |
| 2. 글꼴/글자/문단 | FaceName, CharShape, ParaShape 파서 | ~200줄 |
| 3. BinData | BinDataRef 파서, 자산 로드 | ~80줄 |
| 4. 문단 서식 | PARA_CHAR_SHAPE, 서식 적용 | ~110줄 |
| 5. 표 | CTRL_HEADER, TABLE, LIST_HEADER | ~210줄 |
| 6. 그림/도형 | gso, SHAPE_COMPONENT | ~120줄 |
| 7. 페이지 설정 | PAGE_DEF, Section 분리 | ~45줄 |
| 8. 통합 | parse_bytes 확장, fallback | ~90줄 |
| **합계** | | **~975줄** |

## 위험 요소

1. **공식 문서 불완전성**: HWP 5.0 공식 PDF가 모든 필드를 다루지 않음 → pyhwp, hwpjs 구조 참고로 보완
2. **버전별 차이**: 5.0.0 ~ 5.1.x까지 레코드 구조가 미세하게 다를 수 있음 → 내결함성 원칙으로 대응
3. **압축 문서**: DocInfo도 압축될 수 있음 → 기존 zlib/deflate 감지 로직 재사용
4. **레코드 계층**: level 기반 부모-자식 관계 복원이 복잡할 수 있음 → 스택 기반 파서로 구현
5. **표 셀 매핑**: col_addr/row_addr 기반 2D 배치가 HWPX보다 복잡 → 그리드 맵 방식으로 처리

## 참고

- HWP 파서 현재 코드: `crates/max_viewer_hwp/src/lib.rs`
- 한컴 공식 HWP 5.0 문서 형식 PDF
- pyhwp 프로젝트 (구조 참고, 라이선스 주의)
- hwpjs 프로젝트 (Rust 코어 구조 참고)
- 코어 문서 모델: `crates/max_viewer_core/src/lib.rs`

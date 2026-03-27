# HWP/HWPX 포맷 조사

작성일: 2026-03-26  
조사 목적: MAX-Viewer에서 `.hwp`, `.hwpx` 문서를 직접 읽어 화면에 표시하기 위한 기술적 근거 정리

## 1. 핵심 결론

- `HWPX`는 ZIP + XML 기반이라 파싱 난이도가 상대적으로 낮고, 신규 개발의 시작점으로 적합합니다.
- `HWP`는 Microsoft Compound File Binary Format(CFB) 위에 한컴 고유 레코드 구조가 얹힌 바이너리 포맷이라 난이도가 높습니다.
- 실제 제품은 둘 다 읽어야 하므로, 포맷별 파서를 바로 UI에 연결하지 말고 공통 내부 문서 모델로 정규화해야 합니다.
- 암호화 문서, DRM 문서, 배포용 문서, OLE 개체, 일부 도형/수식은 초기 버전에서 제한 지원으로 두는 편이 현실적입니다.

## 2. 왜 두 포맷을 모두 지원해야 하는가

한컴 FAQ에 따르면 한컴은 2010년부터 HWP 구조 공개 및 HWPX 지원을 시작했고, 2021년에 HWPX를 기본 저장 포맷으로 전환했습니다. 즉, 앞으로 생성되는 문서는 HWPX 비중이 커지겠지만, 기존 공공/기업 문서 자산은 여전히 HWP가 많습니다.

MAX-Viewer가 실사용 도구가 되려면:

- 신규 문서 대응을 위해 `HWPX`가 필요하고
- 레거시 문서 대응을 위해 `HWP`가 필요합니다.

이 때문에 구현 전략은 "HWPX 먼저, HWP는 별도 어댑터로 추가"가 가장 합리적입니다.

## 3. HWPX 포맷 구조

한컴테크의 2025년 HWPX 구조 글에 따르면 HWPX는 국가표준 `KS X 6101`(OWPML)을 따르는 개방형 문서 포맷이며, ZIP 안에 여러 XML과 바이너리 리소스를 담는 구조입니다.

### 3.1 주요 구성요소

- `mimetype`
  - 파일이 HWPX임을 식별하는 시그니처 역할
- `version.xml`
  - OWPML 버전과 저장 환경 정보
- `settings.xml`
  - 커서 위치 등 외부 설정 정보
- `Contents/content.hpf`
  - 패키지 메타데이터, manifest, spine
  - `spine` 순서대로 문서를 읽어야 함
- `Contents/header.xml`
  - 글자 모양, 문단 모양, 번호/스타일 등 문서 공통 매핑 정보
- `Contents/section0.xml`, `section1.xml`, ...
  - 구역별 본문 내용
  - 문단은 `<hp:p>`, 텍스트 run은 `<hp:run>`, 실제 텍스트는 `<hp:t>`에 저장
- `BinData/`
  - 이미지, OLE 등 바이너리 리소스
- `META-INF/`
  - 컨테이너 정보, 암호화 문서 관련 정보
- `Preview/`
  - 미리보기 이미지/텍스트
- `Scripts/`
  - 스크립트 정보

### 3.2 구현 시사점

- ZIP 파서를 이용해 엔트리를 안전하게 읽은 뒤 XML만 순회하면 되므로, 파서 구현 속도가 빠릅니다.
- `content.hpf`의 `spine`을 기준으로 section을 읽어야 문서 순서가 보장됩니다.
- `header.xml`을 먼저 읽어 스타일 테이블을 메모리에 올린 뒤 `sectionN.xml`에서 참조를 해석해야 합니다.
- `BinData`는 본문 파싱과 분리해서 로드하고, 이미지나 OLE placeholder와 연결해야 합니다.
- 초기 버전에서는 `<hp:t>` 기반 텍스트 추출과 문단/문자 스타일 적용만으로도 실용적인 뷰어를 빠르게 만들 수 있습니다.

## 4. HWP 포맷 구조

한컴 공식 PDF와 한컴테크 글을 종합하면 HWP 5.x는 복합 파일(Compound File) 구조를 사용합니다. 다시 말해, 하나의 `.hwp` 파일 안에 storage와 stream이 들어 있고, 그 안쪽 데이터는 다시 한컴 고유 레코드 구조로 저장됩니다.

### 4.1 주요 스토리지/스트림

- `FileHeader`
  - 파일 시그니처, 버전, 속성 플래그
  - 공식 문서 기준 속성 비트에는 압축 여부, 암호 설정 여부, 배포용 문서 여부, 스크립트 저장 여부 등이 포함됨
- `DocInfo`
  - 글꼴, 글자 모양, 문단 모양, 스타일, BinData 참조 정보 등 공통 정보
- `BodyText/Section0`, `Section1`, ...
  - 실제 본문
- `BinData`
  - 이미지, OLE 등 바이너리 데이터
- `PrvText`
  - 미리보기 텍스트
- `PrvImage`
  - 미리보기 이미지
- `DocOptions`, `Scripts`, `XMLTemplate`, `DocHistory`
  - 옵션, 스크립트, XML 템플릿, 이력 정보

### 4.2 HWP 레코드 구조

한컴 공식 HWP 5.0 문서 형식 PDF의 데이터 레코드 설명에 따르면 레코드 헤더는 32비트이며 다음으로 구성됩니다.

- `TagID` 10비트
- `Level` 10비트
- `Size` 12비트

`Size`가 모두 1인 경우(`4095` 이상)에는 뒤에 실제 길이를 담는 `DWORD`가 추가됩니다. 즉, HWP 파서는 "스트림 -> 압축 해제 -> 레코드 헤더 반복 해석 -> Tag/Level 기반 계층 복원" 흐름이 되어야 합니다.

### 4.3 HWP에서 특히 중요한 점

- `FileHeader`의 압축 플래그를 먼저 확인해야 합니다.
- `DocInfo`, `BodyText`, `DocHistory` 등 일부 스트림은 압축 해제 후 처리해야 합니다.
- `DocInfo`에서 글꼴/문단/스타일/BinData 매핑을 먼저 읽어야 `BodyText`를 올바르게 해석할 수 있습니다.
- `BodyText`는 레코드 계층 구조를 따라 문단, 컨트롤, 표, 페이지 설정 등을 재구성해야 합니다.
- 문서 버전에 따라 레코드 길이가 가변적일 수 있으므로, "알려진 필드만 읽고 남은 바이트를 스킵"하는 내결함성이 필수입니다.

## 5. MAX-Viewer에 중요한 구현 포인트

### 5.1 HWPX는 "문서 구조 파서"로 접근

HWPX는 XML 구조가 분명하므로 다음 순서가 적절합니다.

1. ZIP 열기
2. `mimetype` 검증
3. `content.hpf` 파싱
4. `header.xml` 스타일 테이블 로드
5. `sectionN.xml` 순차 파싱
6. `BinData` 연결
7. 내부 문서 모델 생성

### 5.2 HWP는 "컨테이너 + 레코드 파서"로 접근

HWP는 두 계층을 분리해 생각해야 합니다.

1. `CFB/OLE` 컨테이너 읽기
2. `FileHeader` 해석
3. 압축 여부 확인 및 스트림 해제
4. `DocInfo` 레코드 파싱
5. `BodyText/SectionN` 레코드 파싱
6. 스타일/리소스 참조 복원
7. 내부 문서 모델 생성

### 5.3 두 포맷의 공통점

결국 뷰어에 필요한 건 다음과 같은 공통 정보입니다.

- 문서 메타데이터
- 페이지 설정
- 문단 목록
- 텍스트 run
- 문자/문단 스타일
- 표 구조
- 이미지 자산
- 머리글/바닥글
- 각주/미주

그래서 MAX-Viewer는 HWP/HWPX 각각을 직접 화면에 그리기보다, 먼저 같은 구조의 중간 문서 모델로 통합하는 것이 맞습니다.

## 6. 참고할 공개 구현체

### 6.1 `hancom-io/hwpx-owpml-model`

한컴이 공개한 OWPML 모델 저장소입니다. README에 따르면 OWPML 구조 기반으로 문서 엘리먼트를 추출하고 저장할 수 있으며, 텍스트 추출 예제도 포함되어 있습니다.

활용 포인트:

- HWPX 파서 설계 참고
- OWPML 문서 구조 이해
- section 기반 순회 방식 참고

### 6.2 `pyhwp`

`pyhwp`는 HWP v5 파서/프로세서로, 내부 스트림 분석과 텍스트 변환 기능을 제공합니다.

활용 포인트:

- HWP 바이너리 분석 흐름 참고
- DocInfo/BodyText 분리 구조 참고
- 테스트 코퍼스 구성 시 비교 기준으로 활용

주의:

- AGPL 계열 라이선스이므로 직접 코드 차용은 신중해야 하고, 구조 참고 및 테스트 비교 대상으로 보는 편이 안전합니다.

### 6.3 `hwpjs`

`hwpjs`는 HWP를 읽고 JSON/Markdown/HTML로 변환하는 공개 프로젝트입니다. 저장소 설명 기준으로 Rust 핵심 파서와 Web/Node 바인딩 구조를 사용합니다.

활용 포인트:

- "Rust core + Web 패키지" 아키텍처 참고
- 브라우저 뷰어/HTML 변환 접근 참고
- 장기적으로 WASM 배포를 고려할 때 좋은 구조적 참고 사례

### 6.4 LibreOffice `hwpfilter`

LibreOffice 문서에는 한국어 HWP 포맷용 필터가 있지만, 문서 자체에 "새로운 버전의 포맷을 제대로 처리하지 못할 수 있다"는 취지의 경고가 있습니다.

시사점:

- LibreOffice 필터를 MAX-Viewer의 핵심 엔진으로 삼는 것은 위험합니다.
- 다만 회귀 테스트용 비교군, 또는 실패 사례 수집용 참고 대상으로는 의미가 있습니다.

## 7. 추천 구현 전략

### 7.1 우선순위

1. `HWPX 텍스트/스타일/표/이미지` 뷰어 완성
2. 공통 내부 문서 모델 안정화
3. `HWP 텍스트/스타일/표/이미지` 파서 추가
4. 페이지 정밀도, 머리글/바닥글, 각주/미주, 번호 매기기 강화
5. 도형, 수식, OLE, 변경 추적, 암호 문서 대응 확대

### 7.2 초기 버전에서 과감히 제한할 항목

- 암호화 문서 전체 지원
- DRM 문서
- 스크립트 실행
- OLE 편집
- 차트/복합 도형의 완전한 재현
- 문서 편집 기능

## 8. 실무적인 판단

MAX-Viewer는 처음부터 "한컴과 100% 동일한 편집 엔진"을 목표로 잡으면 일정이 통제되지 않습니다. 대신 아래처럼 잡는 것이 현실적입니다.

- 목표 1: 문서를 열 수 있어야 한다.
- 목표 2: 대부분의 공공 문서를 읽을 수 있어야 한다.
- 목표 3: 검색/복사/텍스트 추출이 가능해야 한다.
- 목표 4: 표와 이미지가 크게 무너지지 않아야 한다.
- 목표 5: 페이지 모드 정확도는 점진적으로 끌어올린다.

즉, "읽기 호환성 중심의 뷰어"로 출발하고, 정확도는 샘플 문서와 스냅샷 테스트를 통해 단계적으로 높이는 편이 맞습니다.

## 9. 출처

### 공식/1차 자료

- 한컴 FAQ, HWP 개방성 및 HWPX 기본 포맷 전환  
  <https://recruit.hancom.co.kr/support/faqCenter/faq/detail/3129>
- 한컴테크, HWP 포맷 구조 살펴보기  
  <https://tech.hancom.com/%ED%95%9C-%EA%B8%80-%EB%AC%B8%EC%84%9C-%ED%8C%8C%EC%9D%BC-%ED%98%95%EC%8B%9D-hwp-%ED%8F%AC%EB%A7%B7-%EA%B5%AC%EC%A1%B0-%EC%82%B4%ED%8E%B4%EB%B3%B4%EA%B8%B0/>
- 한컴테크, HWPX 포맷 구조 살펴보기  
  <https://tech.hancom.com/hwpxformat/>
- 한컴 공식 PDF, 한글 문서 파일 형식 5.0  
  <https://cdn.hancom.com/link/docs/%ED%95%9C%EA%B8%80%EB%AC%B8%EC%84%9C%ED%8C%8C%EC%9D%BC%ED%98%95%EC%8B%9D_5.0_revision1.2.pdf>
- Microsoft Learn, MS-CFB  
  <https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b>

### 참고 구현체

- Hancom OWPML model  
  <https://github.com/hancom-io/hwpx-owpml-model>
- pyhwp  
  <https://github.com/mete0r/pyhwp>
- hwpjs  
  <https://github.com/ohah/hwpjs>
- LibreOffice hwpfilter  
  <https://docs.libreoffice.org/hwpfilter.html>

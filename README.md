# PPTX 테이블 파서

PPTX 슬라이드 XML에서 테이블 XML을 추출하고, 이를 그리드 JSON으로 변환하는 유틸리티입니다.

## 스크립트

- `table_extractor/extract_table.py`
  - 슬라이드 XML에서 `<a:tbl>` 노드를 추출합니다.
  - 출력: `table_extractor/extract_results/<slide_stem>_0001.xml`, ...

- `table_parser/parse_table.py`
  - 추출된 테이블 XML을 JSON 그리드 형태로 파싱합니다.
  - 출력: `table_parser/parsing_results/<input_stem>_grid.json`

- `table_parser/tableMaker.py`
  - 테이블 JSON을 Markdown/HTML/CSV로 렌더링합니다.
  - 배치(인자 없음) 실행 시 `table_parser/parsing_results/*.json`을
    `table_parser/tables/*.md`로 변환합니다.

## 요구사항

- Python 3.10+ (권장 3.11+)

## 사용법

### 1) 테이블 XML 추출

특정 파일 지정 실행:

```bash
python3 table_extractor/extract_table.py table_extractor/target_slides/slide2.xml table_extractor/target_slides/slide10.xml
```

인자 없이 실행(기본 모드):

```bash
python3 table_extractor/extract_table.py
```

기본 모드에서는 `table_extractor/target_slides/`의 모든 `*.xml`을 처리합니다.

### 2) 테이블 XML -> JSON 파싱

특정 추출 XML 지정 실행:

```bash
python3 table_parser/parse_table.py table_extractor/extract_results/slide2_0001.xml table_extractor/extract_results/slide10_0001.xml
```

인자 없이 실행(기본 모드):

```bash
python3 table_parser/parse_table.py
```

기본 모드에서는 `table_extractor/extract_results/`의 모든 `*.xml`을 처리합니다.

### 3) JSON -> 테이블 렌더링 (`tableMaker.py`)

단일 JSON 파일 변환:

```bash
python3 table_parser/tableMaker.py table_parser/parsing_results/slide2_0001_grid.json
```

인자 없이 실행(배치 모드):

```bash
python3 table_parser/tableMaker.py
```

배치 모드에서는 `table_parser/parsing_results/*.json`(manifest 제외)을
`table_parser/tables/*.md`로 변환합니다.

## 참고

- 진행 로그는 `SCRIPT`와 `WRITE` 단계로 구분되어 출력됩니다.
- 출력 폴더가 없으면 자동 생성됩니다.
- 각 출력 폴더에는 `manifest.json`이 생성됩니다.

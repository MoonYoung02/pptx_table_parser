# table_extractor

`table_extractor/extract_table.py`는 슬라이드 XML에서 OpenXML 테이블 노드(`<a:tbl>`)를 추출해 개별 XML 파일로 저장하는 모듈입니다.

아래 경로 예시는 저장소 루트에서 실행하는 기준입니다.

## 입력

- 슬라이드 XML 파일 경로(1개 이상)
  - 형식: `*.xml`
  - 예: `python3 table_extractor/extract_table.py table_extractor/target_slides/slide2.xml`
- 입력 인자 생략
  - 동작: `table_extractor/target_slides/slide*.xml` 전체 처리

## 출력

출력 디렉터리: `table_extractor/extract_results/`

- 추출 테이블 XML
  - 파일명 형식: `<slide_stem>_0001.xml`, `<slide_stem>_0002.xml`, ...
  - 예: `slide2_0001.xml`
- 실행 매니페스트
  - `table_extractor/extract_results/manifest.json`
  - 내용: 입력 소스별 추출 개수, 생성 파일 목록, 전체 테이블 개수

## 사용법

### 1) 단일 파일 처리

```bash
python3 table_extractor/extract_table.py table_extractor/target_slides/slide2.xml
```

### 2) 복수 파일 처리

```bash
python3 table_extractor/extract_table.py table_extractor/target_slides/slide2.xml table_extractor/target_slides/slide10.xml
```

### 3) 기본 경로 일괄 처리

```bash
python3 table_extractor/extract_table.py
```

## 종료 코드

- `0`: 정상 완료
- `1`: 입력 파일 없음/파일 경로 오류


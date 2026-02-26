# PPTX Table Parser

Utilities for extracting table XML from PPTX slide XML and converting it to grid JSON.

## Scripts

- `table_extractor/extract_table.py`
  - Extracts `<a:tbl>` nodes from slide XML files.
  - Output: `table_extractor/extract_results/<slide_stem>_0001.xml`, ...

- `table_parser/parse_table.py`
  - Parses extracted table XML into JSON grid format.
  - Output: `table_parser/parsing_results/<input_stem>_grid.json`

- `table_parser/tableMaker.py`
  - Renders table JSON into Markdown/HTML/CSV.
  - In batch mode (no input argument), converts `table_parser/parsing_results/*.json`
    into `table_parser/tables/*.md`.

## Requirements

- Python 3.10+ (3.11+ recommended)

## Usage

### 1) Extract table XML

Run with specific files:

```bash
python3 table_extractor/extract_table.py table_extractor/target_slides/slide2.xml table_extractor/target_slides/slide10.xml
```

Run with no arguments (default mode):

```bash
python3 table_extractor/extract_table.py
```

Default mode reads all `*.xml` in `table_extractor/target_slides/`.

### 2) Parse table XML to JSON

Run with specific extracted XML files:

```bash
python3 table_parser/parse_table.py table_extractor/extract_results/slide2_0001.xml table_extractor/extract_results/slide10_0001.xml
```

Run with no arguments (default mode):

```bash
python3 table_parser/parse_table.py
```

Default mode reads all `*.xml` in `table_extractor/extract_results/`.

### 3) Render JSON to table (`tableMaker.py`)

Convert a single JSON file:

```bash
python3 table_parser/tableMaker.py table_parser/parsing_results/slide2_0001_grid.json
```

Run with no arguments (batch mode):

```bash
python3 table_parser/tableMaker.py
```

Batch mode converts `table_parser/parsing_results/*.json` (excluding `manifest.json`)
to `table_parser/tables/*.md`.

## Notes

- Progress logs are shown in `SCRIPT` and `WRITE` stages.
- Output folders are created automatically if missing.
- `manifest.json` is generated in each output folder.

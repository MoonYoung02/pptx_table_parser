#!/usr/bin/env python3
"""
End-to-end Surya pipeline for all PPTX files in target_pptx.

Pipeline:
1) convert_pptx_to_pdf.py: target_pptx -> target_pdf
2) surya_layout on each PDF
3) surya_ocr on each PDF
4) surya_table on each PDF
5) normalize_surya_results.py for each stem
6) build_structure_ready_from_normalized.py for each stem
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path
from typing import Dict, List


def run_cmd(cmd: List[str], cwd: Path, allow_fail: bool = False) -> bool:
    proc = subprocess.run(cmd, cwd=str(cwd), text=True, capture_output=True)
    if proc.returncode != 0:
        if allow_fail:
            print(
                "[WARN] command failed but continuing\n"
                f"cmd: {' '.join(cmd)}\n"
                f"stdout:\n{proc.stdout}\n"
                f"stderr:\n{proc.stderr}"
            )
            return False
        raise RuntimeError(
            "command failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )
    if proc.stdout.strip():
        print(proc.stdout.strip())
    return True


def run_surya_cli(kind: str, input_path: Path, output_dir: Path, cwd: Path) -> None:
    cli_map = {
        "layout": ("surya.scripts.detect_layout", "detect_layout_cli"),
        "ocr": ("surya.scripts.ocr_text", "ocr_text_cli"),
        "table": ("surya.scripts.table_recognition", "table_recognition_cli"),
    }
    if kind not in cli_map:
        raise ValueError(f"unsupported surya cli kind: {kind}")
    module_name, cli_name = cli_map[kind]
    launcher = (
        "import sys; "
        f"from {module_name} import {cli_name} as cli; "
        "cli.main(args=sys.argv[1:], standalone_mode=False)"
    )
    run_cmd(
        [
            sys.executable,
            "-c",
            launcher,
            str(input_path),
            "--output_dir",
            str(output_dir),
        ],
        cwd=cwd,
    )


def prettify_json_file(json_path: Path) -> None:
    if not json_path.exists() or not json_path.is_file():
        return
    try:
        data = json.loads(json_path.read_text(encoding="utf-8"))
    except Exception:
        return
    json_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Run convert + Surya + normalize for target_pptx/*.pptx")
    parser.add_argument("--surya-dir", default=".", help="Path to surya_pipeline directory.")
    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-run layout/table/normalize even if output files already exist.",
    )
    args = parser.parse_args()

    surya_dir = Path(args.surya_dir).resolve()
    target_pptx = surya_dir / "target_pptx"
    target_pdf = surya_dir / "target_pdf"
    output_layout = surya_dir / "output" / "layout_result"
    output_ocr = surya_dir / "output" / "ocr"
    output_table = surya_dir / "output" / "table"
    output_norm = surya_dir / "output" / "normalized"
    output_struct = surya_dir / "output" / "structure_ready"

    for required_dir in [
        target_pptx,
        target_pdf,
        output_layout,
        output_ocr,
        output_table,
        output_norm,
        output_struct,
    ]:
        if required_dir.exists() and not required_dir.is_dir():
            raise NotADirectoryError(f"required path exists but is not a directory: {required_dir}")
        required_dir.mkdir(parents=True, exist_ok=True)

    # Step 1: PPTX -> PDF
    print("[1/4] Converting PPTX to PDF...")
    run_cmd([sys.executable, "convert_pptx_to_pdf.py"], cwd=surya_dir, allow_fail=True)

    # Build stem -> pptx map
    pptx_map: Dict[str, Path] = {}
    for p in sorted(target_pptx.glob("*.pptx")):
        pptx_map[p.stem] = p

    pdf_files = sorted(target_pdf.glob("*.pdf"))
    if not pdf_files:
        print("[INFO] No PDF files found in target_pdf after conversion.")
        return 0

    try:
        __import__("surya")
    except Exception as e:  # noqa: BLE001
        raise RuntimeError("surya package not importable from current Python. Activate the project venv first.") from e

    # Step 2 + 3 + 4 + 5 + 6 per PDF
    for idx, pdf_path in enumerate(pdf_files, start=1):
        stem = pdf_path.stem
        print(f"[PDF {idx}/{len(pdf_files)}] {pdf_path.name}")

        layout_json = output_layout / stem / "results.json"
        ocr_json = output_ocr / stem / "results.json"
        table_json = output_table / stem / "results.json"
        normalized_json = output_norm / f"{stem}_normalized.json"
        structure_dir = output_struct / stem
        structure_manifest = structure_dir / "structure_analysis_manifest.json"

        # Step 2: layout
        if args.force or not layout_json.exists():
            print("  [2/5] Running surya_layout...")
            run_surya_cli("layout", pdf_path, output_layout, cwd=surya_dir)
        else:
            print("  [2/5] Skip surya_layout (exists)")
        prettify_json_file(layout_json)

        # Step 3: ocr
        if args.force or not ocr_json.exists():
            print("  [3/5] Running surya_ocr...")
            run_surya_cli("ocr", pdf_path, output_ocr, cwd=surya_dir)
        else:
            print("  [3/5] Skip surya_ocr (exists)")
        prettify_json_file(ocr_json)

        # Step 4: table
        if args.force or not table_json.exists():
            print("  [4/5] Running surya_table...")
            run_surya_cli("table", pdf_path, output_table, cwd=surya_dir)
        else:
            print("  [4/5] Skip surya_table (exists)")
        prettify_json_file(table_json)

        # Step 5: normalize
        pptx_path = pptx_map.get(stem)
        if pptx_path is None:
            print(f"  [5/6] Skip normalize/build (matching pptx not found for stem: {stem})")
            continue
        if args.force or not normalized_json.exists():
            print("  [5/6] Running normalize_surya_results.py...")
            run_cmd(
                [
                    sys.executable,
                    "normalize_surya_results.py",
                    "--layout-json",
                    str(layout_json),
                    "--table-json",
                    str(table_json),
                    "--ocr-json",
                    str(ocr_json),
                    "--pptx-path",
                    str(pptx_path),
                    "--output-json",
                    str(normalized_json),
                ],
                cwd=surya_dir,
            )
        else:
            print("  [5/6] Skip normalize (exists)")
        prettify_json_file(normalized_json)

        # Step 6: structure-ready
        if args.force or not structure_manifest.exists():
            print("  [6/6] Running build_structure_ready_from_normalized.py...")
            run_cmd(
                [
                    sys.executable,
                    "build_structure_ready_from_normalized.py",
                    "--normalized-json",
                    str(normalized_json),
                    "--pptx-path",
                    str(pptx_path),
                    "--output-dir",
                    str(structure_dir),
                ],
                cwd=surya_dir,
            )
        else:
            print("  [6/6] Skip structure-ready (exists)")
        prettify_json_file(structure_manifest)

    print("[DONE] Pipeline finished.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

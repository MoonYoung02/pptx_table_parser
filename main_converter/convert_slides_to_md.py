#!/usr/bin/env python3
"""
Convert extracted PPTX package(s) to markdown.

Usage:
  python convert_slides_to_md.py [ppt_root_dir ...]

Rules:
  - If no positional args are provided, process all package roots in ./target_pptx.
  - Package root example: ./target_pptx/sample1
  - Each package must contain: ./ppt/slides
  - Output is always written to ./output (created automatically).
  - Input slide XML order is assumed to be the final reading order.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def run_structure_analysis_stage(
    repo_root: Path,
    slide_xmls: Sequence[Path],
    mode: str,
    surya_dir: Optional[Path],
) -> Tuple[Dict[str, Path], Path]:
    if mode == "surya" and surya_dir is None:
        raise ValueError("structure-analysis mode 'surya' requires --surya-dir")

    ro_script = repo_root / "structure_analyzer" / "extract_reading_order_slide.py"
    if not ro_script.exists():
        raise FileNotFoundError(f"reading_order script not found: {ro_script}")

    ro_output = Path(tempfile.mkdtemp(prefix=f"struct_analysis_{mode}_", dir="/tmp"))
    cmd = [
        "python3",
        str(ro_script),
        "--mode",
        mode,
        "--output-dir",
        str(ro_output),
    ]
    if mode == "surya" and surya_dir is not None:
        cmd.extend(["--surya-dir", str(surya_dir.resolve())])
    cmd.extend(str(p.resolve()) for p in slide_xmls)

    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(
            "structure_analysis stage failed\n"
            f"cmd: {' '.join(cmd)}\n"
            f"stdout:\n{proc.stdout}\n"
            f"stderr:\n{proc.stderr}"
        )

    manifest_path = ro_output / "structure_analysis_manifest.json"
    if not manifest_path.exists():
        legacy_manifest = ro_output / "reading_order_manifest.json"
        if legacy_manifest.exists():
            manifest_path = legacy_manifest
        else:
            raise FileNotFoundError(f"structure_analysis manifest not found: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    failed = manifest.get("failed", [])
    if isinstance(failed, list) and failed:
        raise RuntimeError(f"structure_analysis stage reported failures: {json.dumps(failed, ensure_ascii=False)}")

    mapping: Dict[str, Path] = {}
    processed = manifest.get("processed", [])
    if isinstance(processed, list):
        for row in processed:
            if not isinstance(row, dict):
                continue
            src = row.get("input_xml")
            out = row.get("output_xml")
            if isinstance(src, str) and isinstance(out, str):
                mapping[str(Path(src).resolve())] = Path(out).resolve()
    return mapping, ro_output


def ensure_imports(repo_root: Path) -> None:
    # Support both old and new repository layouts.
    candidates = [
        repo_root / "table_parser",
        repo_root / "pptx_table_parser" / "table_parser",
        repo_root / "pptx_table_parser",
    ]
    for cand in candidates:
        if cand.exists() and cand.is_dir() and str(cand) not in sys.path:
            sys.path.insert(0, str(cand))


def default_target_dirs(cwd: Path) -> List[Path]:
    # Prefer local main_converter/target_pptx, then fallback to parent-level.
    cands = [
        cwd / "target_pptx",
        cwd.parent / "target_pptx",
    ]
    out: List[Path] = []
    for c in cands:
        if c.exists() and c.is_dir():
            out.append(c)
    return out


def natural_key(name: str) -> Tuple:
    parts = re.split(r"(\d+)", name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def normalize_text(s: str) -> str:
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def pick_packages(target_dirs: Sequence[Path], raw_inputs: Sequence[str]) -> List[Path]:
    pkgs: List[Path] = []
    if raw_inputs:
        for item in raw_inputs:
            p = Path(item)
            candidates = [p]
            for td in target_dirs:
                candidates.append(td / item)
            picked = None
            for c in candidates:
                if c.exists() and c.is_dir():
                    picked = c.resolve()
                    break
            if picked is not None:
                pkgs.append(picked)
    else:
        for target_dir in target_dirs:
            for p in sorted(target_dir.iterdir(), key=lambda x: natural_key(x.name)):
                if p.is_dir():
                    pkgs.append(p.resolve())

    out: List[Path] = []
    uniq: Dict[str, Path] = {}
    for pkg in pkgs:
        slides_dir = pkg / "ppt" / "slides"
        if slides_dir.exists() and slides_dir.is_dir():
            uniq[str(pkg)] = pkg
    for _, v in uniq.items():
        out.append(v)
    return out


def parse_slide_number(filename: str, default_idx: int) -> int:
    m = re.search(r"slide(\d+)", filename, re.IGNORECASE)
    if m:
        return int(m.group(1))
    return default_idx


def slide_base_name(slide_xml: Path) -> str:
    m = re.search(r"(slide\d+)", slide_xml.stem, re.IGNORECASE)
    if m:
        return f"{m.group(1)}.xml"
    return slide_xml.name


def find_sidecar_json(slide_xml: Path) -> Optional[Path]:
    # e.g. slide2.reordered.xml -> slide2.structure_analysis.json
    stem = slide_xml.stem
    candidates = []
    if stem.endswith(".reordered"):
        candidates.append(slide_xml.with_name(f"{stem[:-len('.reordered')]}.structure_analysis.json"))
        candidates.append(slide_xml.with_name(f"{stem[:-len('.reordered')]}.reading_order.json"))
    candidates.append(slide_xml.with_name(f"{stem}.structure_analysis.json"))
    candidates.append(slide_xml.with_name(f"{stem}.reading_order.json"))
    for c in candidates:
        if c.exists() and c.is_file():
            return c
    return None


def load_heading_hints(slide_xml: Path) -> Dict[str, Dict[str, object]]:
    """
    Load heading hints produced by structure analysis stage.
    key: shape_id (string)
    value: subset of hint fields
    """
    sidecar = find_sidecar_json(slide_xml)
    if sidecar is None:
        return {}
    try:
        payload = json.loads(sidecar.read_text(encoding="utf-8"))
    except Exception:
        return {}
    rows = payload.get("structure_order")
    if not isinstance(rows, list):
        rows = payload.get("reading_order")
    if not isinstance(rows, list):
        return {}
    out: Dict[str, Dict[str, object]] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        sid = str(row.get("shape_id", "")).strip()
        if not sid:
            continue
        out[sid] = {
            "is_heading_candidate": bool(row.get("is_heading_candidate", False)),
            "heading_score": float(row.get("heading_score", 0.0)),
            "heading_depth_hint": row.get("heading_depth_hint"),
        }
    return out


def rels_from_sidecar(slide_xml: Path) -> Optional[Path]:
    sidecar = find_sidecar_json(slide_xml)
    if sidecar is None:
        return None
    try:
        payload = json.loads(sidecar.read_text(encoding="utf-8"))
    except Exception:
        return None
    src = payload.get("input_xml")
    if not isinstance(src, str) or not src:
        return None
    src_path = Path(src)
    if not src_path.exists():
        return None
    rels = src_path.parent / "_rels" / f"{src_path.name}.rels"
    if rels.exists():
        return rels

    # Fallback: try finding the original slide XML under */ppt/slides/.
    base_name = slide_base_name(slide_xml)
    repo_root = Path(__file__).resolve().parent.parent
    for cand in repo_root.rglob(base_name):
        if "/ppt/slides/" not in cand.as_posix():
            continue
        rels2 = cand.parent / "_rels" / f"{cand.name}.rels"
        if rels2.exists():
            return rels2
    return None


def fallback_rels(slide_xml: Path) -> Optional[Path]:
    base = slide_base_name(slide_xml)
    cands = [
        slide_xml.parent / "_rels" / f"{base}.rels",
        slide_xml.with_name(f"{base}.rels"),
        slide_xml.parent / "_rels" / f"{slide_xml.name}.rels",
    ]
    for c in cands:
        if c.exists() and c.is_file():
            return c
    return None


def build_rels_map(rels_path: Optional[Path]) -> Dict[str, str]:
    if rels_path is None or not rels_path.exists():
        return {}
    root = ET.parse(rels_path).getroot()
    out: Dict[str, str] = {}
    for rel in root.findall("rel:Relationship", REL_NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            out[rid] = target
    return out


def choose_rels_in_package(slide_xml: Path) -> Optional[Path]:
    # Strictly stay inside same ppt package to avoid cross-package mismatches.
    sidecar = rels_from_sidecar(slide_xml)
    if sidecar:
        return sidecar
    return fallback_rels(slide_xml)


def resolve_image_path(
    slide_xml: Path,
    rels_map: Dict[str, str],
    rels_path: Optional[Path],
    r_embed: Optional[str],
) -> Tuple[str, Optional[str]]:
    if not r_embed:
        return "[unresolved-image]", "missing r:embed"
    if rels_path is None:
        return f"[unresolved-image:{r_embed}]", "missing slide rels file"
    if not rels_map:
        return f"[unresolved-image:{r_embed}]", "empty or unreadable rels map"
    target = rels_map.get(r_embed)
    if not target:
        return f"[unresolved-image:{r_embed}]", f"relationship not found: {r_embed}"

    abs_path = (rels_path.parent.parent / target).resolve()

    cwd = Path.cwd().resolve()
    try:
        rel = abs_path.relative_to(cwd)
        return rel.as_posix(), None
    except ValueError:
        return str(abs_path), None


def extract_shape_text(shape_elem: ET.Element) -> str:
    paragraphs: List[str] = []
    for p in shape_elem.findall(".//p:txBody/a:p", NS):
        runs = []
        for t in p.findall(".//a:t", NS):
            if t.text:
                runs.append(t.text.strip())
        text = normalize_text(" ".join(x for x in runs if x))
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs).strip()


def infer_heading_depth_fallback(text: str, text_block_index: int) -> Optional[int]:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return None
    if raw.startswith(("▶", "-", "*", "√")):
        return None
    if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", raw):
        return 3
    if re.match(r"^\d+[.)]\s+", raw):
        return 2
    # First meaningful text on a slide is often the slide title.
    if text_block_index == 0 and len(raw) <= 80:
        return 1
    return None


def shape_id_of(elem: ET.Element) -> str:
    c_nv_pr = elem.find(".//p:cNvPr", NS)
    if c_nv_pr is None:
        return ""
    return c_nv_pr.attrib.get("id", "")


def convert_table_to_markdown(graphic_frame: ET.Element) -> Tuple[Optional[str], Optional[str]]:
    tbl = graphic_frame.find(".//a:tbl", NS)
    if tbl is None:
        return None, "graphicFrame without a:tbl"

    import parse_table  # type: ignore
    import tableMaker  # type: ignore

    with tempfile.NamedTemporaryFile("wb", suffix=".xml", delete=True) as tmp:
        tmp.write(ET.tostring(tbl, encoding="utf-8"))
        tmp.flush()
        parsed = parse_table.parse_table_xml(Path(tmp.name))
        dense = tableMaker._dense_grid_from_parsed_table(parsed, fill_merged=tableMaker.FILL_BOTH)
        md = tableMaker._render_markdown_flat(dense=dense, header_rows=1, use_header_rows=True)
    return md, None


def convert_one_slide(
    slide_xml: Path,
    page_no: int,
) -> Tuple[str, Dict[str, object]]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        raise ValueError("missing p:cSld/p:spTree")

    rels_path = choose_rels_in_package(slide_xml)
    rels_map = build_rels_map(rels_path)
    heading_hints = load_heading_hints(slide_xml)

    lines: List[str] = [f"[Page_{page_no}]", ""]
    used_headings: set = set()
    text_block_index = 0

    stats = {
        "blocks_total": 0,
        "text_blocks": 0,
        "image_blocks": 0,
        "table_blocks": 0,
        "unsupported_blocks": 0,
        "skipped_blocks": 0,
        "resolved_images": 0,
        "unresolved_images": 0,
        "warnings": [],
        "rels_path": str(rels_path) if rels_path else None,
    }

    for child in list(sp_tree):
        tag = local_name(child.tag)
        if tag not in {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}:
            continue
        stats["blocks_total"] += 1

        if tag == "cxnSp":
            stats["skipped_blocks"] += 1
            continue

        if tag in {"sp", "grpSp"}:
            ph = child.find(".//p:ph", NS)
            ph_type = ph.attrib.get("type") if ph is not None else None
            if ph_type in {"sldNum", "ftr", "dt"}:
                stats["skipped_blocks"] += 1
                continue
            text = extract_shape_text(child)
            if not text:
                stats["skipped_blocks"] += 1
                continue
            if re.fullmatch(r"\d+", text):
                stats["skipped_blocks"] += 1
                continue
            sid = shape_id_of(child)
            hint = heading_hints.get(sid, {})
            depth = hint.get("heading_depth_hint")
            score = float(hint.get("heading_score", 0.0))
            is_candidate = bool(hint.get("is_heading_candidate", False))

            if not is_candidate:
                fb_depth = infer_heading_depth_fallback(text, text_block_index)
                if fb_depth is not None:
                    depth = fb_depth
                    score = 0.8
                    is_candidate = True

            rendered = text
            # Final markdown heading level rendering is converter responsibility.
            if is_candidate and isinstance(depth, int) and 1 <= depth <= 6 and score >= 0.7:
                key = normalize_text(text)
                if key not in used_headings:
                    rendered = f"{'#' * depth} {text}"
                    used_headings.add(key)
                else:
                    # Deduplicate repeated heading text.
                    stats["skipped_blocks"] += 1
                    continue

            lines.append(rendered)
            lines.append("")
            stats["text_blocks"] += 1
            text_block_index += 1
            continue

        if tag == "pic":
            blip = child.find(".//a:blip", NS)
            embed = blip.attrib.get(f"{{{NS['r']}}}embed") if blip is not None else None
            img_path, warn = resolve_image_path(slide_xml, rels_map, rels_path, embed)
            if warn:
                stats["unresolved_images"] += 1
                stats["warnings"].append(warn)
            else:
                stats["resolved_images"] += 1
            lines.append(f'[img(src="{img_path}")]')
            lines.append("")
            stats["image_blocks"] += 1
            continue

        if tag == "graphicFrame":
            table_md, err = convert_table_to_markdown(child)
            if table_md is not None:
                lines.append(table_md.strip())
                lines.append("")
                stats["table_blocks"] += 1
            else:
                lines.append("[unsupported: graphicFrame(non-table)]")
                lines.append("")
                stats["unsupported_blocks"] += 1
                if err:
                    stats["warnings"].append(err)
            continue

    md_text = "\n".join(lines).rstrip() + "\n"
    return md_text, stats


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Convert extracted PPTX package(s) to markdown."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="Package root dir(s). If omitted, process ./target_pptx/*.",
    )
    parser.add_argument(
        "--reading-order",
        choices=("xml", "surya"),
        default="xml",
        help="Reading-order strategy. Default uses legacy XML-only ordering.",
    )
    parser.add_argument(
        "--surya-dir",
        default=None,
        help="Directory containing Surya per-slide JSON outputs (required when --reading-order surya).",
    )
    args = parser.parse_args()

    cwd = Path.cwd()
    output_dir = cwd / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    repo_root = Path(__file__).resolve().parent.parent
    ensure_imports(repo_root)
    surya_dir = Path(args.surya_dir).resolve() if args.surya_dir else None

    target_dirs = default_target_dirs(cwd)
    packages = pick_packages(target_dirs, args.inputs)
    if not packages:
        print("No valid PPTX package directories found.")
        if target_dirs:
            print("Checked default directories:")
            for d in target_dirs:
                print(f"- {d.resolve()}")
        else:
            print("Checked default directories: none found (expected ./target_pptx).")
        return 0

    manifest = {
        "started_at": datetime.utcnow().isoformat() + "Z",
        "packages": [],
        "summary": {
            "processed_packages": 0,
            "processed_slides": 0,
            "failed": 0,
            "resolved_images": 0,
            "unresolved_images": 0,
            "table_blocks": 0,
        },
    }

    for pkg in packages:
        pkg_name = pkg.name
        slides_dir = pkg / "ppt" / "slides"
        slide_xmls = sorted(
            [p for p in slides_dir.glob("slide*.xml") if p.is_file()],
            key=lambda p: natural_key(p.name),
        )
        pkg_out = output_dir / pkg_name
        per_slide_dir = pkg_out / "per_slide"
        pkg_out.mkdir(parents=True, exist_ok=True)
        per_slide_dir.mkdir(parents=True, exist_ok=True)

        pkg_row = {
            "package": str(pkg),
            "name": pkg_name,
            "slides": [],
            "result_md": str(pkg_out / "result.md"),
            "structure_analysis_mode": args.reading_order,
        }

        all_chunks: List[str] = []
        try:
            ro_map, ro_output = run_structure_analysis_stage(
                repo_root=repo_root,
                slide_xmls=slide_xmls,
                mode=args.reading_order,
                surya_dir=surya_dir,
            )
            pkg_row["structure_analysis_output_dir"] = str(ro_output)
        except Exception as e:  # noqa: BLE001
            for slide_xml in slide_xmls:
                row = {
                    "page": parse_slide_number(slide_xml.name, 0),
                    "source_xml": str(slide_xml),
                    "status": "failed",
                    "error": f"structure_analysis stage failed: {e}",
                    "warnings": [],
                }
                pkg_row["slides"].append(row)
                manifest["summary"]["failed"] += 1
            manifest["packages"].append(pkg_row)
            print(f"[{pkg_name}] structure_analysis failed: {e}")
            continue

        for i, slide_xml in enumerate(slide_xmls, 1):
            page_no = parse_slide_number(slide_xml.name, i)
            ordered_slide_xml = ro_map.get(str(slide_xml.resolve()), slide_xml)
            row = {
                "page": page_no,
                "source_xml": str(slide_xml),
                "structure_analysis_xml": str(ordered_slide_xml),
                "status": "ok",
                "warnings": [],
            }
            try:
                md_text, stats = convert_one_slide(ordered_slide_xml, page_no)
                out_md = per_slide_dir / f"{slide_xml.stem}.md"
                out_md.write_text(md_text, encoding="utf-8")
                all_chunks.append(md_text.rstrip())

                row.update(
                    {
                        "status": "ok",
                        "output_md": str(out_md),
                        "blocks_total": stats["blocks_total"],
                        "text_blocks": stats["text_blocks"],
                        "image_blocks": stats["image_blocks"],
                        "table_blocks": stats["table_blocks"],
                        "unsupported_blocks": stats["unsupported_blocks"],
                        "skipped_blocks": stats["skipped_blocks"],
                        "rels_path": stats["rels_path"],
                        "warnings": stats["warnings"],
                    }
                )
                manifest["summary"]["processed_slides"] += 1
                manifest["summary"]["resolved_images"] += stats["resolved_images"]
                manifest["summary"]["unresolved_images"] += stats["unresolved_images"]
                manifest["summary"]["table_blocks"] += stats["table_blocks"]
                print(f"[{pkg_name}] Processed: {slide_xml.name}")
            except Exception as e:  # noqa: BLE001
                row.update({"status": "failed", "error": str(e)})
                manifest["summary"]["failed"] += 1
                print(f"[{pkg_name}] Failed: {slide_xml.name} -> {e}")
            pkg_row["slides"].append(row)

        merged = "\n\n".join(all_chunks).strip()
        if merged:
            merged += "\n"
        merged_path = pkg_out / "result.md"
        merged_path.write_text(merged, encoding="utf-8")
        manifest["summary"]["processed_packages"] += 1
        manifest["packages"].append(pkg_row)

    manifest["finished_at"] = datetime.utcnow().isoformat() + "Z"
    manifest_path = output_dir / "convert_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote package outputs under: {output_dir.resolve()}")
    print(f"Wrote: {manifest_path.resolve()}")
    print(
        "Summary: "
        f"packages={manifest['summary']['processed_packages']} "
        f"slides={manifest['summary']['processed_slides']} "
        f"failed={manifest['summary']['failed']} "
        f"tables={manifest['summary']['table_blocks']} "
        f"images_resolved={manifest['summary']['resolved_images']} "
        f"images_unresolved={manifest['summary']['unresolved_images']}"
    )
    return 1 if manifest["summary"]["failed"] > 0 else 0


if __name__ == "__main__":
    raise SystemExit(main())

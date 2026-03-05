#!/usr/bin/env python3
"""
Build reordered slide XML + structure-analysis sidecars from normalized Surya output.
"""

from __future__ import annotations

import argparse
from difflib import SequenceMatcher
import json
import os
import re
import shutil
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def parse_int(v: Any, default: int = 10**18) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def parse_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def canonical_text(s: str) -> str:
    text = normalize_text(s).lower()
    text = re.sub(r"^\d+(?:[.)]|\.\d+)*\s*", "", text)
    text = re.sub(r"[^0-9a-zA-Z가-힣]+", "", text)
    return text


def text_similarity(a: str, b: str) -> float:
    left = canonical_text(a)
    right = canonical_text(b)
    if not left or not right:
        return 0.0
    if left == right:
        return 1.0
    if left in right or right in left:
        shorter = min(len(left), len(right))
        longer = max(len(left), len(right))
        if longer > 0:
            return max(0.85, shorter / longer)
    return SequenceMatcher(None, left, right).ratio()


def extract_pptx_to_temp_root(pptx_path: Path) -> Tuple[Path, Path]:
    if not pptx_path.exists() or not pptx_path.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx_path}")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_xml_"))
    extracted_root = temp_dir / pptx_path.stem
    extracted_root.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(extracted_root)
    return temp_dir, extracted_root


def copy_slide_relationship_bundle(slide_xml: Path, output_dir: Path) -> Optional[Path]:
    rels_src = slide_xml.parent / "_rels" / f"{slide_xml.name}.rels"
    if not rels_src.exists() or not rels_src.is_file():
        return None

    ppt_root = slide_xml.parent.parent
    rels_dst_dir = output_dir / "_rels"
    rels_dst_dir.mkdir(parents=True, exist_ok=True)
    rels_dst = rels_dst_dir / f"{slide_xml.name}.rels"

    tree = ET.parse(rels_src)
    root = tree.getroot()
    for rel in root.findall("rel:Relationship", REL_NS):
        target = rel.attrib.get("Target")
        target_mode = rel.attrib.get("TargetMode")
        if not target or target_mode == "External":
            continue
        resolved = Path(os.path.normpath(str(slide_xml.parent / target)))
        if not resolved.exists() or not resolved.is_file():
            continue
        try:
            asset_rel = resolved.relative_to(ppt_root)
        except ValueError:
            asset_rel = Path(resolved.name)
        copied_asset = output_dir / asset_rel
        copied_asset.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(resolved, copied_asset)
        rel.attrib["Target"] = asset_rel.as_posix()

    tree.write(rels_dst, encoding="utf-8", xml_declaration=True)
    return rels_dst


def get_nvpr_paths(tag: str) -> Tuple[str, str]:
    if tag == "sp":
        return "./p:nvSpPr/p:cNvPr", "./p:nvSpPr/p:nvPr/p:ph"
    if tag == "pic":
        return "./p:nvPicPr/p:cNvPr", "./p:nvPicPr/p:nvPr/p:ph"
    if tag == "graphicFrame":
        return "./p:nvGraphicFramePr/p:cNvPr", "./p:nvGraphicFramePr/p:nvPr/p:ph"
    if tag == "grpSp":
        return "./p:nvGrpSpPr/p:cNvPr", "./p:nvGrpSpPr/p:nvPr/p:ph"
    if tag == "cxnSp":
        return "./p:nvCxnSpPr/p:cNvPr", "./p:nvCxnSpPr/p:nvPr/p:ph"
    return ".//p:cNvPr", ".//p:ph"


def first_off(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:off",
        "./p:grpSpPr/a:xfrm/a:off",
        "./p:xfrm/a:off",
        ".//a:off",
    ):
        off = elem.find(p, NS)
        if off is not None:
            return off
    return None


def first_ext(elem: ET.Element) -> Optional[ET.Element]:
    for p in (
        "./p:spPr/a:xfrm/a:ext",
        "./p:grpSpPr/a:xfrm/a:ext",
        "./p:xfrm/a:ext",
        ".//a:ext",
    ):
        ext = elem.find(p, NS)
        if ext is not None:
            return ext
    return None


def extract_bbox(elem: ET.Element) -> Optional[List[float]]:
    off = first_off(elem)
    ext = first_ext(elem)
    if off is None or ext is None:
        return None
    x = parse_float(off.attrib.get("x"))
    y = parse_float(off.attrib.get("y"))
    w = parse_float(ext.attrib.get("cx"))
    h = parse_float(ext.attrib.get("cy"))
    if w <= 0 or h <= 0:
        return None
    return [x, y, x + w, y + h]


def extract_shape_text(elem: ET.Element) -> str:
    parts: List[str] = []
    for t in elem.findall(".//a:t", NS):
        if t.text and t.text.strip():
            parts.append(t.text.strip())
    return normalize_text(" ".join(parts))


def extract_font_pt(elem: ET.Element) -> Optional[float]:
    sizes: List[float] = []
    for tag in ("a:rPr", "a:endParaRPr"):
        for node in elem.findall(f".//{tag}", NS):
            sz = node.attrib.get("sz")
            if sz is None:
                continue
            val = parse_float(sz, -1.0)
            if val > 0:
                sizes.append(val / 100.0)
    if not sizes:
        return None
    return max(sizes)


def parse_slide_xml_objects(slide_xml: Path) -> List[Dict[str, Any]]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []

    out: List[Dict[str, Any]] = []
    xml_index = 0
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        xml_index += 1
        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = ch.find(c_nv_path, NS)
        ph = ch.find(ph_path, NS)
        bbox = extract_bbox(ch)
        off = first_off(ch)
        x = parse_int(off.attrib.get("x")) if off is not None else parse_int(bbox[0] if bbox else None)
        y = parse_int(off.attrib.get("y")) if off is not None else parse_int(bbox[1] if bbox else None)
        out.append(
            {
                "shape_id": c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else "",
                "xml_index": xml_index,
                "tag": tag,
                "name": c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else "",
                "ph_type": ph.attrib.get("type") if ph is not None else None,
                "ph_idx": ph.attrib.get("idx") if ph is not None else None,
                "x": x,
                "y": y,
                "bbox": bbox,
                "text": extract_shape_text(ch),
                "font_pt": extract_font_pt(ch),
                "is_footer": (ph.attrib.get("type") if ph is not None else None) in {"sldNum", "ftr", "dt"},
                "is_decorative": tag == "cxnSp" or (not extract_shape_text(ch) and tag not in {"pic", "graphicFrame"}),
                "is_title_placeholder": (ph.attrib.get("type") if ph is not None else None) in TITLE_TYPES,
            }
        )
    return out


def load_normalized_pages(normalized_json: Path) -> Dict[int, Dict[str, Any]]:
    payload = load_json(normalized_json)
    pages = payload.get("pages", [])
    out: Dict[int, Dict[str, Any]] = {}
    if not isinstance(pages, list):
        return out
    for row in pages:
        if not isinstance(row, dict):
            continue
        page_no = parse_int(row.get("page"), 0)
        if page_no > 0:
            out[page_no] = row
    return out


def choose_unique_shape_matches(reading_order: Sequence[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen: set[str] = set()
    out: List[Dict[str, Any]] = []
    for row in reading_order:
        if not isinstance(row, dict):
            continue
        shape_id = str(row.get("matched_shape_id") or "").strip()
        if not shape_id or shape_id in seen:
            continue
        seen.add(shape_id)
        out.append(row)
    return out


def heading_fields_from_surya(row: Dict[str, Any]) -> Dict[str, Any]:
    decision = row.get("heading_decision")
    if not isinstance(decision, dict):
        return {
            "is_heading": False,
            "is_heading_candidate": False,
            "heading_score": 0.0,
            "heading_depth_hint": None,
            "reason": "Surya heading decision unavailable",
        }
    is_heading = bool(decision.get("is_heading", False))
    level = parse_int(decision.get("level"), 0)
    return {
        "is_heading": is_heading,
        "is_heading_candidate": is_heading,
        "heading_score": float(decision.get("score", 0.0)),
        "heading_depth_hint": (level if is_heading and level > 0 else None),
        "reason": "Surya heading decision",
    }


def find_alias_anchor(
    obj: Dict[str, Any],
    structure_rows: Sequence[Dict[str, Any]],
) -> Optional[Dict[str, Any]]:
    obj_text = normalize_text(str(obj.get("text", "")))
    if not obj_text:
        return None

    best: Optional[Dict[str, Any]] = None
    best_score = 0.0
    obj_is_title = bool(obj.get("is_title_placeholder"))
    obj_text_len = len(canonical_text(obj_text))
    for row in structure_rows:
        if row.get("coord_source") == "xml-unmatched":
            continue
        row_text = normalize_text(str(row.get("text", "")))
        if not row_text:
            continue
        score = text_similarity(obj_text, row_text)
        if obj_is_title and bool(row.get("is_heading_candidate")):
            score += 0.15
        if obj_text_len <= 8 and len(canonical_text(row_text)) <= 8:
            score += 0.10
        if parse_int(obj.get("y"), 10**18) != 10**18 and parse_int(row.get("y"), 10**18) != 10**18:
            dy = abs(parse_int(obj.get("y"), 10**18) - parse_int(row.get("y"), 10**18))
            if dy <= 300000:
                score += 0.05
        if score > best_score:
            best_score = score
            best = row

    threshold = 0.92 if obj_text_len <= 6 else 0.82
    if obj_is_title:
        threshold = 0.65
    if best_score < threshold:
        return None
    return best


def bucket_for_object(obj: Dict[str, Any], heading_fields: Dict[str, Any]) -> int:
    if obj.get("is_footer"):
        return 4
    if obj.get("is_decorative"):
        return 5
    depth = heading_fields.get("heading_depth_hint")
    text = str(obj.get("text", ""))
    if isinstance(depth, int) and depth > 1 and re.match(r"^\d+(?:[.)]|\.\d+)", text):
        return 0
    if heading_fields.get("is_heading_candidate"):
        return 1
    return 2


def build_ordered_xml_indexes(
    xml_objects: Sequence[Dict[str, Any]],
    reading_order: Sequence[Dict[str, Any]],
) -> Tuple[List[int], List[Dict[str, Any]], Dict[str, int]]:
    by_shape_id = {str(obj.get("shape_id", "")): obj for obj in xml_objects}
    matched_ids: List[str] = []
    matched_rows: List[Dict[str, Any]] = []
    alias_rows_by_anchor: Dict[str, List[Dict[str, Any]]] = {}
    unmatched_rows: List[Dict[str, Any]] = []

    unique_matches = choose_unique_shape_matches(reading_order)
    for rank, row in enumerate(unique_matches):
        shape_id = str(row.get("matched_shape_id") or "").strip()
        obj = by_shape_id.get(shape_id)
        if obj is None:
            continue
        matched_ids.append(shape_id)
        heading_fields = heading_fields_from_surya(row)
        matched_rows.append(
            {
                "shape_id": obj["shape_id"],
                "xml_index": obj["xml_index"],
                "surya_rank": parse_int(row.get("position"), rank),
                "tag": obj["tag"],
                "name": obj["name"],
                "ph_type": obj["ph_type"],
                "ph_idx": obj["ph_idx"],
                "x": obj["x"],
                "y": obj["y"],
                "coord_source": "surya-normalized",
                "text": normalize_text(str(row.get("xml_text") or row.get("ocr_text") or obj.get("text", ""))),
                "bbox": row.get("bbox") if isinstance(row.get("bbox"), list) else obj.get("bbox"),
                "label": row.get("label"),
                "is_footer": bool(obj.get("is_footer")),
                "is_decorative": bool(obj.get("is_decorative")),
                "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                **heading_fields,
            }
        )

    unmatched_count = 0
    alias_count = 0
    for obj in xml_objects:
        shape_id = str(obj.get("shape_id", ""))
        if shape_id in matched_ids:
            continue
        alias_anchor = None
        if (
            not bool(obj.get("is_footer"))
            and not bool(obj.get("is_decorative"))
            and normalize_text(str(obj.get("text", "")))
        ):
            alias_anchor = find_alias_anchor(obj, matched_rows)
        if alias_anchor is not None:
            matched_ids.append(shape_id)
            alias_count += 1
            heading_fields = {
                "is_heading": bool(alias_anchor.get("is_heading", False)),
                "is_heading_candidate": bool(alias_anchor.get("is_heading_candidate", False)),
                "heading_score": float(alias_anchor.get("heading_score", 0.0)),
                "heading_depth_hint": alias_anchor.get("heading_depth_hint"),
                "reason": "Alias match to Surya-ordered text block",
            }
            alias_rows_by_anchor.setdefault(str(alias_anchor.get("shape_id", "")), []).append(
                {
                    "shape_id": obj["shape_id"],
                    "xml_index": obj["xml_index"],
                    "surya_rank": alias_anchor.get("surya_rank"),
                    "tag": obj["tag"],
                    "name": obj["name"],
                    "ph_type": obj["ph_type"],
                    "ph_idx": obj["ph_idx"],
                    "x": obj["x"],
                    "y": obj["y"],
                    "coord_source": "surya-alias",
                    "text": normalize_text(str(obj.get("text", ""))),
                    "bbox": obj.get("bbox"),
                    "label": alias_anchor.get("label"),
                    "is_footer": bool(obj.get("is_footer")),
                    "is_decorative": bool(obj.get("is_decorative")),
                    "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                    **heading_fields,
                }
            )
            continue
        unmatched_count += 1
        heading_fields = {
            "is_heading": False,
            "is_heading_candidate": False,
            "heading_score": 0.0,
            "heading_depth_hint": None,
            "reason": "Unmatched XML object appended after Surya-ordered shapes",
        }
        unmatched_rows.append(
            {
                "shape_id": obj["shape_id"],
                "xml_index": obj["xml_index"],
                "surya_rank": None,
                "tag": obj["tag"],
                "name": obj["name"],
                "ph_type": obj["ph_type"],
                "ph_idx": obj["ph_idx"],
                "x": obj["x"],
                "y": obj["y"],
                "coord_source": "xml-unmatched",
                "text": normalize_text(str(obj.get("text", ""))),
                "bbox": obj.get("bbox"),
                "label": None,
                "is_footer": bool(obj.get("is_footer")),
                "is_decorative": bool(obj.get("is_decorative")),
                "is_title_placeholder": bool(obj.get("is_title_placeholder")),
                **heading_fields,
            }
        )

    structure_rows: List[Dict[str, Any]] = []
    for row in matched_rows:
        structure_rows.append(row)
    alias_rows: List[Dict[str, Any]] = []
    for row in matched_rows:
        aliases = alias_rows_by_anchor.get(str(row.get("shape_id", "")), [])
        aliases.sort(
            key=lambda item: (
                parse_int(item.get("surya_rank"), 10**18),
                parse_int(item.get("y"), 10**18),
                parse_int(item.get("x"), 10**18),
                parse_int(item.get("xml_index"), 10**18),
            )
        )
        alias_rows.extend(aliases)
    structure_rows.extend(alias_rows)
    unmatched_rows.sort(key=lambda row: parse_int(row.get("xml_index"), 10**18))
    structure_rows.extend(unmatched_rows)

    for row in structure_rows:
        row["bucket"] = bucket_for_object(row, row)

    ordered_xml_indexes = [int(row["xml_index"]) for row in structure_rows]
    stats = {
        "matched_shapes": len(matched_ids),
        "unmatched_shapes": unmatched_count,
        "alias_matches": alias_count,
        "duplicate_shape_matches": max(0, len([r for r in reading_order if r.get("matched_shape_id")]) - len(matched_ids)),
    }
    return ordered_xml_indexes, structure_rows, stats


def reorder_tree_by_indexes(tree: ET.ElementTree, ordered_xml_indexes: Sequence[int]) -> None:
    root = tree.getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return

    children = list(sp_tree)
    reorderables = [ch for ch in children if local_name(ch.tag) in REORDERABLE]
    if not reorderables:
        return

    idx_to_elem = {i + 1: elem for i, elem in enumerate(reorderables)}
    reordered_elems = [idx_to_elem[i] for i in ordered_xml_indexes if i in idx_to_elem]
    used = set(ordered_xml_indexes)
    for i, elem in idx_to_elem.items():
        if i not in used:
            reordered_elems.append(elem)

    new_children: List[ET.Element] = []
    inserted = False
    for ch in children:
        if local_name(ch.tag) in REORDERABLE:
            if not inserted:
                new_children.extend(reordered_elems)
                inserted = True
            continue
        new_children.append(ch)
    sp_tree[:] = new_children


def confidence_label(stats: Dict[str, int], total: int) -> str:
    if total <= 0:
        return "low"
    ratio = stats["unmatched_shapes"] / total
    if ratio >= 0.3:
        return "low"
    if ratio >= 0.1:
        return "medium"
    return "high"


def build_structure_analysis_sidecar(
    slide_xml: Path,
    output_xml: Path,
    xml_objects: Sequence[Dict[str, Any]],
    structure_rows: Sequence[Dict[str, Any]],
    ordered_xml_indexes: Sequence[int],
    page_payload: Dict[str, Any],
    stats: Dict[str, int],
) -> Dict[str, Any]:
    tables = page_payload.get("tables", [])
    reading_order = page_payload.get("reading_order", [])
    total = len(xml_objects)
    headings = sum(1 for row in structure_rows if bool(row.get("is_heading_candidate")))
    return {
        "input_xml": str(slide_xml),
        "output_xml": str(output_xml),
        "mode": "surya",
        "reading_order_source": "surya_pipeline/normalized",
        "confidence": confidence_label(stats, total),
        "counts": {
            "total": total,
            "text": sum(1 for obj in xml_objects if normalize_text(str(obj.get("text", "")))),
            "graphicFrame": sum(1 for obj in xml_objects if obj.get("tag") == "graphicFrame"),
            "pic": sum(1 for obj in xml_objects if obj.get("tag") == "pic"),
            "footer": sum(1 for obj in xml_objects if obj.get("is_footer")),
            "decorative": sum(1 for obj in xml_objects if obj.get("is_decorative")),
            "surya_reading_blocks": len(reading_order) if isinstance(reading_order, list) else 0,
            "matched_shapes": stats["matched_shapes"],
            "unmatched_shapes": stats["unmatched_shapes"],
            "alias_matches": stats.get("alias_matches", 0),
            "duplicate_shape_matches": stats["duplicate_shape_matches"],
            "headings": headings,
            "tables": len(tables) if isinstance(tables, list) else 0,
        },
        "ordered_xml_indexes": list(ordered_xml_indexes),
        "structure_order": list(structure_rows),
        "raw_xml_order": [
            {
                "shape_id": obj["shape_id"],
                "xml_index": obj["xml_index"],
                "tag": obj["tag"],
                "name": obj["name"],
                "ph_type": obj["ph_type"],
                "ph_idx": obj["ph_idx"],
                "x": obj["x"],
                "y": obj["y"],
                "text": obj["text"],
                "bbox": obj["bbox"],
            }
            for obj in xml_objects
        ],
        "surya_page": {
            "page": page_payload.get("page"),
            "image_bbox": page_payload.get("image_bbox"),
        },
        "tables": tables if isinstance(tables, list) else [],
    }


def write_structure_ready_outputs(
    slide_xml: Path,
    page_payload: Dict[str, Any],
    output_dir: Path,
) -> Dict[str, Any]:
    output_dir.mkdir(parents=True, exist_ok=True)
    xml_objects = parse_slide_xml_objects(slide_xml)
    reading_order = page_payload.get("reading_order", [])
    if not isinstance(reading_order, list):
        reading_order = []
    ordered_xml_indexes, structure_rows, stats = build_ordered_xml_indexes(xml_objects, reading_order)

    tree = ET.parse(slide_xml)
    reorder_tree_by_indexes(tree, ordered_xml_indexes)

    stem = slide_xml.stem
    output_xml = output_dir / f"{stem}.reordered.xml"
    output_json = output_dir / f"{stem}.structure_analysis.json"
    sidecar = build_structure_analysis_sidecar(
        slide_xml=slide_xml,
        output_xml=output_xml,
        xml_objects=xml_objects,
        structure_rows=structure_rows,
        ordered_xml_indexes=ordered_xml_indexes,
        page_payload=page_payload,
        stats=stats,
    )
    tree.write(output_xml, encoding="utf-8", xml_declaration=True)
    rels_dst = copy_slide_relationship_bundle(slide_xml, output_dir)
    if rels_dst is not None:
        sidecar["output_rels"] = str(rels_dst)
    output_json.write_text(json.dumps(sidecar, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "page": page_payload.get("page"),
        "input_xml": str(slide_xml),
        "output_xml": str(output_xml),
        "output_json": str(output_json),
        "output_rels": str(rels_dst) if rels_dst is not None else None,
        "matched_shapes": stats["matched_shapes"],
        "unmatched_shapes": stats["unmatched_shapes"],
        "confidence": sidecar["confidence"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Build reordered XML + sidecars from normalized Surya output.")
    parser.add_argument("--normalized-json", required=True, help="Path to normalized Surya output JSON.")
    parser.add_argument("--ppt-root", default=None, help="Extracted PPT root directory containing ppt/slides.")
    parser.add_argument("--pptx-path", default=None, help="Optional source PPTX path. Used if --ppt-root is absent.")
    parser.add_argument("--output-dir", required=True, help="Output directory for reordered XML and sidecars.")
    args = parser.parse_args()

    normalized_json = Path(args.normalized_json).resolve()
    output_dir = Path(args.output_dir).resolve()
    temp_ppt_dir: Optional[Path] = None
    if args.ppt_root:
        ppt_root = Path(args.ppt_root).resolve()
    elif args.pptx_path:
        temp_ppt_dir, ppt_root = extract_pptx_to_temp_root(Path(args.pptx_path).resolve())
    else:
        raise ValueError("either --ppt-root or --pptx-path is required")

    pages = load_normalized_pages(normalized_json)
    manifest: Dict[str, Any] = {
        "mode": "surya",
        "source": {
            "normalized_json": str(normalized_json),
            "ppt_root": str(ppt_root),
            "pptx_path": str(Path(args.pptx_path).resolve()) if args.pptx_path else None,
        },
        "processed": [],
        "failed": [],
    }

    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])

    for page_no in sorted(pages.keys()):
        slide_xml = ppt_root / "ppt" / "slides" / f"slide{page_no}.xml"
        if not slide_xml.exists():
            manifest["failed"].append({"page": page_no, "error": f"slide xml not found: {slide_xml}"})
            continue
        try:
            row = write_structure_ready_outputs(slide_xml, pages[page_no], output_dir)
            manifest["processed"].append(row)
            print(f"Processed: slide{page_no}.xml")
        except Exception as exc:  # noqa: BLE001
            manifest["failed"].append({"page": page_no, "input_xml": str(slide_xml), "error": str(exc)})
            print(f"Failed: slide{page_no}.xml -> {exc}")

    manifest_path = output_dir / "structure_analysis_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote manifest: {manifest_path}")

    if temp_ppt_dir is not None and temp_ppt_dir.exists():
        shutil.rmtree(temp_ppt_dir, ignore_errors=True)
    return 1 if manifest["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())

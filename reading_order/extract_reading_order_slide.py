#!/usr/bin/env python3
"""
Extract reading order per slide XML and write:
1) reading-order JSON
2) reordered slide XML

Usage:
  python extract_reading_order_slide.py [slide1.xml slide2.xml ...]

If no positional args are given, all *.xml in ./target_slides are processed.
Outputs are always written to ./output (created automatically if missing).
"""

from __future__ import annotations

import argparse
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

REL_NS = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}

SLIDE_LAYOUT_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
FOOTER_TYPES = {"sldNum", "ftr", "dt"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}


@dataclass
class SlideObject:
    shape_id: str
    xml_index: int
    tag: str
    name: str
    ph_type: Optional[str]
    ph_idx: Optional[str]
    x: int
    y: int
    coord_source: str
    text: str
    normalized: str
    is_footer: bool
    is_decorative: bool
    is_heading: bool
    is_title_placeholder: bool
    is_duplicate_title: bool = False
    surya_rank: Optional[int] = None


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def natural_key(path: Path) -> Tuple:
    parts = re.split(r"(\d+)", path.name)
    out: List[object] = []
    for part in parts:
        if part.isdigit():
            out.append(int(part))
        else:
            out.append(part.lower())
    return tuple(out)


def parse_int(v: Optional[str], default: int = 10**18) -> int:
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def normalize_text(s: str) -> str:
    s = s.lower().replace("\n", " ")
    s = re.sub(r"[^0-9a-z\uac00-\ud7a3\.\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def token_set(s: str) -> set:
    return set(re.findall(r"[0-9a-z\uac00-\ud7a3]+", normalize_text(s)))


def is_numbered_heading_text(text: str) -> bool:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    t = normalize_text(text)
    # 1. / 1.1 / 1.1.1 / 1)
    if re.match(r"^(?:\d+\.\d+(?:\.\d+)*\.?|\d+\.)\s+", t):
        return True
    if re.match(r"^\d+\)\s+", raw):
        return True
    return False


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


def resolve_slide_layout(slide_xml: Path) -> Optional[Path]:
    rels = slide_xml.parent / "_rels" / f"{slide_xml.name}.rels"
    if not rels.exists():
        return None

    rel_root = ET.parse(rels).getroot()
    for rel in rel_root.findall("rel:Relationship", REL_NS):
        if rel.attrib.get("Type") != SLIDE_LAYOUT_REL_TYPE:
            continue
        target = rel.attrib.get("Target")
        if not target:
            continue
        resolved = Path(os.path.normpath(str(slide_xml.parent / target)))
        if resolved.exists():
            return resolved
    return None


def parse_layout_placeholders(layout_xml: Optional[Path]) -> Dict[Tuple[str, str], Tuple[int, int]]:
    out: Dict[Tuple[str, str], Tuple[int, int]] = {}
    if layout_xml is None or not layout_xml.exists():
        return out

    root = ET.parse(layout_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return out

    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue

        _, ph_path = get_nvpr_paths(tag)
        ph = ch.find(ph_path, NS)
        if ph is None:
            continue
        ph_type = ph.attrib.get("type", "body")
        ph_idx = ph.attrib.get("idx", "")

        off = first_off(ch)
        if off is None:
            continue
        x = parse_int(off.attrib.get("x"))
        y = parse_int(off.attrib.get("y"))

        key = (ph_type, ph_idx)
        if key not in out:
            out[key] = (x, y)
        key_type = (ph_type, "")
        if key_type not in out:
            out[key_type] = (x, y)

    return out


def looks_heading(text: str, ph_type: Optional[str], tag: str) -> bool:
    raw = (text or "").strip()
    if raw.startswith(("▶", "-", "*", "√")):
        return False

    t = normalize_text(text)
    if not t:
        return False
    if is_numbered_heading_text(text):
        return True
    if ph_type in TITLE_TYPES and len(t) <= 80:
        return True
    if tag == "sp" and len(t) <= 80 and any(
        k in t for k in ("평가 결과", "조사 결과", "summary", "개발 계획")
    ):
        return True
    return False


def is_decorative(tag: str, text: str) -> bool:
    t = normalize_text(text)
    if tag == "cxnSp":
        return True
    if not t and tag not in {"pic", "graphicFrame"}:
        return True
    return False


def parse_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def parse_bbox(raw: Any) -> Optional[Tuple[float, float, float, float]]:
    if isinstance(raw, dict):
        keys = ("x1", "y1", "x2", "y2")
        if all(k in raw for k in keys):
            x1 = parse_float(raw.get("x1"))
            y1 = parse_float(raw.get("y1"))
            x2 = parse_float(raw.get("x2"))
            y2 = parse_float(raw.get("y2"))
            return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
        return None

    if not isinstance(raw, list):
        return None
    if len(raw) == 4 and all(isinstance(x, (int, float)) for x in raw):
        x1, y1, x2, y2 = (float(raw[0]), float(raw[1]), float(raw[2]), float(raw[3]))
        return (min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2))
    if len(raw) == 4 and all(isinstance(x, list) and len(x) >= 2 for x in raw):
        xs = [parse_float(p[0]) for p in raw]
        ys = [parse_float(p[1]) for p in raw]
        return (min(xs), min(ys), max(xs), max(ys))
    return None


def bbox_center_xy(x1: float, y1: float, x2: float, y2: float) -> Tuple[float, float]:
    return ((x1 + x2) / 2.0, (y1 + y2) / 2.0)


def normalize_xy(x: float, y: float, max_x: float, max_y: float) -> Tuple[float, float]:
    nx = (x / max_x) if max_x > 0 else 0.0
    ny = (y / max_y) if max_y > 0 else 0.0
    return (nx, ny)


def find_surya_json(slide_xml: Path, surya_dir: Path) -> Optional[Path]:
    stem = slide_xml.stem
    candidates = [
        surya_dir / f"{stem}.json",
        surya_dir / f"{slide_xml.name}.json",
    ]
    for c in candidates:
        if c.exists() and c.is_file():
            return c

    wildcard_hits = sorted(surya_dir.glob(f"*{stem}*.json"), key=lambda p: p.name.lower())
    for c in wildcard_hits:
        if c.is_file():
            return c
    return None


def _looks_surya_block(node: Dict[str, Any]) -> bool:
    if parse_bbox(node.get("bbox")) is not None:
        return True
    if parse_bbox(node.get("polygon")) is not None:
        return True
    return False


def _collect_surya_blocks(node: Any, out: List[Dict[str, Any]]) -> None:
    if isinstance(node, dict):
        if _looks_surya_block(node):
            out.append(node)
        for v in node.values():
            _collect_surya_blocks(v, out)
        return
    if isinstance(node, list):
        for item in node:
            _collect_surya_blocks(item, out)


def load_surya_blocks(slide_xml: Path, surya_dir: Path) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    surya_json = find_surya_json(slide_xml, surya_dir)
    if surya_json is None:
        raise FileNotFoundError(f"surya result not found for {slide_xml.name} under {surya_dir}")

    payload = json.loads(surya_json.read_text(encoding="utf-8"))
    raw_blocks: List[Dict[str, Any]] = []
    _collect_surya_blocks(payload, raw_blocks)

    parsed: List[Dict[str, Any]] = []
    for row in raw_blocks:
        bbox = parse_bbox(row.get("bbox"))
        if bbox is None:
            bbox = parse_bbox(row.get("polygon"))
        if bbox is None:
            continue

        label_raw = row.get("label") or row.get("type") or row.get("class") or row.get("category")
        text_raw = row.get("text") or row.get("value") or row.get("content") or ""
        position = row.get("position")

        label = str(label_raw).strip() if label_raw is not None else ""
        text = str(text_raw).strip() if text_raw is not None else ""

        x1, y1, x2, y2 = bbox
        cx, cy = bbox_center_xy(x1, y1, x2, y2)
        parsed.append(
            {
                "bbox": [x1, y1, x2, y2],
                "cx": cx,
                "cy": cy,
                "label": label,
                "text": text,
                "position": parse_int(str(position), default=10**18) if position is not None else 10**18,
            }
        )

    parsed.sort(key=lambda b: (b["position"], b["cy"], b["cx"]))
    for i, b in enumerate(parsed):
        if b["position"] >= 10**18:
            b["position"] = i

    return parsed, {"surya_json": str(surya_json), "surya_block_count": len(parsed)}


def heading_depth_from_text_label(text: str, label: str) -> Optional[int]:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    low_label = (label or "").lower()
    if not raw and "header" not in low_label and "title" not in low_label:
        return None

    if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", raw):
        return 3
    if re.match(r"^\d+[.)]\s+", raw):
        return 2
    if any(k in low_label for k in ("section-header", "header", "title")):
        return 1
    if len(raw) <= 80:
        return 1
    return None


def apply_surya_ordering(objects: List[SlideObject], blocks: List[Dict[str, Any]]) -> List[SlideObject]:
    if not objects:
        return objects
    if not blocks:
        return objects

    finite_x = [float(o.x) for o in objects if o.x < 10**18]
    finite_y = [float(o.y) for o in objects if o.y < 10**18]
    max_xml_x = max(finite_x) if finite_x else 1.0
    max_xml_y = max(finite_y) if finite_y else 1.0

    max_surya_x = max(float(b["bbox"][2]) for b in blocks) if blocks else 1.0
    max_surya_y = max(float(b["bbox"][3]) for b in blocks) if blocks else 1.0

    assigned: Dict[int, Dict[str, Any]] = {}
    for blk in blocks:
        bx, by = normalize_xy(float(blk["cx"]), float(blk["cy"]), max_surya_x, max_surya_y)

        best_i = -1
        best_dist = float("inf")
        for i, obj in enumerate(objects):
            ox, oy = normalize_xy(float(obj.x if obj.x < 10**18 else max_xml_x), float(obj.y if obj.y < 10**18 else max_xml_y), max_xml_x, max_xml_y)
            d = (ox - bx) ** 2 + (oy - by) ** 2
            if d < best_dist:
                best_dist = d
                best_i = i
        if best_i < 0:
            continue

        prev = assigned.get(best_i)
        if prev is None or int(blk["position"]) < int(prev["position"]):
            assigned[best_i] = blk

    out: List[SlideObject] = []
    for i, obj in enumerate(objects):
        blk = assigned.get(i)
        text = blk["text"] if blk and blk.get("text") else obj.text
        label = blk["label"] if blk else ""
        depth = heading_depth_from_text_label(text, label)
        is_heading = depth is not None
        out.append(
            SlideObject(
                shape_id=obj.shape_id,
                xml_index=obj.xml_index,
                tag=obj.tag,
                name=obj.name,
                ph_type=obj.ph_type,
                ph_idx=obj.ph_idx,
                x=obj.x,
                y=obj.y,
                coord_source="surya",
                text=text,
                normalized=normalize_text(text),
                is_footer=obj.is_footer,
                is_decorative=obj.is_decorative,
                is_heading=is_heading,
                is_title_placeholder=obj.is_title_placeholder,
                is_duplicate_title=False,
                surya_rank=(int(blk["position"]) if blk else None),
            )
        )
    detect_duplicate_titles(out)
    return out


def detect_duplicate_titles(objs: List[SlideObject]) -> None:
    numbered = [o for o in objs if o.is_heading and is_numbered_heading_text(o.text)]
    if not numbered:
        return

    for obj in objs:
        if not obj.is_title_placeholder:
            continue
        ts = token_set(obj.text)
        if not ts:
            continue
        for heading in numbered:
            hs = token_set(heading.text)
            if not hs:
                continue
            inter = len(ts & hs)
            union = len(ts | hs)
            jaccard = (inter / union) if union else 0.0
            if jaccard >= 0.55:
                obj.is_duplicate_title = True
                break


def bucket(obj: SlideObject) -> int:
    if obj.is_footer:
        return 4
    if obj.is_decorative:
        return 5
    if obj.is_heading and is_numbered_heading_text(obj.text):
        return 0
    if obj.is_heading and not obj.is_duplicate_title:
        return 1
    if obj.is_duplicate_title:
        return 3
    return 2


def reason(obj: SlideObject) -> str:
    if obj.is_footer:
        return "Footer or slide number placeholder"
    if obj.is_decorative:
        return "Decorative connector/shape, low reading priority"
    if obj.is_heading and is_numbered_heading_text(obj.text):
        return "Numbered section heading pattern"
    if obj.is_heading and not obj.is_duplicate_title:
        if obj.coord_source == "layout":
            return "Title placeholder with layout-inherited coordinates"
        return "Short title-like text block"
    if obj.is_duplicate_title:
        return "Duplicate of a section heading title placeholder"
    if obj.tag == "graphicFrame":
        return "Table/chart frame as a body block"
    if obj.tag == "pic":
        return "Image block (no text)"
    return "General body text block (y->x)"


def heading_depth_hint(obj: SlideObject) -> Optional[int]:
    # Final markdown depth hint should be emitted by converter.
    # Here we provide structural hints only.
    if obj.is_duplicate_title:
        return None
    if obj.is_heading and is_numbered_heading_text(obj.text):
        raw = re.sub(r"\s+", " ", (obj.text or "").strip())
        t = normalize_text(obj.text)
        if re.match(r"^\d+\)\s+", raw):
            return 2
        if re.match(r"^\d+\.\d+(?:\.\d+)*\.?\s+", t):
            return 3
        if re.match(r"^\d+\.\s+", t):
            return 2
    if obj.is_heading and not obj.is_duplicate_title:
        return 1
    return None


def heading_score(obj: SlideObject) -> float:
    # Confidence-like score for heading candidacy.
    if obj.is_duplicate_title:
        return 0.2
    if obj.is_heading and is_numbered_heading_text(obj.text):
        return 0.95
    if obj.is_heading and not obj.is_duplicate_title:
        return 0.75
    return 0.0


def confidence(objs: Sequence[SlideObject]) -> str:
    if not objs:
        return "low"
    unknown = sum(1 for o in objs if o.coord_source == "unknown")
    ratio = unknown / len(objs)
    if ratio >= 0.3:
        return "low"
    if ratio >= 0.1:
        return "medium"
    return "high"


def object_to_dict(o: SlideObject) -> Dict[str, object]:
    depth = heading_depth_hint(o)
    score = heading_score(o)
    is_candidate = depth is not None and score >= 0.7
    return {
        "shape_id": o.shape_id,
        "xml_index": o.xml_index,
        "tag": o.tag,
        "name": o.name,
        "ph_type": o.ph_type,
        "ph_idx": o.ph_idx,
        "x": o.x,
        "y": o.y,
        "coord_source": o.coord_source,
        "text": o.text,
        "is_footer": o.is_footer,
        "is_decorative": o.is_decorative,
        "is_heading": o.is_heading,
        "is_title_placeholder": o.is_title_placeholder,
        "is_duplicate_title": o.is_duplicate_title,
        "is_heading_candidate": is_candidate,
        "heading_score": round(score, 3),
        "heading_depth_hint": depth,
        "surya_rank": o.surya_rank,
        "bucket": bucket(o),
        "reason": reason(o),
    }


def extract_slide_objects_xml(slide_xml: Path) -> Tuple[List[SlideObject], Dict[str, object]]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return [], {"error": "Missing p:cSld/p:spTree"}

    layout_xml = resolve_slide_layout(slide_xml)
    layout_map = parse_layout_placeholders(layout_xml)

    objects: List[SlideObject] = []
    xml_idx = 0
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        xml_idx += 1

        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = ch.find(c_nv_path, NS)
        ph = ch.find(ph_path, NS)

        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        name = c_nv_pr.attrib.get("name", "") if c_nv_pr is not None else ""
        ph_type = ph.attrib.get("type") if ph is not None else None
        ph_idx = ph.attrib.get("idx") if ph is not None else None

        off = first_off(ch)
        coord_source = "direct"
        if off is not None:
            x = parse_int(off.attrib.get("x"))
            y = parse_int(off.attrib.get("y"))
        else:
            x = 10**18
            y = 10**18
            if ph_type is not None:
                key = (ph_type, ph_idx or "")
                if key in layout_map:
                    x, y = layout_map[key]
                    coord_source = "layout"
                elif (ph_type, "") in layout_map:
                    x, y = layout_map[(ph_type, "")]
                    coord_source = "layout"
                else:
                    coord_source = "unknown"
            else:
                coord_source = "unknown"

        texts = []
        for t in ch.findall(".//a:t", NS):
            if t.text and t.text.strip():
                texts.append(t.text.strip())
        text = " ".join(texts)
        normalized = normalize_text(text)

        objects.append(
            SlideObject(
                shape_id=shape_id,
                xml_index=xml_idx,
                tag=tag,
                name=name,
                ph_type=ph_type,
                ph_idx=ph_idx,
                x=x,
                y=y,
                coord_source=coord_source,
                text=text,
                normalized=normalized,
                is_footer=(ph_type in FOOTER_TYPES) or bool(re.fullmatch(r"\d+", normalized)),
                is_decorative=is_decorative(tag, text),
                is_heading=looks_heading(text, ph_type, tag),
                is_title_placeholder=ph_type in TITLE_TYPES,
            )
        )

    detect_duplicate_titles(objects)
    meta = {
        "layout_xml": str(layout_xml) if layout_xml else None,
        "layout_placeholder_count": len(layout_map),
    }
    return objects, meta


def extract_slide_objects(
    slide_xml: Path,
    mode: str,
    surya_dir: Optional[Path],
) -> Tuple[List[SlideObject], Dict[str, object]]:
    objects, meta = extract_slide_objects_xml(slide_xml)
    if mode == "xml":
        meta["mode"] = "xml"
        return objects, meta

    if surya_dir is None:
        raise ValueError("surya mode requires --surya-dir")
    blocks, surya_meta = load_surya_blocks(slide_xml, surya_dir)
    if not blocks:
        raise ValueError(f"surya mode has no usable blocks for {slide_xml.name}")

    mapped = apply_surya_ordering(objects, blocks)
    meta.update(surya_meta)
    meta["mode"] = "surya"
    return mapped, meta


def order_objects(objects: Sequence[SlideObject]) -> List[SlideObject]:
    return sorted(
        objects,
        key=lambda o: (
            0 if o.surya_rank is not None else 1,
            o.surya_rank if o.surya_rank is not None else 10**18,
            bucket(o),
            o.y,
            o.x,
            o.xml_index,
        ),
    )


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


def gather_input_files(target_dir: Path, raw_inputs: Sequence[str]) -> List[Path]:
    files: List[Path] = []

    if raw_inputs:
        for item in raw_inputs:
            p = Path(item)
            candidates = [p, target_dir / item]
            if p.suffix == "":
                candidates.append(target_dir / f"{item}.xml")
            selected: Optional[Path] = None
            for c in candidates:
                if c.exists() and c.is_file():
                    selected = c
                    break
            if selected is None:
                continue
            files.append(selected.resolve())
    else:
        if not target_dir.exists():
            return []
        files = sorted(target_dir.glob("*.xml"), key=natural_key)
        files = [f.resolve() for f in files if f.is_file()]

    unique: Dict[str, Path] = {}
    for f in files:
        unique[str(f)] = f
    return list(unique.values())


def write_outputs(
    slide_xml: Path,
    output_dir: Path,
    mode: str,
    surya_dir: Optional[Path],
) -> Dict[str, object]:
    objects, meta = extract_slide_objects(slide_xml, mode=mode, surya_dir=surya_dir)
    ordered = order_objects(objects)

    ordered_indexes = [obj.xml_index for obj in ordered]

    tree = ET.parse(slide_xml)
    reorder_tree_by_indexes(tree, ordered_indexes)

    stem = slide_xml.stem
    json_path = output_dir / f"{stem}.reading_order.json"
    xml_path = output_dir / f"{stem}.reordered.xml"

    report: Dict[str, object] = {
        "input_xml": str(slide_xml),
        "mode": mode,
        "layout_xml": meta.get("layout_xml"),
        "surya_json": meta.get("surya_json"),
        "confidence": confidence(objects),
        "counts": {
            "total": len(objects),
            "text": sum(1 for o in objects if bool(o.normalized)),
            "graphicFrame": sum(1 for o in objects if o.tag == "graphicFrame"),
            "pic": sum(1 for o in objects if o.tag == "pic"),
            "footer": sum(1 for o in objects if o.is_footer),
            "decorative": sum(1 for o in objects if o.is_decorative),
            "layout_coord_used": sum(1 for o in objects if o.coord_source == "layout"),
            "surya_ranked": sum(1 for o in objects if o.surya_rank is not None),
            "surya_blocks": int(meta.get("surya_block_count", 0)),
        },
        "reading_order": [object_to_dict(o) for o in ordered],
        "raw_xml_order": [object_to_dict(o) for o in sorted(objects, key=lambda x: x.xml_index)],
        "ordered_xml_indexes": ordered_indexes,
        "output_reordered_xml": str(xml_path),
    }

    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)

    return {
        "input_xml": str(slide_xml),
        "output_json": str(json_path),
        "output_xml": str(xml_path),
        "confidence": report["confidence"],
        "object_count": len(objects),
    }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Extract reading order from slide XML and write JSON + reordered XML."
    )
    parser.add_argument(
        "inputs",
        nargs="*",
        help="Input slide xml file(s). If omitted, process all *.xml from ./target_slides.",
    )
    parser.add_argument(
        "--mode",
        choices=("xml", "surya"),
        default="xml",
        help="Reading order mode. xml=legacy XML heuristics, surya=Surya OCR/layout based.",
    )
    parser.add_argument(
        "--surya-dir",
        default=None,
        help="Directory containing Surya per-slide JSON outputs. Required in --mode surya.",
    )
    parser.add_argument(
        "--output-dir",
        default="./output",
        help="Output directory for reading-order JSON/XML artifacts.",
    )
    args = parser.parse_args()

    target_dir = Path("./target_slides")
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    surya_dir = Path(args.surya_dir).resolve() if args.surya_dir else None

    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])

    inputs = gather_input_files(target_dir, args.inputs)
    if not inputs:
        print("No input XML files found.")
        print(f"Checked default directory: {target_dir.resolve()}")
        return

    manifest = {"processed": [], "failed": []}

    for slide_xml in sorted(inputs, key=natural_key):
        try:
            row = write_outputs(slide_xml, output_dir, mode=args.mode, surya_dir=surya_dir)
            manifest["processed"].append(row)
            print(f"Processed: {slide_xml.name}")
        except Exception as e:  # noqa: BLE001
            manifest["failed"].append({"input_xml": str(slide_xml), "error": str(e)})
            print(f"Failed: {slide_xml.name} -> {e}")

    manifest_path = output_dir / "reading_order_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote manifest: {manifest_path.resolve()}")
    print(f"Output directory: {output_dir.resolve()}")
    print(f"Processed: {len(manifest['processed'])}, Failed: {len(manifest['failed'])}")


if __name__ == "__main__":
    main()

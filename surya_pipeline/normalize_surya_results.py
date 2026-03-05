#!/usr/bin/env python3
"""
Normalize Surya outputs into per-page reading order, headings, and table metadata.

Heading decision uses multi-signal post-processing:
- Surya layout label/confidence
- position on slide
- text length and numbered pattern
- PPTX XML placeholder type and font size
- OCR text (if OCR json is provided)
"""

from __future__ import annotations

import argparse
from collections import defaultdict
from difflib import SequenceMatcher
import json
import re
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple
import xml.etree.ElementTree as ET
import zipfile


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}

REORDERABLE = {"sp", "pic", "graphicFrame", "grpSp", "cxnSp"}
TITLE_TYPES = {"title", "ctrTitle", "subTitle"}


@dataclass
class XmlObject:
    shape_id: str
    tag: str
    ph_type: Optional[str]
    x: float
    y: float
    w: float
    h: float
    cx: float
    cy: float
    text: str
    font_pt: Optional[float]


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def pick_doc_payload(payload: Any, doc_key: Optional[str]) -> Tuple[str, Any]:
    if isinstance(payload, list):
        return "default", payload
    if not isinstance(payload, dict):
        raise ValueError("invalid json payload type")
    if doc_key:
        if doc_key not in payload:
            raise KeyError(f"doc key not found: {doc_key}")
        return doc_key, payload[doc_key]
    keys = list(payload.keys())
    if not keys:
        raise ValueError("empty payload dictionary")
    return keys[0], payload[keys[0]]


def pick_table_payload(
    payload: Any,
    doc_key: Optional[str],
    fallback_doc_key: str,
) -> Tuple[str, List[Dict[str, Any]]]:
    if isinstance(payload, list):
        rows = [x for x in payload if isinstance(x, dict)]
        return "default", rows
    if not isinstance(payload, dict):
        return fallback_doc_key, []

    if doc_key:
        node = payload.get(doc_key, [])
        if isinstance(node, list):
            return doc_key, [x for x in node if isinstance(x, dict)]
        return doc_key, []

    if fallback_doc_key in payload:
        node = payload.get(fallback_doc_key, [])
        if isinstance(node, list):
            return fallback_doc_key, [x for x in node if isinstance(x, dict)]
        return fallback_doc_key, []

    if not payload:
        return fallback_doc_key, []

    first_key = next(iter(payload.keys()))
    node = payload.get(first_key, [])
    if isinstance(node, list):
        return first_key, [x for x in node if isinstance(x, dict)]
    return first_key, []


def safe_float(v: Any, default: float = 0.0) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def safe_int(v: Any, default: int = 0) -> int:
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def norm_bbox(raw: Any) -> Optional[List[float]]:
    if not isinstance(raw, list) or len(raw) != 4:
        return None
    x1 = safe_float(raw[0])
    y1 = safe_float(raw[1])
    x2 = safe_float(raw[2])
    y2 = safe_float(raw[3])
    return [min(x1, x2), min(y1, y2), max(x1, x2), max(y1, y2)]


def merge_bboxes(bboxes: Sequence[List[float]]) -> Optional[List[float]]:
    if not bboxes:
        return None
    xs1 = [b[0] for b in bboxes]
    ys1 = [b[1] for b in bboxes]
    xs2 = [b[2] for b in bboxes]
    ys2 = [b[3] for b in bboxes]
    return [min(xs1), min(ys1), max(xs2), max(ys2)]


def bbox_area(bbox: Optional[List[float]]) -> float:
    if bbox is None:
        return 0.0
    return max(0.0, bbox[2] - bbox[0]) * max(0.0, bbox[3] - bbox[1])


def intersection_bbox(a: Optional[List[float]], b: Optional[List[float]]) -> Optional[List[float]]:
    if a is None or b is None:
        return None
    x1 = max(a[0], b[0])
    y1 = max(a[1], b[1])
    x2 = min(a[2], b[2])
    y2 = min(a[3], b[3])
    if x2 <= x1 or y2 <= y1:
        return None
    return [x1, y1, x2, y2]


def overlap_ratio(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    inter = intersection_bbox(a, b)
    if inter is None:
        return 0.0
    inter_area = bbox_area(inter)
    denom = min(bbox_area(a), bbox_area(b))
    if denom <= 0:
        return 0.0
    return inter_area / denom


def center_distance_score(a: Optional[List[float]], b: Optional[List[float]]) -> float:
    if a is None or b is None:
        return 0.0
    acx = (a[0] + a[2]) / 2.0
    acy = (a[1] + a[3]) / 2.0
    bcx = (b[0] + b[2]) / 2.0
    bcy = (b[1] + b[3]) / 2.0
    dist = ((acx - bcx) ** 2 + (acy - bcy) ** 2) ** 0.5
    return max(0.0, 1.0 - min(dist / 1.5, 1.0))


def text_similarity(a: str, b: str) -> float:
    left = normalize_text(a)
    right = normalize_text(b)
    if not left or not right:
        return 0.0
    if left in right or right in left:
        return 1.0
    return SequenceMatcher(None, left, right).ratio()


def table_bbox_from_item(item: Dict[str, Any]) -> Optional[List[float]]:
    for key in ("cells", "rows", "cols"):
        rows = item.get(key, [])
        if not isinstance(rows, list) or not rows:
            continue
        bboxes: List[List[float]] = []
        for row in rows:
            if not isinstance(row, dict):
                continue
            b = norm_bbox(row.get("bbox"))
            if b is not None:
                bboxes.append(b)
        merged = merge_bboxes(bboxes)
        if merged is not None:
            return merged
    return None


def local_name(tag: str) -> str:
    return tag.split("}", 1)[-1]


def parse_numbered_heading(text: str) -> Optional[int]:
    raw = re.sub(r"\s+", " ", (text or "").strip())
    if not raw:
        return None
    if re.match(r"^\d+\.\d+\.\d+\.?\s+", raw):
        return 3
    if re.match(r"^\d+\.\d+\.?\s+", raw):
        return 3
    if re.match(r"^\d+[.)]\s+", raw):
        return 2
    return None


def normalize_text(s: str) -> str:
    s = re.sub(r"\s+", " ", (s or "").strip())
    return s


def first_off_and_ext(elem: ET.Element) -> Tuple[Optional[ET.Element], Optional[ET.Element]]:
    off = None
    ext = None
    for p in (
        "./p:spPr/a:xfrm",
        "./p:grpSpPr/a:xfrm",
        "./p:xfrm",
        ".//a:xfrm",
    ):
        xfrm = elem.find(p, NS)
        if xfrm is None:
            continue
        off = xfrm.find("a:off", NS)
        ext = xfrm.find("a:ext", NS)
        if off is not None or ext is not None:
            return off, ext
    return off, ext


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


def extract_shape_text(elem: ET.Element) -> str:
    parts: List[str] = []
    for t in elem.findall(".//a:t", NS):
        if t.text and t.text.strip():
            parts.append(t.text.strip())
    return normalize_text(" ".join(parts))


def extract_font_pt(elem: ET.Element) -> Optional[float]:
    sizes: List[float] = []
    for rpr in elem.findall(".//a:rPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = safe_float(sz, -1.0)
        if val > 0:
            sizes.append(val / 100.0)
    for rpr in elem.findall(".//a:endParaRPr", NS):
        sz = rpr.attrib.get("sz")
        if sz is None:
            continue
        val = safe_float(sz, -1.0)
        if val > 0:
            sizes.append(val / 100.0)
    if not sizes:
        return None
    return max(sizes)


def parse_slide_xml_objects(slide_xml: Path) -> List[XmlObject]:
    root = ET.parse(slide_xml).getroot()
    sp_tree = root.find("p:cSld/p:spTree", NS)
    if sp_tree is None:
        return []

    out: List[XmlObject] = []
    for ch in list(sp_tree):
        tag = local_name(ch.tag)
        if tag not in REORDERABLE:
            continue
        c_nv_path, ph_path = get_nvpr_paths(tag)
        c_nv_pr = ch.find(c_nv_path, NS)
        ph = ch.find(ph_path, NS)
        shape_id = c_nv_pr.attrib.get("id", "") if c_nv_pr is not None else ""
        ph_type = ph.attrib.get("type") if ph is not None else None

        off, ext = first_off_and_ext(ch)
        x = safe_float(off.attrib.get("x")) if off is not None else 0.0
        y = safe_float(off.attrib.get("y")) if off is not None else 0.0
        w = safe_float(ext.attrib.get("cx")) if ext is not None else 0.0
        h = safe_float(ext.attrib.get("cy")) if ext is not None else 0.0
        cx = x + (w / 2.0)
        cy = y + (h / 2.0)

        out.append(
            XmlObject(
                shape_id=shape_id,
                tag=tag,
                ph_type=ph_type,
                x=x,
                y=y,
                w=w,
                h=h,
                cx=cx,
                cy=cy,
                text=extract_shape_text(ch),
                font_pt=extract_font_pt(ch),
            )
        )
    return out


def parse_slide_size_emu(ppt_root: Path) -> Tuple[float, float]:
    pres = ppt_root / "ppt" / "presentation.xml"
    if not pres.exists():
        return (1.0, 1.0)
    root = ET.parse(pres).getroot()
    sld_sz = root.find("p:sldSz", NS)
    if sld_sz is None:
        return (1.0, 1.0)
    cx = safe_float(sld_sz.attrib.get("cx"), 1.0)
    cy = safe_float(sld_sz.attrib.get("cy"), 1.0)
    return (max(cx, 1.0), max(cy, 1.0))


def xml_bbox_to_normalized(
    obj: XmlObject,
    slide_w: float,
    slide_h: float,
) -> Optional[List[float]]:
    if obj.w <= 0 or obj.h <= 0 or slide_w <= 0 or slide_h <= 0:
        return None
    return [
        obj.x / slide_w,
        obj.y / slide_h,
        (obj.x + obj.w) / slide_w,
        (obj.y + obj.h) / slide_h,
    ]


def extract_pptx_to_temp_root(pptx_path: Path) -> Tuple[Path, Path]:
    if not pptx_path.exists() or not pptx_path.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx_path}")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_xml_"))
    extracted_root = temp_dir / pptx_path.stem
    extracted_root.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pptx_path, "r") as zf:
        zf.extractall(extracted_root)
    return temp_dir, extracted_root


def resolve_results_json(raw: Optional[str], default_dir: Path, kind: str) -> Path:
    if raw:
        p = Path(raw).resolve()
        if p.is_file():
            return p
        if p.is_dir():
            hits = sorted(p.glob("*/results.json"))
            if hits:
                return hits[0].resolve()
        raise FileNotFoundError(f"{kind} results not found from: {p}")

    hits = sorted(default_dir.glob("*/results.json"))
    if hits:
        return hits[0].resolve()
    raise FileNotFoundError(f"{kind} results not found under default dir: {default_dir}")


def find_ocr_text_for_block(
    block_bbox: List[float],
    ocr_items: Sequence[Dict[str, Any]],
) -> str:
    rows = collect_ocr_rows_for_block(block_bbox, ocr_items)
    if rows:
        return join_ocr_text(rows)
    x1, y1, x2, y2 = block_bbox
    bx = (x1 + x2) / 2.0
    by = (y1 + y2) / 2.0
    best = ""
    best_dist = float("inf")
    for row in ocr_items:
        b = norm_bbox(row.get("bbox"))
        if b is None:
            b = norm_bbox(row.get("polygon"))
        if b is None:
            continue
        cx = (b[0] + b[2]) / 2.0
        cy = (b[1] + b[3]) / 2.0
        d = ((cx - bx) ** 2) + ((cy - by) ** 2)
        if d < best_dist:
            best_dist = d
            best = normalize_text(str(row.get("text", "")))
    return best


def collect_ocr_text_rows(node: Any, out: List[Dict[str, Any]]) -> None:
    if isinstance(node, dict):
        has_text = "text" in node and node.get("text") not in (None, "")
        has_box = norm_bbox(node.get("bbox")) is not None or norm_bbox(node.get("polygon")) is not None
        if has_text and has_box:
            out.append(node)
        for v in node.values():
            collect_ocr_text_rows(v, out)
        return
    if isinstance(node, list):
        for item in node:
            collect_ocr_text_rows(item, out)


def group_ocr_rows_by_page(ocr_payload: Any, doc_key: Optional[str]) -> Dict[int, List[Dict[str, Any]]]:
    if ocr_payload is None:
        return {}
    _, raw = pick_doc_payload(ocr_payload, doc_key)
    out: Dict[int, List[Dict[str, Any]]] = {}
    if not isinstance(raw, list):
        return out
    for i, page_node in enumerate(raw, start=1):
        rows: List[Dict[str, Any]] = []
        collect_ocr_text_rows(page_node, rows)
        for idx, row in enumerate(rows):
            row.setdefault("order_index", idx)
        out[i] = rows
    return out


def collect_ocr_rows_for_block(
    block_bbox: List[float],
    ocr_items: Sequence[Dict[str, Any]],
) -> List[Dict[str, Any]]:
    hits: List[Tuple[int, Dict[str, Any]]] = []
    for row in ocr_items:
        bbox = norm_bbox(row.get("bbox")) or norm_bbox(row.get("polygon"))
        if bbox is None:
            continue
        ratio = overlap_ratio(block_bbox, bbox)
        cx = (bbox[0] + bbox[2]) / 2.0
        cy = (bbox[1] + bbox[3]) / 2.0
        center_inside = block_bbox[0] <= cx <= block_bbox[2] and block_bbox[1] <= cy <= block_bbox[3]
        if ratio >= 0.2 or center_inside:
            hits.append((safe_int(row.get("order_index"), 10**9), row))
    hits.sort(key=lambda item: item[0])
    return [row for _, row in hits]


def join_ocr_text(rows: Sequence[Dict[str, Any]]) -> str:
    texts: List[str] = []
    seen = set()
    for row in rows:
        text = normalize_text(str(row.get("text", "")))
        if not text or text in seen:
            continue
        seen.add(text)
        texts.append(text)
    return " ".join(texts).strip()


def match_layout_to_xml(
    block_bbox: List[float],
    image_bbox: List[float],
    xml_objects: Sequence[XmlObject],
    slide_w: float,
    slide_h: float,
    block_text: str = "",
) -> Optional[XmlObject]:
    if not xml_objects:
        return None
    best: Optional[XmlObject] = None
    best_score = -1.0
    for obj in xml_objects:
        obj_bbox = xml_bbox_to_normalized(obj, slide_w, slide_h)
        if obj_bbox is None:
            continue
        score = (0.65 * overlap_ratio(block_bbox, obj_bbox)) + (0.20 * center_distance_score(block_bbox, obj_bbox))
        if block_text and obj.text:
            score += 0.25 * text_similarity(block_text, obj.text)
        if obj.tag in {"pic", "graphicFrame", "cxnSp"}:
            score -= 0.10
        if score > best_score:
            best_score = score
            best = obj
    return best


def match_ocr_row_to_xml(
    ocr_row: Dict[str, Any],
    xml_objects: Sequence[XmlObject],
    slide_w: float,
    slide_h: float,
) -> Optional[XmlObject]:
    row_bbox = norm_bbox(ocr_row.get("bbox")) or norm_bbox(ocr_row.get("polygon"))
    if row_bbox is None:
        return None
    row_text = normalize_text(str(ocr_row.get("text", "")))
    best: Optional[XmlObject] = None
    best_score = -1.0
    for obj in xml_objects:
        if not normalize_text(obj.text):
            continue
        obj_bbox = xml_bbox_to_normalized(obj, slide_w, slide_h)
        if obj_bbox is None:
            continue
        score = (0.50 * overlap_ratio(row_bbox, obj_bbox)) + (0.20 * center_distance_score(row_bbox, obj_bbox))
        if row_text and obj.text:
            score += 0.35 * text_similarity(row_text, obj.text)
        if obj.ph_type in TITLE_TYPES:
            score += 0.05
        if score > best_score:
            best_score = score
            best = obj
    if best_score < 0.25:
        return None
    return best


def build_reading_order_from_ocr(
    layout_rows: Sequence[Dict[str, Any]],
    page_ocr_rows: Sequence[Dict[str, Any]],
    xml_objects: Sequence[XmlObject],
    slide_w: float,
    slide_h: float,
) -> List[Dict[str, Any]]:
    if not page_ocr_rows or not xml_objects:
        return []

    by_shape_id = {obj.shape_id: obj for obj in xml_objects if obj.shape_id}
    matched_ocr: Dict[str, Dict[str, Any]] = {}
    for row in page_ocr_rows:
        obj = match_ocr_row_to_xml(row, xml_objects, slide_w, slide_h)
        if obj is None or not obj.shape_id:
            continue
        shape_id = obj.shape_id
        row_bbox = norm_bbox(row.get("bbox")) or norm_bbox(row.get("polygon"))
        item = matched_ocr.setdefault(
            shape_id,
            {
                "obj": obj,
                "position": safe_int(row.get("order_index"), 10**9),
                "texts": [],
                "bboxes": [],
            },
        )
        item["position"] = min(item["position"], safe_int(row.get("order_index"), 10**9))
        text = normalize_text(str(row.get("text", "")))
        if text and text not in item["texts"]:
            item["texts"].append(text)
        if row_bbox is not None:
            item["bboxes"].append(row_bbox)

    support_rows: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    fallback_rows: List[Dict[str, Any]] = []
    for row in layout_rows:
        shape_id = str(row.get("matched_shape_id") or "").strip()
        if shape_id:
            support_rows[shape_id].append(row)
        else:
            fallback_rows.append(row)

    reading: List[Dict[str, Any]] = []
    covered_shape_ids = set()
    for shape_id, item in matched_ocr.items():
        obj = item["obj"]
        support = None
        candidates = support_rows.get(shape_id, [])
        if candidates:
            candidates.sort(
                key=lambda r: (
                    safe_int(r.get("position"), 10**9),
                    -(float(((r.get("heading_decision") or {}).get("score", 0.0)))),
                )
            )
            support = candidates[0]
        reading.append(
            {
                "position": item["position"],
                "label": support.get("label") if support else "OCRText",
                "confidence": support.get("confidence") if support else 1.0,
                "bbox": merge_bboxes(item["bboxes"]) or xml_bbox_to_normalized(obj, slide_w, slide_h),
                "ocr_text": " ".join(item["texts"]).strip(),
                "xml_text": obj.text,
                "matched_shape_id": shape_id,
                "matched_placeholder_type": obj.ph_type,
                "matched_font_pt": obj.font_pt,
                "heading_decision": (support.get("heading_decision") if support else {
                    "is_heading": False,
                    "score": 0.0,
                    "threshold": 0.0,
                    "level": 0,
                    "reasons": ["No matching layout block for OCR-derived shape"],
                }),
            }
        )
        covered_shape_ids.add(shape_id)

    for shape_id, rows in support_rows.items():
        if shape_id in covered_shape_ids or shape_id not in by_shape_id:
            continue
        rows.sort(key=lambda r: safe_int(r.get("position"), 10**9))
        row = dict(rows[0])
        row["position"] = 10**6 + safe_int(row.get("position"), 10**5)
        reading.append(row)

    used_fallback_rows = set()
    for obj in xml_objects:
        shape_id = obj.shape_id
        if not shape_id or shape_id in covered_shape_ids:
            continue
        obj_text = normalize_text(obj.text)
        if not obj_text or obj.tag in {"pic", "graphicFrame", "cxnSp"}:
            continue
        obj_bbox = xml_bbox_to_normalized(obj, slide_w, slide_h)
        if obj_bbox is None:
            continue
        best_idx = None
        best_score = -1.0
        for idx, row in enumerate(fallback_rows):
            if idx in used_fallback_rows:
                continue
            row_bbox = row.get("bbox") if isinstance(row.get("bbox"), list) else None
            row_text = normalize_text(str(row.get("ocr_text") or row.get("xml_text") or ""))
            score = (0.45 * overlap_ratio(obj_bbox, row_bbox)) + (0.15 * center_distance_score(obj_bbox, row_bbox))
            if row_text:
                score += 0.40 * text_similarity(obj_text, row_text)
            if score > best_score:
                best_score = score
                best_idx = idx
        if best_idx is None or best_score < 0.35:
            continue
        used_fallback_rows.add(best_idx)
        row = dict(fallback_rows[best_idx])
        row["matched_shape_id"] = shape_id
        row["matched_placeholder_type"] = obj.ph_type
        row["matched_font_pt"] = obj.font_pt
        row["xml_text"] = obj.text
        row["position"] = 10**5 + safe_int(row.get("position"), 10**5)
        reading.append(row)
        covered_shape_ids.add(shape_id)

    reading.sort(
        key=lambda x: (
            safe_int(x.get("position"), 10**9),
            ((x.get("bbox") or [0.0, 0.0, 0.0, 0.0])[1]),
            ((x.get("bbox") or [0.0, 0.0, 0.0, 0.0])[0]),
        )
    )
    return reading


def score_heading(
    label: str,
    confidence: float,
    block_bbox: List[float],
    image_bbox: List[float],
    text: str,
    xml_obj: Optional[XmlObject],
) -> Tuple[float, int, List[str]]:
    score = 0.0
    reasons: List[str] = []
    heading_level = 0

    if label in {"SectionHeader", "PageHeader"}:
        score += 0.45
        reasons.append("layout label indicates header")
    if confidence >= 0.98:
        score += 0.15
        reasons.append("high layout confidence")
    elif confidence >= 0.90:
        score += 0.08
        reasons.append("good layout confidence")

    raw = normalize_text(text)
    if raw:
        if 3 <= len(raw) <= 80:
            score += 0.08
            reasons.append("short heading-like text length")
        elif len(raw) > 140:
            score -= 0.20
            reasons.append("too long for heading")
        num_depth = parse_numbered_heading(raw)
        if num_depth:
            score += 0.22
            heading_level = num_depth
            reasons.append("numbered heading pattern")

    ih = max(image_bbox[3] - image_bbox[1], 1.0)
    y_ratio = (block_bbox[1] - image_bbox[1]) / ih
    if y_ratio <= 0.20:
        score += 0.18
        reasons.append("top area on slide")
    elif y_ratio <= 0.35:
        score += 0.08
        reasons.append("upper area on slide")
    else:
        score -= 0.10
        reasons.append("mid/lower area penalty")

    if label in {"PageFooter", "Footnote"}:
        score -= 0.50
        reasons.append("footer/footnote label")

    if xml_obj is not None:
        if xml_obj.ph_type in TITLE_TYPES:
            score += 0.35
            reasons.append("ppt placeholder is title/subtitle")
            if heading_level == 0:
                heading_level = 1
        if xml_obj.font_pt is not None:
            if xml_obj.font_pt >= 26:
                score += 0.20
                reasons.append("large font size")
                if heading_level == 0:
                    heading_level = 1
            elif xml_obj.font_pt >= 20:
                score += 0.12
                reasons.append("medium-large font size")
                if heading_level == 0:
                    heading_level = 2
            elif xml_obj.font_pt <= 12:
                score -= 0.08
                reasons.append("small font penalty")
        if xml_obj.tag in {"pic", "graphicFrame"}:
            score -= 0.15
            reasons.append("non-text object penalty")

    if heading_level == 0 and label in {"SectionHeader", "PageHeader"}:
        heading_level = 2
    return score, heading_level, reasons


def main() -> int:
    script_dir = Path(__file__).resolve().parent

    parser = argparse.ArgumentParser(description="Normalize Surya layout/table/ocr outputs.")
    parser.add_argument(
        "--layout-json",
        default=None,
        help="Path to layout results.json OR layout result directory. If omitted, auto-picks under ./output/layout_result.",
    )
    parser.add_argument(
        "--table-json",
        default=None,
        help="Path to table results.json OR table result directory. If omitted, auto-picks under ./output/table.",
    )
    parser.add_argument(
        "--ocr-json",
        default=None,
        help="Optional path to Surya OCR results.json",
    )
    parser.add_argument(
        "--pptx-path",
        default=None,
        help="Optional path to source .pptx. If set, script extracts XML internally.",
    )
    parser.add_argument(
        "--ppt-root",
        default=None,
        help="Path to extracted pptx root directory (contains ppt/slides).",
    )
    parser.add_argument(
        "--output-json",
        default=str(script_dir / "output" / "normalized" / "normalized_results.json"),
        help="Path to output normalized json",
    )
    parser.add_argument(
        "--doc-key",
        default=None,
        help="Optional document key in input JSON. If omitted, first key is used.",
    )
    parser.add_argument(
        "--exclude-reading-labels",
        default="PageFooter,Footnote",
        help="Comma-separated labels excluded from reading-order output.",
    )
    parser.add_argument(
        "--final-heading-threshold",
        type=float,
        default=0.55,
        help="Threshold for final heading decision score.",
    )
    args = parser.parse_args()

    default_layout_dir = script_dir / "output" / "layout_result"
    default_table_dir = script_dir / "output" / "table"
    layout_path = resolve_results_json(args.layout_json, default_layout_dir, kind="layout")
    table_path = resolve_results_json(args.table_json, default_table_dir, kind="table")
    output_path = Path(args.output_json).resolve()
    ocr_path = Path(args.ocr_json).resolve() if args.ocr_json else None
    pptx_path = Path(args.pptx_path).resolve() if args.pptx_path else None

    temp_ppt_dir: Optional[Path] = None
    if args.ppt_root:
        ppt_root = Path(args.ppt_root).resolve()
    elif pptx_path is not None:
        temp_ppt_dir, ppt_root = extract_pptx_to_temp_root(pptx_path)
    else:
        raise ValueError("either --ppt-root or --pptx-path is required")

    layout_payload = load_json(layout_path)
    table_payload = load_json(table_path)
    ocr_payload = load_json(ocr_path) if ocr_path and ocr_path.exists() else None
    layout_key, layout_pages_raw = pick_doc_payload(layout_payload, args.doc_key)
    table_key, table_items_raw = pick_table_payload(table_payload, args.doc_key, fallback_doc_key=layout_key)

    if not isinstance(layout_pages_raw, list):
        raise ValueError("layout payload doc value must be a list")
    if not isinstance(table_items_raw, list):
        table_items_raw = []

    ocr_by_page = group_ocr_rows_by_page(ocr_payload, args.doc_key)
    slide_w, slide_h = parse_slide_size_emu(ppt_root)

    excluded_labels = {x.strip() for x in args.exclude_reading_labels.split(",") if x.strip()}
    pages: Dict[int, Dict[str, Any]] = {}

    for page in layout_pages_raw:
        if not isinstance(page, dict):
            continue
        page_no = safe_int(page.get("page"), 0)
        if page_no <= 0:
            continue
        image_bbox = norm_bbox(page.get("image_bbox")) or [0.0, 0.0, 1.0, 1.0]

        slide_xml = ppt_root / "ppt" / "slides" / f"slide{page_no}.xml"
        xml_objects = parse_slide_xml_objects(slide_xml) if slide_xml.exists() else []
        page_ocr_rows = ocr_by_page.get(page_no, [])

        bboxes = page.get("bboxes", [])
        if not isinstance(bboxes, list):
            bboxes = []

        layout_rows: List[Dict[str, Any]] = []
        final_headings: List[Dict[str, Any]] = []

        for blk in bboxes:
            if not isinstance(blk, dict):
                continue
            label = str(blk.get("label", "")).strip()
            bbox = norm_bbox(blk.get("bbox"))
            if bbox is None:
                continue
            position = safe_int(blk.get("position"), 10**9)
            conf = safe_float(blk.get("confidence"), 0.0)

            block_ocr_rows = collect_ocr_rows_for_block(bbox, page_ocr_rows) if page_ocr_rows else []
            ocr_text = join_ocr_text(block_ocr_rows) if block_ocr_rows else ""
            xml_obj = match_layout_to_xml(bbox, image_bbox, xml_objects, slide_w, slide_h, ocr_text)
            xml_text = xml_obj.text if xml_obj else ""
            text_for_heading = ocr_text or xml_text
            score, heading_level, reasons = score_heading(
                label=label,
                confidence=conf,
                block_bbox=bbox,
                image_bbox=image_bbox,
                text=text_for_heading,
                xml_obj=xml_obj,
            )
            is_final_heading = score >= args.final_heading_threshold

            row = {
                "position": (
                    min((safe_int(r.get("order_index"), 10**9) for r in block_ocr_rows), default=10**6 + position)
                ),
                "label": label,
                "confidence": conf,
                "bbox": bbox,
                "ocr_text": ocr_text,
                "xml_text": xml_text,
                "matched_shape_id": xml_obj.shape_id if xml_obj else None,
                "matched_placeholder_type": xml_obj.ph_type if xml_obj else None,
                "matched_font_pt": xml_obj.font_pt if xml_obj else None,
                "heading_decision": {
                    "is_heading": is_final_heading,
                    "score": round(score, 4),
                    "threshold": args.final_heading_threshold,
                    "level": heading_level if is_final_heading else 0,
                    "reasons": reasons,
                },
            }

            if label not in excluded_labels:
                layout_rows.append(row)
            if is_final_heading:
                final_headings.append(row)

        reading = build_reading_order_from_ocr(
            layout_rows=layout_rows,
            page_ocr_rows=page_ocr_rows,
            xml_objects=xml_objects,
            slide_w=slide_w,
            slide_h=slide_h,
        )
        if not reading:
            reading = list(layout_rows)

        reading.sort(key=lambda x: (x["position"], x["bbox"][1], x["bbox"][0]))
        final_headings.sort(key=lambda x: (x["position"], x["bbox"][1], x["bbox"][0]))

        pages[page_no] = {
            "page": page_no,
            "image_bbox": image_bbox,
            "reading_order": [
                {
                    "order_index": idx,
                    **r,
                }
                for idx, r in enumerate(reading)
            ],
            "headings": final_headings,
            "tables": [],
        }

    for item in table_items_raw:
        if not isinstance(item, dict):
            continue
        page_no = safe_int(item.get("page"), 0)
        if page_no <= 0:
            continue
        if page_no not in pages:
            pages[page_no] = {
                "page": page_no,
                "image_bbox": norm_bbox(item.get("image_bbox")) or [0.0, 0.0, 1.0, 1.0],
                "reading_order": [],
                "headings": [],
                "tables": [],
            }
        table_bbox = table_bbox_from_item(item)
        pages[page_no]["tables"].append(
            {
                "table_idx": safe_int(item.get("table_idx"), len(pages[page_no]["tables"])),
                "bbox": table_bbox,
                "row_count": len(item.get("rows", [])) if isinstance(item.get("rows"), list) else 0,
                "col_count": len(item.get("cols", [])) if isinstance(item.get("cols"), list) else 0,
            }
        )

    ordered_pages = [pages[k] for k in sorted(pages.keys())]
    for p in ordered_pages:
        p["tables"].sort(key=lambda t: (safe_int(t.get("table_idx"), 10**9), (t.get("bbox") or [0, 0, 0, 0])[1]))
        p["counts"] = {
            "reading_blocks": len(p["reading_order"]),
            "headings": len(p["headings"]),
            "tables": len(p["tables"]),
        }

    out = {
        "source": {
            "layout_json": str(layout_path),
            "layout_doc_key": layout_key,
            "table_json": str(table_path),
            "table_doc_key": table_key,
            "ocr_json": str(ocr_path) if ocr_path and ocr_path.exists() else None,
            "pptx_path": str(pptx_path) if pptx_path and pptx_path.exists() else None,
            "ppt_root": str(ppt_root),
        },
        "rules": {
            "final_heading_threshold": args.final_heading_threshold,
            "exclude_reading_labels": sorted(excluded_labels),
            "signals": [
                "layout_label",
                "layout_confidence",
                "position_y_ratio",
                "text_length",
                "numbered_pattern",
                "xml_placeholder",
                "xml_font_size",
                "xml_object_type",
                "ocr_text",
            ],
        },
        "pages": ordered_pages,
        "summary": {
            "page_count": len(ordered_pages),
            "reading_blocks_total": sum(len(p["reading_order"]) for p in ordered_pages),
            "headings_total": sum(len(p["headings"]) for p in ordered_pages),
            "tables_total": sum(len(p["tables"]) for p in ordered_pages),
        },
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote: {output_path}")
    print(
        f"Summary: pages={out['summary']['page_count']} "
        f"reading_blocks={out['summary']['reading_blocks_total']} "
        f"headings={out['summary']['headings_total']} "
        f"tables={out['summary']['tables_total']}"
    )

    if temp_ppt_dir is not None and temp_ppt_dir.exists():
        shutil.rmtree(temp_ppt_dir, ignore_errors=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

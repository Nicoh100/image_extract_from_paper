from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
import pdfplumber
from rich.console import Console

from .types import AssetKind
from .utils import clamp_box_to_page, expand_box, mm_to_points, non_max_suppression

console = Console()


def _blocks_rawdict(page: fitz.Page) -> Dict:
    return page.get_text("rawdict")


def _detect_image_blocks(page: fitz.Page) -> List[Tuple[List[float], float]]:
    rd = _blocks_rawdict(page)
    boxes: List[List[float]] = []
    scores: List[float] = []
    for block in rd.get("blocks", []):
        if block.get("type") == 1:  # image
            bbox = [float(x) for x in block.get("bbox", [0, 0, 0, 0])]
            w = max(1.0, bbox[2] - bbox[0])
            h = max(1.0, bbox[3] - bbox[1])
            area = w * h
            boxes.append(bbox)
            scores.append(area)
    keep_idx = non_max_suppression(boxes, scores, iou_threshold=0.3)
    return [(boxes[i], scores[i]) for i in keep_idx]


def _detect_tables_pdfplumber(pdf_path: str, page_index: int) -> List[List[float]]:
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_index < 0 or page_index >= len(pdf.pages):
                return []
            page = pdf.pages[page_index]
            tables = page.find_tables()  # heuristic
            bboxes: List[List[float]] = []
            for t in tables:
                if hasattr(t, "bbox") and t.bbox:
                    # pdfplumber bbox is (x0, top, x1, bottom) in PDF points
                    bboxes.append([float(t.bbox[0]), float(t.bbox[1]), float(t.bbox[2]), float(t.bbox[3])])
            return bboxes
    except Exception:
        return []


def detect_assets_on_page(
    doc: fitz.Document,
    page_index: int,
    pdf_path: str,
    pad_mm: float = 4.0,
) -> List[Tuple[AssetKind, List[float]]]:
    page = doc[page_index]
    page_rect = [page.rect.x0, page.rect.y0, page.rect.x1, page.rect.y1]

    # Images from PyMuPDF rawdict
    img_candidates = _detect_image_blocks(page)  # (bbox, score)
    img_boxes = [b for b, _ in img_candidates]

    # Tables from pdfplumber
    table_boxes = _detect_tables_pdfplumber(pdf_path, page_index)

    # Merge with simple NMS per class
    pad = mm_to_points(pad_mm)

    final: List[Tuple[AssetKind, List[float]]] = []

    if img_boxes:
        scores = [(b[2] - b[0]) * (b[3] - b[1]) for b in img_boxes]
        keep = non_max_suppression(img_boxes, scores, iou_threshold=0.3)
        for i in keep:
            b = clamp_box_to_page(expand_box(img_boxes[i], pad), page_rect)
            final.append(("image", b))

    if table_boxes:
        scores = [(b[2] - b[0]) * (b[3] - b[1]) for b in table_boxes]
        keep = non_max_suppression(table_boxes, scores, iou_threshold=0.2)
        for i in keep:
            b = clamp_box_to_page(expand_box(table_boxes[i], pad), page_rect)
            final.append(("table", b))

    # If nothing found, return empty (MVP doesn't force figures)
    return final


CAPTION_RE = re.compile(r"\b(?:Figure|Fig\.\s*|Abb\.|Abbildung|Table|Tab\.|Tabelle)\s*\d+", re.I)


def find_caption_near_bbox(page: fitz.Page, bbox: Sequence[float], radius_mm: float) -> Optional[str]:
    radius_pt = mm_to_points(radius_mm)
    x0, y0, x1, y1 = bbox
    search_rect = fitz.Rect(x0 - radius_pt, y0 - radius_pt, x1 + radius_pt, y1 + radius_pt)
    blocks = page.get_text("blocks")
    nearby_lines: List[str] = []
    for b in blocks:
        bx0, by0, bx1, by1, text, *_ = b
        rect = fitz.Rect(bx0, by0, bx1, by1)
        if rect.intersects(search_rect):
            s = (text or "").strip()
            if s:
                nearby_lines.append(s)
    for line in nearby_lines:
        if CAPTION_RE.search(line):
            return line
    # fallback: join first few lines
    if nearby_lines:
        return nearby_lines[0][:200]
    return None

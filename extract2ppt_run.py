from __future__ import annotations
import argparse, json, os
from pathlib import Path
from typing import List, Sequence, Tuple, Optional, Dict, DefaultDict
import re
import fitz
import pdfplumber
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from rich.console import Console
from tqdm import tqdm
import numpy as np
import cv2

console = Console()

def mm_to_points(mm: float) -> float:
    return mm * 72.0 / 25.4

def expand_box(bbox: Sequence[float], pad_points: float, page_rect: fitz.Rect) -> List[float]:
    x0, y0, x1, y1 = bbox
    x0 -= pad_points; y0 -= pad_points; x1 += pad_points; y1 += pad_points
    x0 = max(page_rect.x0, x0)
    y0 = max(page_rect.y0, y0)
    x1 = min(page_rect.x1, x1)
    y1 = min(page_rect.y1, y1)
    return [float(x0), float(y0), float(x1), float(y1)]

def compute_iou(a,b):
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    inter_x0 = max(ax0, bx0)
    inter_y0 = max(ay0, by0)
    inter_x1 = min(ax1, bx1)
    inter_y1 = min(ay1, by1)
    inter_w = max(0.0, inter_x1 - inter_x0)
    inter_h = max(0.0, inter_y1 - inter_y0)
    inter = inter_w * inter_h
    if inter <= 0: return 0.0
    area_a = (ax1-ax0)*(ay1-ay0)
    area_b = (bx1-bx0)*(by1-by0)
    union = area_a + area_b - inter
    return inter/union if union>0 else 0.0

def nms(boxes, scores, thr=0.3):
    order = sorted(range(len(boxes)), key=lambda i: scores[i], reverse=True)
    keep=[]
    while order:
        i = order.pop(0)
        keep.append(i)
        new=[]
        for j in order:
            if compute_iou(boxes[i], boxes[j]) <= thr:
                new.append(j)
        order=new
    return keep

def union_boxes(boxes: List[Sequence[float]]) -> List[float]:
    x0 = min(b[0] for b in boxes)
    y0 = min(b[1] for b in boxes)
    x1 = max(b[2] for b in boxes)
    y1 = max(b[3] for b in boxes)
    return [float(x0), float(y0), float(x1), float(y1)]

def detect_image_boxes(page: fitz.Page):
    rd = page.get_text('rawdict')
    boxes = []
    scores = []
    for block in rd.get('blocks', []):
        if block.get('type') == 1:
            bbox = [float(x) for x in block.get('bbox', [0,0,0,0])]
            w = max(1.0, bbox[2]-bbox[0]); h = max(1.0, bbox[3]-bbox[1])
            if w < 16 or h < 16:
                continue
            boxes.append(bbox); scores.append(w*h)
    keep = nms(boxes, scores, 0.3)
    return [boxes[i] for i in keep]

def detect_figure_drawing_boxes(page: fitz.Page, page_area_threshold: float = 0.005):
    rects=[]; scores=[]
    page_area = float(page.rect.width * page.rect.height)
    try:
        drawings = page.get_drawings()
    except Exception:
        drawings = []
    for d in drawings:
        r = d.get('rect', None)
        if not r:
            try:
                pts = []
                for it in d.get('items', []):
                    for p in it[1]:
                        pts.append((p.x, p.y))
                if not pts:
                    continue
                xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
                r = fitz.Rect(min(xs), min(ys), max(xs), max(ys))
            except Exception:
                continue
        w = float(r.x1 - r.x0); h = float(r.y1 - r.y0)
        area = w*h
        if w < 20 or h < 20:
            continue
        if area / max(1.0, page_area) < page_area_threshold:
            continue
        rects.append([float(r.x0), float(r.y0), float(r.x1), float(r.y1)])
        scores.append(area)
    if not rects:
        return []
    keep = nms(rects, scores, thr=0.2)
    return [rects[i] for i in keep]

CAPTION_FIG_RE = re.compile(r"\b(?:fig(?:ure)?|abb(?:ildung)?)\s*(\d+)", re.I)
CAPTION_TAB_RE = re.compile(r"\b(?:tab(?:le)?|tabelle)\s*(\d+)", re.I)
CONTINUED_RE = re.compile(r"\bcontinued\b", re.I)

def parse_caption_label(text: Optional[str]) -> Optional[Tuple[str,int,bool]]:
    if not text:
        return None
    cont = bool(CONTINUED_RE.search(text))
    m = CAPTION_FIG_RE.search(text)
    if m:
        try:
            return ("figure", int(m.group(1)), cont)
        except Exception:
            pass
    m = CAPTION_TAB_RE.search(text)
    if m:
        try:
            return ("table", int(m.group(1)), cont)
        except Exception:
            pass
    return None

def find_caption_line(page: fitz.Page, bbox, radius_mm: float) -> Tuple[Optional[str], Optional[List[float]]]:
    radius_pt = mm_to_points(radius_mm)
    x0,y0,x1,y1 = bbox
    search_rect = fitz.Rect(x0 - radius_pt, y0 - radius_pt, x1 + radius_pt, y1 + radius_pt)
    blocks = page.get_text('blocks')
    nearby=[]
    for b in blocks:
        bx0, by0, bx1, by1, text, *_ = b
        rect = fitz.Rect(bx0, by0, bx1, by1)
        if rect.intersects(search_rect):
            s = (text or '').strip()
            if s:
                nearby.append((s, [float(bx0), float(by0), float(bx1), float(by1)]))
    for text, rect in nearby:
        if CAPTION_FIG_RE.search(text) or CAPTION_TAB_RE.search(text):
            return text, rect
    if nearby:
        return nearby[0][0], nearby[0][1]
    return None, None


def render_region(page: fitz.Page, rect: Sequence[float], render_dpi: int) -> Image.Image:
    clip = fitz.Rect(*rect)
    scale = render_dpi / 72.0
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
    mode = 'RGB' if pix.n < 4 else 'RGBA'
    img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if mode == 'RGBA':
        img = img.convert('RGB')
    return img


def smart_trim_table_pil(pil_img: Image.Image) -> Image.Image:
    arr = np.array(pil_img)
    if arr.ndim == 3 and arr.shape[2] == 4:
        arr = cv2.cvtColor(arr, cv2.COLOR_RGBA2RGB)
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    _, bin_inv = cv2.threshold(gray, 248, 255, cv2.THRESH_BINARY_INV)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (7, 7))
    mask = cv2.morphologyEx(bin_inv, cv2.MORPH_CLOSE, kernel, iterations=1)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return pil_img
    areas = [cv2.contourArea(c) for c in contours]
    idx = int(np.argmax(areas))
    x, y, w, h = cv2.boundingRect(contours[idx])
    pad_in = 1
    x = max(0, x + pad_in)
    y = max(0, y + pad_in)
    w = max(1, w - 2 * pad_in)
    h = max(1, h - 2 * pad_in)
    cropped = arr[y:y+h, x:x+w]
    return Image.fromarray(cropped)


def page_caption_lines(page: fitz.Page) -> List[Tuple[str, List[float]]]:
    out = []
    for bx0, by0, bx1, by1, text, *_ in page.get_text('blocks'):
        s = (text or '').strip()
        if not s:
            continue
        if CAPTION_FIG_RE.search(s) or CAPTION_TAB_RE.search(s):
            out.append((s, [float(bx0), float(by0), float(bx1), float(by1)]))
    return out


def refine_caption_rect(page: fitz.Page, cap_rect: Sequence[float], cap_text: Optional[str]) -> List[float]:
    """Return a very tight rectangle around just the 'Fig N …' line (± one wrapped line).

    This ignores block-level boxes and derives the bottom from the exact line glyphs.
    """
    rd = page.get_text("rawdict")
    lines: List[Tuple[float, float, str, List[float], float]] = []  # (y0, x0, text, rect, size)
    # Only consider lines that intersect the original caption rect
    capR = fitz.Rect(float(cap_rect[0]), float(cap_rect[1]), float(cap_rect[2]), float(cap_rect[3]))
    for block in rd.get("blocks", []):
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue
            bx0 = min(float(s.get("bbox", [0, 0, 0, 0])[0]) for s in spans)
            by0 = min(float(s.get("bbox", [0, 0, 0, 0])[1]) for s in spans)
            bx1 = max(float(s.get("bbox", [0, 0, 0, 0])[2]) for s in spans)
            by1 = max(float(s.get("bbox", [0, 0, 0, 0])[3]) for s in spans)
            rect = fitz.Rect(bx0, by0, bx1, by1)
            if not rect.intersects(capR):
                continue
            text = "".join(s.get("text", "") for s in spans).strip()
            sizes = [float(s.get("size", 0.0)) for s in spans if float(s.get("size", 0.0)) > 0]
            size = float(np.median(sizes)) if sizes else 0.0
            lines.append((by0, bx0, text, [bx0, by0, bx1, by1], size))

    if not lines:
        # fallback: very small height around original top
        tiny_h = mm_to_points(6)
        return [float(cap_rect[0]), float(cap_rect[1]), float(cap_rect[2]), float(cap_rect[1]) + tiny_h]

    lines.sort(key=lambda x: (x[0], x[1]))
    start_idx = None
    for i, (_, _, t, r, sz) in enumerate(lines):
        if CAPTION_FIG_RE.search(t or ""):
            start_idx = i
            break
    if start_idx is None:
        # fallback: choose the top-most intersecting line
        y0, x0, t, r, sz = lines[0]
        return [r[0], r[1], r[2], r[3]]

    used = [lines[start_idx]]
    base_size = used[0][4]
    base_left = used[0][1]
    prev_bottom = used[0][3][3]
    # Consider at most one wrap line if alignment and size match
    if start_idx + 1 < len(lines):
        _, _, t2, r2, sz2 = lines[start_idx + 1]
        head = (t2.split()[0].lower() if t2 else "")
        cond_head = head not in {"results", "methods", "introduction", "discussion", "conclusion", "acknowledgments"}
        cond_align = abs(r2[0] - base_left) <= mm_to_points(3)
        cond_size = sz2 <= base_size + 0.4
        cond_gap = (r2[1] - prev_bottom) <= (r2[3] - r2[1]) * 0.4
        if cond_head and cond_align and cond_size and cond_gap:
            used.append(lines[start_idx + 1])
            prev_bottom = r2[3]

    bx0 = min(u[3][0] for u in used)
    by0 = min(u[3][1] for u in used)
    bx1 = max(u[3][2] for u in used)
    by1 = max(u[3][3] for u in used)
    return [bx0, by0, bx1, by1]

def select_plumber_table_below_caption(pp_page, y_cap: float) -> Optional[List[float]]:
    best = None
    best_delta = None
    try:
        tables = pp_page.find_tables()
        for t in tables:
            if getattr(t, 'bbox', None):
                tb = [float(t.bbox[0]), float(t.bbox[1]), float(t.bbox[2]), float(t.bbox[3])]
                if tb[1] >= y_cap - 5:
                    delta = tb[1] - y_cap
                    if delta < 0:
                        continue
                    if best_delta is None or delta < best_delta:
                        best = tb
                        best_delta = delta
        if best is None and tables:
            best = max(( [float(t.bbox[0]), float(t.bbox[1]), float(t.bbox[2]), float(t.bbox[3])] for t in tables if getattr(t, 'bbox', None) ), key=lambda b:(b[2]-b[0])*(b[3]-b[1]), default=None)
    except Exception:
        return None
    return best


def find_figure_region(page: fitz.Page, cap_rect: Sequence[float], render_dpi: int) -> Optional[List[float]]:
    # Heuristic: look for drawings/images within a compact vertical band just above the caption.
    # This avoids capturing entire columns of body text and focuses on the cluster directly tied to the caption.
    page_height = float(page.rect.y1 - page.rect.y0)
    page_width = float(page.rect.x1 - page.rect.x0)

    # Band height limited by both relative page height and an absolute mm cap (wider to capture full diagram)
    band_height = min(0.80 * page_height, mm_to_points(220))
    band_top = max(page.rect.y0, float(cap_rect[3]) - band_height)
    band = [page.rect.x0 + mm_to_points(6), band_top, page.rect.x1 - mm_to_points(6), float(cap_rect[3])]

    draw_boxes = detect_figure_drawing_boxes(page)
    img_boxes = detect_image_boxes(page)

    min_width = 0.18 * page_width  # accept narrower primitives than before; figures can be composed of many medium boxes
    candidates: List[List[float]] = []
    for b in draw_boxes + img_boxes:
        # Keep items that lie inside the band and are above the caption line
        if b[3] <= float(cap_rect[3]) + mm_to_points(2) and b[1] >= band_top - mm_to_points(6):
            if (b[2] - b[0]) >= min_width:
                # Slightly grow each box to better connect nearby panel elements
                grown = [b[0] - mm_to_points(1.5), b[1] - mm_to_points(1.5), b[2] + mm_to_points(1.5), b[3] + mm_to_points(1.5)]
                candidates.append(grown)

    if candidates:
        bbox = union_boxes(candidates)
        # Clamp to the search band
        return [max(band[0], bbox[0]), max(band[1], bbox[1]), min(band[2], bbox[2]), min(band[3], bbox[3])]

    # Fallback: render the band and detect the largest connected component of ink
    band_img = render_region(page, band, render_dpi)
    arr = np.array(band_img.convert("RGB"))
    gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY)
    _, bin_inv = cv2.threshold(gray, 248, 255, cv2.THRESH_BINARY_INV)
    # Close small gaps so a multi-panel figure becomes a single blob
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (9, 9))
    mask = cv2.morphologyEx(bin_inv, cv2.MORPH_CLOSE, kernel, iterations=1)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return [band[0], band[1], band[2], band[3]]
    areas = [cv2.contourArea(c) for c in contours]
    idx = int(np.argmax(areas))
    x, y, w, h = cv2.boundingRect(contours[idx])
    # Map back to page coords
    scale = (72.0 / render_dpi)
    fig_rect = [band[0] + x * scale, band[1] + y * scale, band[0] + (x + w) * scale, band[1] + (y + h) * scale]
    return fig_rect


def page_has_figure_candidates(page: fitz.Page, cap_rect: Sequence[float]) -> bool:
    """Return True if drawings/images likely form a figure above the caption on this page.

    This guards against picking up text-only pages that merely mention the figure in body text.
    """
    page_height = float(page.rect.y1 - page.rect.y0)
    page_width = float(page.rect.x1 - page.rect.x0)
    band_height = min(0.55 * page_height, mm_to_points(160))
    band_top = max(page.rect.y0, float(cap_rect[3]) - band_height)

    draw_boxes = detect_figure_drawing_boxes(page)
    img_boxes = detect_image_boxes(page)
    min_width = 0.18 * page_width

    for b in draw_boxes + img_boxes:
        if b[3] <= float(cap_rect[3]) + mm_to_points(2) and b[1] >= band_top - mm_to_points(6):
            if (b[2] - b[0]) >= min_width:
                return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('input_pdf')
    ap.add_argument('--output-dir', default='exports')
    ap.add_argument('--target-ppi', type=int, default=220)
    ap.add_argument('--min-ppi', type=int, default=180)
    ap.add_argument('--render-dpi', type=int, default=900)
    ap.add_argument('--caption-radius', type=float, default=60.0)
    ap.add_argument('--images-only', action='store_true', default=True)
    ap.add_argument('--pad-mm', type=float, default=2.0)
    args = ap.parse_args()

    config = {
        'source_pdf': os.path.abspath(args.input_pdf),
        'output_dir': os.path.abspath(args.output_dir),
        'target_ppi': args.target_ppi,
        'min_ppi': args.min_ppi,
        'render_dpi': args.render_dpi,
        'caption_radius_mm': args.caption_radius,
        'images_only': bool(args.images_only),
        'pad_mm': args.pad_mm,
    }
    pdf_path = Path(config['source_pdf'])
    exports_root = Path(config['output_dir'])
    figures_dir = exports_root / 'figures'
    tables_dir = exports_root / 'tables'
    for d in (figures_dir, tables_dir):
        d.mkdir(parents=True, exist_ok=True)

    grouped: Dict[str, Dict] = {}
    with fitz.open(str(pdf_path)) as doc, pdfplumber.open(str(pdf_path)) as plumber:
        for page_index in tqdm(range(len(doc)), desc='Pages', unit='p'):
            page = doc[page_index]
            for cap_text, cap_rect in page_caption_lines(page):
                parsed = parse_caption_label(cap_text)
                if not parsed:
                    continue
                label_kind, label_idx, is_continued = parsed
                key = f"{label_kind}:{label_idx}"
                segment = {'page': page_index, 'bbox': None, 'caption': cap_text, 'caption_rect': cap_rect, 'continued': is_continued}
                entry = grouped.get(key)
                if not entry:
                    grouped[key] = {'kind': label_kind, 'label_index': label_idx, 'segments': [segment]}
                else:
                    entry['segments'].append(segment)

            img_boxes = detect_image_boxes(page)
            draw_boxes = detect_figure_drawing_boxes(page)
            plumber_tables = []
            if page_index < len(plumber.pages):
                try:
                    ptables = plumber.pages[page_index].find_tables()
                    for t in ptables:
                        if getattr(t, 'bbox', None):
                            plumber_tables.append([float(t.bbox[0]), float(t.bbox[1]), float(t.bbox[2]), float(t.bbox[3])])
                except Exception:
                    pass
            candidates = [('image', b) for b in img_boxes] + [('figure', b) for b in draw_boxes] + [('table', b) for b in plumber_tables]

            for kind, bbox in candidates:
                cap_text, cap_rect = find_caption_line(page, bbox, config['caption_radius_mm'])
                parsed = parse_caption_label(cap_text)
                if not parsed:
                    continue
                label_kind, label_idx, is_continued = parsed
                key = f"{label_kind}:{label_idx}"
                segment = {'page': page_index, 'bbox': bbox, 'caption': cap_text, 'caption_rect': cap_rect, 'continued': is_continued}
                entry = grouped.get(key)
                if not entry:
                    grouped[key] = {'kind': label_kind, 'label_index': label_idx, 'segments': [segment]}
                else:
                    entry['segments'].append(segment)

        pad_pts_cache: Dict[int, float] = {}
        for key, entry in grouped.items():
            kind = entry['kind']
            page_to_items: DefaultDict[int, Dict[str, object]] = {}  # type: ignore
            page_order: List[int] = []
            for seg in entry['segments']:
                p = seg['page']
                if p not in page_to_items:
                    page_to_items[p] = {'boxes': [], 'caption': seg.get('caption'), 'caption_rect': seg.get('caption_rect')}
                    page_order.append(p)
                if seg.get('bbox') is not None:
                    page_to_items[p]['boxes'].append(seg['bbox'])  # type: ignore
                if seg.get('caption_rect'):
                    page_to_items[p]['caption_rect'] = seg['caption_rect']
                if seg.get('caption'):
                    page_to_items[p]['caption'] = seg['caption']
            page_order.sort()

            if kind == 'figure':
                rendered_imgs: List[Image.Image] = []
                widths: List[int] = []
                with fitz.open(str(pdf_path)) as doc2:
                    for p in page_order:
                        page = doc2[p]
                        pad_pts = pad_pts_cache.get(p)
                        if pad_pts is None:
                            pad_pts = mm_to_points(config['pad_mm'])
                            pad_pts_cache[p] = pad_pts
                        cap_rect = page_to_items[p].get('caption_rect')  # type: ignore
                        if not cap_rect:
                            continue
                        # Shrink caption box to just the caption lines
                        refined_cap = refine_caption_rect(page, cap_rect, page_to_items[p].get('caption'))  # type: ignore
                        # Skip pages that don't actually have figure candidates above the caption
                        if not page_has_figure_candidates(page, refined_cap):
                            continue
                        fig_rect = find_figure_region(page, refined_cap, config['render_dpi'])
                        if not fig_rect:
                            continue
                        # Bottom is strictly the caption bottom; allow full figure above
                        bbox = union_boxes([fig_rect, refined_cap])
                        bbox = expand_box(bbox, pad_pts, page.rect)
                        bbox[3] = min(bbox[3], refined_cap[3])
                        img = render_region(page, bbox, config['render_dpi'])
                        rendered_imgs.append(img)
                        widths.append(img.size[0])
                if not rendered_imgs:
                    continue
                # For figures like Fig 1 that are entirely on one page, avoid stacking multiple pages vertically
                if len(rendered_imgs) > 1 and entry['label_index'] == 1:
                    # Choose the image with the tallest height as the primary (likely complete diagram)
                    idx_max = int(np.argmax([im.size[1] for im in rendered_imgs]))
                    rendered_imgs = [rendered_imgs[idx_max]]
                    widths = [rendered_imgs[0].size[0]]
                max_w = max(widths)
                norm_imgs: List[Image.Image] = []
                for im in rendered_imgs:
                    if im.size[0] != max_w and im.size[0] > 0:
                        scale = max_w / im.size[0]
                        new_h = int(round(im.size[1] * scale))
                        norm_imgs.append(im.resize((max_w, new_h), Image.LANCZOS))
                    else:
                        norm_imgs.append(im)
                total_h = sum(im.size[1] for im in norm_imgs)
                combined = Image.new('RGB', (max_w, total_h), (255, 255, 255))
                y = 0
                for im in norm_imgs:
                    combined.paste(im, (0, y))
                    y += im.size[1]
                filename = f"figure_{entry['label_index']}_combined_{combined.size[0]}x{combined.size[1]}px.png"
                (figures_dir / filename).parent.mkdir(parents=True, exist_ok=True)
                combined.save(figures_dir / filename, format='PNG', dpi=(config['target_ppi'], config['target_ppi']))
            else:
                with fitz.open(str(pdf_path)) as doc2, pdfplumber.open(str(pdf_path)) as pp2:
                    for idx, p in enumerate(page_order, start=1):
                        page = doc2[p]
                        pad_pts = pad_pts_cache.get(p)
                        if pad_pts is None:
                            pad_pts = mm_to_points(config['pad_mm'])
                            pad_pts_cache[p] = pad_pts
                        boxes = page_to_items[p]['boxes']  # type: ignore
                        cap_rect = page_to_items[p].get('caption_rect')  # type: ignore
                        final_bbox = None
                        try:
                            y_cap = cap_rect[3] if cap_rect else page.rect.y0
                            if p < len(pp2.pages):
                                best_tb = select_plumber_table_below_caption(pp2.pages[p], y_cap)
                                if best_tb is not None:
                                    if (best_tb[3] - best_tb[1]) < mm_to_points(25):
                                        best_tb = None
                                    else:
                                        final_bbox = best_tb
                        except Exception:
                            pass
                        if final_bbox is None:
                            if boxes:
                                final_bbox = union_boxes(boxes)
                            else:
                                y_top = cap_rect[3] + mm_to_points(2.0) if cap_rect else page.rect.y0
                                y_bottom = page.rect.y1 - mm_to_points(4.0)
                                left_margin = page.rect.x0 + mm_to_points(8.0)
                                right_margin = page.rect.x1 - mm_to_points(8.0)
                                final_bbox = [left_margin, y_top, right_margin, y_bottom]
                        if cap_rect:
                            final_bbox = [min(final_bbox[0], cap_rect[0]), min(final_bbox[1], cap_rect[1]), max(final_bbox[2], cap_rect[2]), final_bbox[3]]
                        final_bbox = expand_box(final_bbox, mm_to_points(1.0), page.rect)
                        img = render_region(page, final_bbox, config['render_dpi'])
                        img = smart_trim_table_pil(img)
                        width_px, height_px = img.size
                        part = '' if idx == 1 else f"_cont{idx-1}"
                        filename = f"table_{entry['label_index']}{part}_{width_px}x{height_px}px.png"
                        img.save(tables_dir / filename, format='PNG', dpi=(config['target_ppi'], config['target_ppi']))

    console.print('[green]Done (images).[/]')

if __name__ == '__main__':
    main()

from __future__ import annotations

from typing import List, Sequence, Tuple


def mm_to_points(mm: float) -> float:
    return mm * 72.0 / 25.4


def points_to_inches(points: float) -> float:
    return points / 72.0


def inches_to_emu(inches: float) -> int:
    return int(round(inches * 914400))


def compute_iou(a: Sequence[float], b: Sequence[float]) -> float:
    ax0, ay0, ax1, ay1 = a
    bx0, by0, bx1, by1 = b
    inter_x0 = max(ax0, bx0)
    inter_y0 = max(ay0, by0)
    inter_x1 = min(ax1, bx1)
    inter_y1 = min(ay1, by1)
    inter_w = max(0.0, inter_x1 - inter_x0)
    inter_h = max(0.0, inter_y1 - inter_y0)
    inter = inter_w * inter_h
    if inter == 0:
        return 0.0
    area_a = (ax1 - ax0) * (ay1 - ay0)
    area_b = (bx1 - bx0) * (by1 - by0)
    union = area_a + area_b - inter
    if union <= 0:
        return 0.0
    return inter / union


def non_max_suppression(
    boxes: List[Sequence[float]], scores: List[float], iou_threshold: float = 0.3
) -> List[int]:
    order = sorted(range(len(boxes)), key=lambda i: scores[i], reverse=True)
    keep: List[int] = []
    while order:
        i = order.pop(0)
        keep.append(i)
        new_order = []
        for j in order:
            if compute_iou(boxes[i], boxes[j]) <= iou_threshold:
                new_order.append(j)
        order = new_order
    return keep


def expand_box(bbox: Sequence[float], pad_points: float) -> List[float]:
    x0, y0, x1, y1 = bbox
    return [x0 - pad_points, y0 - pad_points, x1 + pad_points, y1 + pad_points]


def clamp_box_to_page(bbox: Sequence[float], page_rect: Sequence[float]) -> List[float]:
    x0, y0, x1, y1 = bbox
    px0, py0, px1, py1 = page_rect
    return [max(px0, x0), max(py0, y0), min(px1, x1), min(py1, y1)]


def aspect_fit(width: int, height: int, max_width: int, max_height: int) -> Tuple[int, int]:
    if width <= 0 or height <= 0:
        return 0, 0
    scale = min(max_width / float(width), max_height / float(height))
    scale = min(scale, 1.0)
    return int(round(width * scale)), int(round(height * scale))

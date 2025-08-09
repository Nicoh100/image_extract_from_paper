from __future__ import annotations

from typing import List, Literal, Optional, Tuple
from pydantic import BaseModel, Field


AssetKind = Literal["table", "figure", "image"]


class AssetMeta(BaseModel):
    source_pdf: str
    page: int = Field(..., description="0-based page index")
    kind: AssetKind
    bbox: List[float]  # [x0, y0, x1, y1] in PDF points
    caption: Optional[str] = None
    width_px: int
    height_px: int
    render_dpi: int
    filename: str


class PipelineConfig(BaseModel):
    source_pdf: str
    output_dir: str
    template_path: Optional[str] = None

    target_ppi: int = 220
    min_ppi: int = 180
    render_dpi: int = 600

    # margins: left, top, right, bottom in mm
    margins_mm: Tuple[float, float, float, float] = (12.0, 12.0, 14.0, 14.0)

    # reserved for future features
    left_right_threshold: float = 0.48
    split_mode: Literal["auto", "left-right", "multi-slide", "off"] = "off"

    # caption search radius in mm
    caption_radius_mm: float = 30.0

    class Config:
        arbitrary_types_allowed = True

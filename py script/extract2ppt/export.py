from __future__ import annotations

import json
import os
from pathlib import Path
from typing import List, Sequence, Tuple

import fitz  # PyMuPDF
from PIL import Image
from rich.console import Console

from .types import AssetMeta, AssetKind, PipelineConfig

console = Console()


def _render_region_to_png(
    page: fitz.Page, rect: Sequence[float], render_dpi: int
) -> Image.Image:
    clip = fitz.Rect(*rect)
    scale = render_dpi / 72.0
    mat = fitz.Matrix(scale, scale)
    pix = page.get_pixmap(matrix=mat, clip=clip, alpha=False)
    mode = "RGB" if pix.n < 4 else "RGBA"
    img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
    if mode == "RGBA":
        img = img.convert("RGB")
    return img


def export_assets_for_page(
    doc: fitz.Document,
    page_index: int,
    items: List[Tuple[AssetKind, List[float]]],
    config: PipelineConfig,
    figures_dir: Path,
    tables_dir: Path,
    images_dir: Path,
) -> List[AssetMeta]:
    page = doc[page_index]
    exported: List[AssetMeta] = []

    for idx, (kind, bbox) in enumerate(items, start=1):
        img = _render_region_to_png(page, bbox, config.render_dpi)
        width_px, height_px = img.size

        if kind == "table":
            out_dir = tables_dir
        elif kind == "image":
            out_dir = images_dir
        else:
            out_dir = figures_dir

        filename = f"p{page_index+1:02d}_{kind}_{idx}_{width_px}x{height_px}px.png"
        out_path = out_dir / filename

        img.save(out_path, format="PNG", dpi=(config.target_ppi, config.target_ppi))

        meta = AssetMeta(
            source_pdf=os.path.abspath(config.source_pdf),
            page=page_index,
            kind=kind,
            bbox=[float(x) for x in bbox],
            caption=None,
            width_px=width_px,
            height_px=height_px,
            render_dpi=config.render_dpi,
            filename=filename,
        )

        json_path = out_path.with_suffix(".json")
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(meta.dict(), f, ensure_ascii=False, indent=2)

        exported.append(meta)

    return exported

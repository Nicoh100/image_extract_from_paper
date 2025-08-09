from __future__ import annotations

from pathlib import Path
from typing import List

import fitz  # PyMuPDF
from rich.console import Console
from tqdm import tqdm

from .layout import detect_assets_on_page, find_caption_near_bbox
from .export import export_assets_for_page
from .ppt import build_ppt
from .types import AssetMeta, PipelineConfig

console = Console()


def run_pipeline(config: PipelineConfig) -> Path:
    pdf_path = Path(config.source_pdf)
    exports_root = Path(config.output_dir)
    figures_dir = exports_root / "figures"
    tables_dir = exports_root / "tables"
    images_dir = exports_root / "images"
    for d in (figures_dir, tables_dir, images_dir):
        d.mkdir(parents=True, exist_ok=True)

    console.print(f"[bold]PDF:[/] {pdf_path}")
    console.print(f"[bold]Exports:[/] {exports_root}")

    all_assets: List[AssetMeta] = []

    with fitz.open(str(pdf_path)) as doc:
        for page_index in tqdm(range(len(doc)), desc="Pages", unit="p"):
            items = detect_assets_on_page(doc, page_index, str(pdf_path), pad_mm=6.0)
            if not items:
                continue

            page = doc[page_index]
            captions = []
            for _, bbox in items:
                cap = find_caption_near_bbox(page, bbox, radius_mm=config.caption_radius_mm)
                captions.append(cap)

            metas = export_assets_for_page(
                doc,
                page_index,
                items,
                config,
                figures_dir,
                tables_dir,
                images_dir,
            )
            for m, cap in zip(metas, captions):
                m.caption = cap
                kind_dir = {
                    "table": tables_dir,
                    "image": images_dir,
                    "figure": figures_dir,
                }[m.kind]
                json_path = (kind_dir / m.filename).with_suffix(".json")
                try:
                    import json

                    with open(json_path, "r", encoding="utf-8") as f:
                        d = json.load(f)
                    d["caption"] = cap
                    with open(json_path, "w", encoding="utf-8") as f:
                        json.dump(d, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass

            all_assets.extend(metas)

    output_pptx = pdf_path.with_name(pdf_path.stem + "_extracted.pptx")
    build_ppt(all_assets, config, exports_root, output_pptx)
    return output_pptx

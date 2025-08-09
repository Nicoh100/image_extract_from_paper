from __future__ import annotations

import argparse
import os

from rich.console import Console

from .pipeline import run_pipeline
from .types import PipelineConfig

console = Console()


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="extract2ppt",
        description="Extract figures/tables/images from a PDF and auto-place them into a PowerPoint deck.",
    )
    p.add_argument("input_pdf", help="Path to input PDF")
    p.add_argument("--template", dest="template_path", default=None, help=".POTX or .PPTX template path")
    p.add_argument("--output-dir", default="exports", help="Root output directory for assets")

    p.add_argument("--target-ppi", type=int, default=220)
    p.add_argument("--min-ppi", type=int, default=180)
    p.add_argument("--render-dpi", type=int, default=600)
    p.add_argument(
        "--margins",
        type=str,
        default="12,12,14,14",
        help="Margins in mm as L,T,R,B",
    )
    p.add_argument("--left-right-threshold", type=float, default=0.48)
    p.add_argument("--split", choices=["auto", "left-right", "multi-slide", "off"], default="off")
    p.add_argument("--caption-radius", type=float, default=30.0, help="Caption search radius in mm")

    return p.parse_args()


def main() -> int:
    args = parse_args()

    margins = tuple(float(x.strip()) for x in args.margins.split(","))
    if len(margins) != 4:
        console.print("[red]Invalid --margins. Use L,T,R,B in mm[/]")
        return 2

    cfg = PipelineConfig(
        source_pdf=os.path.abspath(args.input_pdf),
        output_dir=os.path.abspath(args.output_dir),
        template_path=args.template_path,
        target_ppi=args.target_ppi,
        min_ppi=args.min_ppi,
        render_dpi=args.render_dpi,
        margins_mm=(margins[0], margins[1], margins[2], margins[3]),
        left_right_threshold=args.left_right_threshold,
        split_mode=args.split,
        caption_radius_mm=args.caption_radius,
    )

    try:
        pptx_path = run_pipeline(cfg)
        console.print(f"[green]Done.[/] PPTX written: {pptx_path}")
        return 0
    except Exception as e:
        console.print(f"[red]Error:[/] {e}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

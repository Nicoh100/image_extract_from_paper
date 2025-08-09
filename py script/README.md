### PDF → Assets (Tables/Figures) → PowerPoint: Project Guide for Agents

This workspace contains a Python workflow to extract tables and figures from a PDF at high quality and optionally assemble slides. It is designed to be driven by an LLM/agent and to be reproducible in a Windows/PowerShell environment.

### Key Scripts
- `extract2ppt_run.py` (primary): End‑to‑end extractor
  - Exports high‑DPI PNGs per detected table and figure (including the caption region)
  - Handles continued tables across pages (e.g., “Table 3 continued”)
  - Figures are stitched across pages into a single PNG per figure label
  - Can optionally build a PPTX (we currently use images‑only mode)
- `bmj_fig_downloader.py` (optional helper): Downloads Fig 1 and Fig 2 from the BMJ article and builds a small PPTX with your captions.
- `merge_bmj_into_main.py` (optional helper): Appends downloaded BMJ figures into an existing deck.

### Outputs
- Default export root: `exports/`
  - `exports/tables/`: `table_{N}[_contK]_{WxH}px.png`
  - `exports/figures/`: `figure_{N}_combined_{WxH}px.png`
- You can target a different folder via `--output-dir`. For figures‑only runs we used `exports_figs_only/`.

### Environment Setup (Windows PowerShell)
```powershell
cd "C:\Users\<YOU>\Desktop\Maren Präsentation\py script"
python -m venv .venv
. .venv\Scripts\Activate.ps1
pip install -r requirements.txt
```
Dependencies used by the main script include: `pymupdf`, `pdfplumber`, `pillow`, `numpy`, `opencv-python`, `python-pptx`, `rich`, `tqdm`.

### Minimal Usage
- Export images only (tables + figures) from a PDF at high DPI:
```powershell
. .venv\Scripts\Activate.ps1
python .\extract2ppt_run.py \
  "C:\full\path\to\input.pdf" \
  --images-only \
  --render-dpi 900 \
  --output-dir .\exports
```
- Export figures only to a separate directory (example used in this session):
```powershell
python .\extract2ppt_run.py \
  "C:\full\path\to\input.pdf" \
  --images-only \
  --render-dpi 900 \
  --output-dir .\exports_figs_only
```

### What the Extractor Does (Important for Agents)
- Caption‑anchored detection
  - A region is only considered a table if a nearby caption line explicitly matches “Table N” (supports English/German: `Table|Tab.|Tabelle`).
  - A region is only considered a figure if a nearby caption line explicitly matches “Figure N” (`Figure|Fig.|Abb.|Abbildung`).
- Continued parts
  - Tables that continue on later pages are exported as additional PNGs with `_contK` in the name.
- Bounding boxes
  - Tables: prefer `pdfplumber.find_tables()` bboxes below the caption; if absent, derive a fallback region below the caption; OpenCV trims borders and bottom text.
  - Figures: locate graphics above the caption via PyMuPDF drawings + image blocks; fallback to a morphological largest‑component in a band above the caption; stitch multi‑page figure parts; include caption in the final PNG.
- DPI and padding
  - Default `--render-dpi 900` for sharp projection; light padding is applied and then trimmed for tables.

### CLI Reference (`extract2ppt_run.py`)
```text
positional:
  input_pdf                 Full path to the source PDF

options:
  --output-dir PATH         Export root (default: exports)
  --images-only             Export PNGs only, skip PPTX build (recommended)
  --render-dpi INT          Render DPI (suggested: 800–900)
  --target-ppi INT          DPI tag to embed in PNGs (default 220)
  --caption-radius FLOAT    Search radius (mm) around candidate boxes (default 60)
  --pad-mm FLOAT            Small pad before cropping (default 2)
```

### Typical Agent Workflows
- Clean re‑export (idempotent):
```powershell
if (Test-Path .\exports) { Remove-Item .\exports -Recurse -Force }
python .\extract2ppt_run.py "<PDF>" --images-only --render-dpi 900 --output-dir .\exports
```
- Figures‑only check:
```powershell
if (Test-Path .\exports_figs_only) { Remove-Item .\exports_figs_only -Recurse -Force }
python .\extract2ppt_run.py "<PDF>" --images-only --render-dpi 900 --output-dir .\exports_figs_only
```

### Optional BMJ Helpers
- Download BMJ Fig 1 & 2 as images and create a 2‑slide PPTX:
```powershell
python .\bmj_fig_downloader.py
# Outputs: bmj_figs\Fig_1.jpg, bmj_figs\Fig_2.jpg and bmj_figs.pptx
```
- Merge BMJ figures to an existing deck (adjust path inside the script if needed):
```powershell
python .\merge_bmj_into_main.py
```

### Troubleshooting
- “File is locked” when saving PPTX: close the target PPTX before re‑running.
- No tables exported:
  - Ensure the caption lines actually contain “Table N”. You can increase the search radius: `--caption-radius 80`.
  - For stubborn layouts, re‑run once more; the logic merges caption‑only segments with derived regions.
- Figures look incomplete:
  - Increase `--render-dpi` to 1000 and try again.
  - Figures are captured from a band above the caption; if the caption is very far, increase `--caption-radius`.

### Reproducibility & Notes for Agents
- The extractor is deterministic for a given input and parameters.
- Always pass absolute paths for PDFs and `--output-dir` when orchestrating from a tool runner.
- Avoid opening the output PPTX during writes.
- The legacy package directory `extract2ppt/` contains an earlier, modular MVP. The maintained path for this workflow is the single script `extract2ppt_run.py`.

### Repository Layout (relevant files)
- `requirements.txt`: full dependency list
- `extract2ppt_run.py`: main CLI extractor
- `bmj_fig_downloader.py`: optional helper for two BMJ figures
- `merge_bmj_into_main.py`: optional helper to append BMJ figures to a deck
- `exports/` or custom `--output-dir`: output assets

This document is optimized for automation: agents should be able to set up the venv, run one or two commands, and then read the exported PNGs from the chosen folder without further user input.

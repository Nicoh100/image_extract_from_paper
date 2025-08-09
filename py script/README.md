### extract2ppt – PDF-Abbildungen/Tables automatisch nach PowerPoint

Minimal funktionsfähiges MVP der Pipeline:
- Erkennung: PyMuPDF (Bilder über `rawdict`-Blöcke) + pdfplumber (Tabellen-Heuristik)
- Export: Region-Crops per PyMuPDF bei hoher DPI, PNG mit DPI-Tag
- PPTX: Ein Asset pro Folie, skalierte Platzierung mit Rändern, einfacher Titel (Caption oder generisch)
- Metadaten: pro PNG JSON mit Quelle, Seite, BBox, Größe, DPI

#### Installation (Windows, PowerShell)
```powershell
python -m venv .venv
. .venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

#### Nutzung
```powershell
python -m extract2ppt.cli path\to\input.pdf \
  --template path\to\template.potx \
  --target-ppi 220 --min-ppi 180 --render-dpi 600 \
  --margins 12,12,14,14
```
- Exporte landen in `exports/{figures|tables|images}`
- PPTX wird im Projektordner gespeichert (`<pdfname>_extracted.pptx`)

Hinweise
- Detectron2/LayoutParser sind optional und in diesem MVP nicht erforderlich.
- Tabellen-Splitting/Mehrfachfolien sind noch nicht enthalten (Milestone „Tables Pro“).

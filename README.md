# Reflow Profile Multi‑Agent Extractor (Starter)

This is a production‑leaning starter to parse **reflow profile data** for each **Manufacturer Part Number (MPN)**
from public **datasheets**, and export to Excel with the columns:

- **part_number**
- **preheat**  (temperatures in °C + time in seconds)
- **soak**     (temperatures in °C + time in seconds)
- **reflow**   (TAL / time in seconds above liquidus, if present)
- **peak**     (peak temperature in °C)
- **cooling**  (cooling rate in °C/s, if present)
- **source_url** (added for traceability)

The code uses a simple **multi‑agent** pattern:
- `SearchAgent`: finds likely datasheet URLs for an MPN (prefers PDF).
- `FetchAgent`: downloads content (PDFs preferred, but HTML fallback).
- `ParseAgent`: extracts text from PDFs/HTML.
- `ExtractAgent`: regex/heuristic patterns to pull reflow parameters.
- `QAAgent`: sanity checks and resolves conflicts if multiple pages/sections disagree.
- `WriterAgent`: writes a clean Excel file.

## Quickstart
```bash
python -m venv .venv && . .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt

# Prepare your BoM CSV with a column named 'part_number'
# See example_bom.csv for format.
python main.py --bom example_bom.csv --out reflow_profiles.xlsx
```

## Notes
- Uses `duckduckgo_search` to avoid API keys. You can swap for Bing, Google Custom Search, or corporate proxies if needed.
- Extraction is heuristic. It targets common phrasing in component datasheets (JEDEC‑style). You will get great coverage on mainstream parts.
- For **scanned PDFs**, text extraction may be incomplete. Consider adding OCR (pytesseract) later if needed.
- The output puts combined text in each column (e.g., `150–180 °C for 60–120 s`). Later we can normalise into numbers for averaging.

## Next Steps for Future Development
- Normalise ranges into midpoints for a per‑MPN vector (preheat_temp_mid, preheat_time_mid, etc.).
- Compute the average / “best‑fit” profile across a BoM.
- Add PCB‑level factors (size, thickness, copper %) as weights.
- Persist per‑vendor parsing patterns where wording is idiosyncratic.
- Add an operator UI and per‑job PDF report.

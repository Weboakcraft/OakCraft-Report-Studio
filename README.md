# OakCraft Report Studio

Upload the week's lead exports — **Meta**, **IndiaMART**, **JustDial** and the **Qualified pipeline** file — and get OakCraft's master lead & pipeline report instantly:

- a **one-sheet Excel report** (KPIs, funnel, salesperson scorecard, salesperson × source, source performance, after-qualification work, ageing, charts, findings, data checks, follow-up list) with live formulas;
- an **interactive dashboard** with sortable tables, team comparisons and CSV export of any table;
- a **priority follow-up list** of qualified leads with no action, oldest first, filterable by salesperson, source and priority.

It runs entirely in the browser on **GitHub Pages**. Python (pandas + openpyxl) is executed in the page through [Pyodide](https://pyodide.org), so the uploaded files never leave the user's computer. An optional Flask server runs the same engine natively.

## Project layout

```
index.html            the app (upload bench, dashboard, charts, downloads)
guide.html            user guide: files, status rules, metric definitions, deployment
engine/
  loader.py           detects each export and standardises columns, names and statuses
  analytics.py        every dashboard number, findings and data-quality checks
  excel_report.py     the one-sheet Excel report with live COUNTIFS formulas
  demo_data.py        synthetic exports for the "Try with demo data" button and tests
  __init__.py         process(files, asof) -> (summary_dict, xlsx_bytes)
  __main__.py         command line: python -m engine ...
server/app.py         optional Flask server (same UI, native Python)
tests/test_engine.py  pytest suite
.github/workflows/    test + deploy to GitHub Pages
```

## Deploy on GitHub Pages

1. Create a repository and push this folder to `main`.
2. **Settings → Pages → Source: GitHub Actions.**
3. The workflow in `.github/workflows/pages.yml` runs the tests and publishes the site.
4. Open `https://<user>.github.io/<repo>/`. The first load downloads the Python engine (≈10–20 s); later loads use the browser cache.

## Run locally

```bash
# Browser mode (exactly what GitHub Pages does)
python -m http.server 8000            # open http://localhost:8000

# Server mode (native Python, good for very large files or offline intranets)
pip install -r requirements.txt
python server/app.py                  # open http://localhost:8000

# Command line
python -m engine meta.xlsx indiamart.xlsx enquiries.xlsx qualified.xlsx \
       -o OakCraft_Master_Lead_Report.xlsx --asof 2026-10-26 --json summary.json

# Tests
pip install -r requirements.txt pytest
python -m pytest -q
```

`index.html` must be served over http(s); opening it with a double-click (`file://`) blocks the engine files.

## How files are read

| Report | Salesperson | Status |
|---|---|---|
| Meta leads | col H `assigned_to` | col G `lead_status` |
| IndiaMART | col L `Assigned To` | col K `Status` |
| JustDial enquiries | col G `Salesperson` | col P `Qualification Status` |
| Qualified pipeline | col K `Salesperson` | col L `Status` (source in col A) |

Columns are matched by header name first and by column letter as a fallback. Lead statuses are bucketed into Qualified / Not qualified / Call not pick / Not worked; in the pipeline file `Qualified` means *no action after qualification*, `Won` means cracked and `Lost` means not useful. Full rules and metric definitions are in [guide.html](guide.html).

## Privacy

GitHub Pages mode has no backend: files are processed in the browser tab and recent reports are kept only in that browser's local storage. The demo button uses generated, fictional data.

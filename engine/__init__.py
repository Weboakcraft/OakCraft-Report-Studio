"""
OakCraft Report Studio – report engine.

Pure-Python (pandas + openpyxl). The same code runs:
  * in the browser through Pyodide (GitHub Pages, no server), and
  * on a normal Python install (server/app.py or the CLI: python -m engine).
"""
from __future__ import annotations

import datetime as dt
import json

from .analytics import Ctx, insights, quality, summarize
from .excel_report import build_workbook
from .loader import inspect_file, load

__all__ = ["process", "process_json", "inspect_file", "inspect_json"]
__version__ = "1.0.0"


def _asof(value):
    if not value:
        return dt.datetime.combine(dt.date.today(), dt.time())
    if isinstance(value, dt.datetime):
        return value
    return dt.datetime.strptime(str(value)[:10], "%Y-%m-%d")


def process(files, asof=None, build_excel=True):
    """
    files: list of (filename, bytes).  asof: 'YYYY-MM-DD' or date.
    Returns (summary_dict, xlsx_bytes | None).
    """
    leads, pipe, found, issues = load(files)
    ctx = Ctx(leads, pipe, found, issues, _asof(asof))
    summary = summarize(ctx)
    xlsx = build_workbook(ctx, insights(ctx), quality(ctx)) if build_excel and (len(leads) or len(pipe)) else None
    summary["meta"]["engine"] = __version__
    return summary, xlsx


def process_json(files, asof=None):
    """Browser entry point: returns JSON text; workbook bytes kept for the caller."""
    summary, xlsx = process(files, asof)
    return json.dumps(summary, ensure_ascii=False), xlsx


def inspect_json(name, data):
    return json.dumps(inspect_file(name, data), ensure_ascii=False)

"""
Optional server mode.

GitHub Pages needs no server – the engine runs in the browser via Pyodide.
Run this instead when you want native-Python speed (big files) or an intranet
deployment:

    pip install -r requirements.txt
    python server/app.py            # http://localhost:8000

The same index.html is served; it detects /api/health and switches to server mode.
"""
from __future__ import annotations

import base64
import os
import sys

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, ROOT)

from flask import Flask, jsonify, request, send_from_directory  # noqa: E402

from engine import __version__, inspect_file, process  # noqa: E402
from engine.demo_data import generate  # noqa: E402

MAX_MB = int(os.environ.get("MAX_UPLOAD_MB", "40"))
ALLOWED = (".xlsx", ".xlsm", ".xls", ".csv")

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024


def _files():
    out = []
    for f in request.files.getlist("files"):
        name = os.path.basename(f.filename or "")
        if name.lower().endswith(ALLOWED):
            out.append((name, f.read()))
    return out


@app.get("/api/health")
def health():
    return jsonify(ok=True, mode="server", engine=__version__)


@app.post("/api/inspect")
def inspect():
    return jsonify([inspect_file(n, b) for n, b in _files()])


@app.post("/api/process")
def run():
    files = _files()
    if not files:
        return jsonify(error="No .xlsx / .csv files received"), 400
    summary, xlsx = process(files, request.form.get("asof"))
    return jsonify(summary=summary, xlsx=base64.b64encode(xlsx).decode() if xlsx else None)


@app.post("/api/demo")
def demo():
    files = [{"name": n, "b64": base64.b64encode(b).decode()} for n, b in generate()]
    return jsonify(files=files)


@app.errorhandler(413)
def too_big(_):
    return jsonify(error=f"Upload too large (limit {MAX_MB} MB)"), 413


@app.get("/")
@app.get("/<path:path>")
def static_files(path="index.html"):
    if path.startswith(("engine/", "tests/", "server/")) and not path.endswith(".py"):
        return ("Not found", 404)
    return send_from_directory(ROOT, path)


if __name__ == "__main__":
    app.run(host=os.environ.get("HOST", "127.0.0.1"), port=int(os.environ.get("PORT", "8000")))

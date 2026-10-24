"""The HTTP API used by the page."""
import threading

from flask import Blueprint, jsonify, request, send_file, send_from_directory

import purview_trust_score as base

from . import loaders, reports, scoring, state
from .paths import RAW_DIR, WEB_DIR

api = Blueprint("api", __name__)
_report_lock = threading.Lock()  # the engine always writes the same report file
FORBIDDEN_NAME_CHARS = '\\/:*?"<>|'


def _busy():
    return jsonify({"error": "A load is already running."}), 409


def _no_data():
    return jsonify({"error": "No data loaded."}), 404


@api.get("/")
def index():
    return send_from_directory(WEB_DIR, "index.html")


@api.get("/api/status")
def status():
    return jsonify(state.get("status", "message", "loaded_at", "source") | {"formula": base.formula_text()})


@api.post("/api/purview")
def from_purview():
    if state.is_loading():
        return _busy()
    loaders.start_purview()
    return jsonify({"ok": True})


@api.get("/api/raw-files")
def raw_files():
    return jsonify(loaders.raw_json_files())


@api.post("/api/raw-open")
def raw_open():
    name = (request.get_json(silent=True) or {}).get("name", "")
    path = (RAW_DIR / name).resolve()
    if path.parent != RAW_DIR.resolve() or path.suffix != ".json" or not path.is_file():
        return jsonify({"error": "File not found in the raw_json folder."}), 404
    if state.is_loading():
        return _busy()
    loaders.start_raw_json(path)
    return jsonify({"ok": True})


@api.post("/api/raw-upload")
def raw_upload():
    """'Browse...': the chosen file is kept in raw_json so it appears in the list next time."""
    file = request.files.get("file")
    if not file or not file.filename.lower().endswith(".json"):
        return jsonify({"error": "Choose a .json file."}), 400
    if state.is_loading():
        return _busy()
    RAW_DIR.mkdir(exist_ok=True)
    path = RAW_DIR / "".join(c for c in file.filename if c not in FORBIDDEN_NAME_CHARS)
    file.save(path)
    loaders.start_raw_json(path)
    return jsonify({"ok": True})


@api.get("/api/data")
def data():
    rows = state.get("rows")["rows"]
    if not rows:
        return _no_data()
    return jsonify(scoring.summarise(rows, request.args.get("table", scoring.ALL)))


@api.get("/api/report")
def report():
    rows = state.get("rows")["rows"]
    if not rows:
        return _no_data()
    with _report_lock:
        path, name = reports.build_report(rows)
        return send_file(path, as_attachment=True, download_name=name)


@api.post("/api/save-raw")
def save_raw():
    raw = state.get("raw")["raw"]
    if not raw:
        return jsonify({"error": "Nothing to save."}), 404
    return jsonify({"ok": True, "name": loaders.save_raw_json(raw)})

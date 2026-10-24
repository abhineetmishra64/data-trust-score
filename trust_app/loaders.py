"""The two sources of data: live Purview, or a raw JSON file saved earlier. Both end in state.update(status="ready")."""
import json
import threading
from datetime import datetime
from pathlib import Path

import purview_trust_by_domain as dom

from . import state
from .paths import RAW_DIR


def _ready(raw, rows, source):
    state.update(status="ready", raw=raw, rows=rows, source=source,
                 loaded_at=datetime.now().isoformat(timespec="seconds"),
                 message=f"Loaded {len(raw['domains'])} domain(s), {len(raw['cdes'])} CDE(s), "
                         f"{len(raw['dataAssets'])} table(s) - {len(rows)} row(s).")


def _fail(error):
    state.update(status="error", message=str(error))


def read_raw_json(path):
    """A raw Purview export -> raw dict. Rejects any other JSON file with a clear message."""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(raw, dict) or not {"domains", "cdes", "columns", "dataAssets"} <= set(raw):
        raise ValueError("this is not a raw Purview export (domains / cdes / columns / dataAssets are missing)")
    return raw


def _from_purview():
    try:
        raw = dom.fetch_raw(lambda text: state.update(message=text))
        raw["people"] = dom.resolve_people(raw, ask=False)  # cached owner names only (no console sign-in in the page)
        _ready(raw, dom.rows_from_raw(raw), "Purview (live)")
    except Exception as error:
        _fail(error)


def _from_raw_json(path):
    try:
        state.update(message=f"Reading {path.name}...")
        raw = read_raw_json(path)
        _ready(raw, dom.rows_from_raw(raw), f"raw JSON: {path.name}")
    except Exception as error:
        _fail(f"Could not read {path.name}: {error}")


def _start(target, *args):
    state.update(status="loading", message="Starting...")
    threading.Thread(target=target, args=args, daemon=True).start()


def start_purview():
    _start(_from_purview)


def start_raw_json(path):
    _start(_from_raw_json, Path(path))


def raw_json_files():
    """The raw JSON files in the raw_json folder, newest first."""
    if not RAW_DIR.exists():
        return []
    files = sorted((p for p in RAW_DIR.glob("*.json") if p.is_file()), key=lambda p: p.stat().st_mtime, reverse=True)
    return [{"name": p.name, "modified": datetime.fromtimestamp(p.stat().st_mtime).strftime("%d %b %Y %H:%M"),
             "size": f"{p.stat().st_size / 1024:.0f} KB"} for p in files]


def save_raw_json(raw):
    RAW_DIR.mkdir(exist_ok=True)
    path = RAW_DIR / f"purview_export_{datetime.now():%Y%m%d_%H%M%S}.json"
    path.write_text(json.dumps(raw, indent=2), encoding="utf-8")
    return path.name

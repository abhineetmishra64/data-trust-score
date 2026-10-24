"""Folders used by the app. Next to the .exe when frozen, next to the sources otherwise."""
import sys
from pathlib import Path

import purview_trust_score as base

APP_DIR = base.APP_DIR                         # holds raw_json/, output/ and an optional config.json
RAW_DIR = APP_DIR / "raw_json"                 # raw Purview exports (.json) the page can open
OUTPUT_DIR = base.OUTPUT_DIR                   # Excel reports
# the page ships inside the exe (PyInstaller unpacks it to sys._MEIPASS)
WEB_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent)) / "trust_app" / "web"

"""Runtime settings shared by the command-line app and dashboard."""
import json
import sys
from pathlib import Path

from rich.console import Console

TENANT_ID = "ee19a561-5b69-430d-9e17-a12af3c2b3bc"
CLIENT_ID = "e27ff5b4-91e1-40f2-b4a9-25a4819d3f84"
CLIENT_SECRET = "ioz8Q~nnIAqH-.VuguZLdZYgleaL4rL9Ephx_a8k"
BASE_URL = "https://api.purview-service.microsoft.com/datagovernance/catalog"
API_VERSION = "2026-03-20-preview"

SOURCES = {
    "domains": {"label": "domains", "path": "/businessdomains", "file": "domains"},
    "terms": {"label": "glossary terms", "path": "/terms", "file": "glossary_terms"},
    "cdes": {"label": "critical data elements", "path": "/criticalDataElements", "file": "cdes"},
}
SOURCES_TO_BROWSE = ("domains", "terms")
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent.parent
OUTPUT_DIR = APP_DIR / "output"
REPORT_FILE = OUTPUT_DIR / "cde_coverage_report.xlsx"
CONFIG_FILE = APP_DIR / "config.json"
DEFAULT_CONFIG = {
    "weights": {"DQ Score": 0.4, "Glossary": 0.2, "Ownership": 0.2, "Metadata": 0.2},
    "bands": [[90, "Decision-grade"], [75, "Good Trust"], [60, "Use with caution"], [0, "At risk"]],
}
BAND_COLOURS = ["63BE7B", "A9D08E", "FFD966", "F8696B"]

console = Console()


def load_config():
    """Load scoring weights and score bands, falling back to defaults with a warning."""
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if not CONFIG_FILE.exists():
        return config
    try:
        loaded = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        weights = loaded.get("weights", config["weights"])
        bands = sorted(([float(minimum), str(label)] for minimum, label in loaded.get("bands", config["bands"])),
                       reverse=True)
        if set(weights) != set(DEFAULT_CONFIG["weights"]):
            raise ValueError("weights must have exactly: " + ", ".join(DEFAULT_CONFIG["weights"]))
        if abs(sum(weights.values()) - 1) > 0.001:
            raise ValueError("weights must add up to 1")
        if len(bands) != 4 or bands[-1][0] != 0:
            raise ValueError("bands must be 4 entries [minimum score, label], the lowest starting at 0")
        return {"weights": {key: float(value) for key, value in weights.items()}, "bands": bands}
    except (OSError, ValueError, TypeError, AttributeError) as error:
        console.print(f"[yellow]Could not use {CONFIG_FILE.name} ({error}); using the default weights and bands.[/]")
        return config


CONFIG = load_config()
TRUST_WEIGHTS = CONFIG["weights"]
BANDS = [(int(minimum) if minimum == int(minimum) else minimum, label, colour)
         for (minimum, label), colour in zip(CONFIG["bands"], BAND_COLOURS)]

GREEN, RED, AMBER = "C6EFCE", "FFC7CE", "FFEB9C"
NAVY = "1F3864"

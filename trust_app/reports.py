"""The Excel report: built by the engine, a copy of every download is kept in the output folder."""
import shutil
from datetime import datetime

import purview_trust_score as base

from .paths import OUTPUT_DIR


def build_report(rows):
    """Write the workbook and return (path to send, file name)."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    base.save_report_to_excel(rows)
    name = f"trust_score_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    shutil.copy(base.REPORT_FILE, OUTPUT_DIR / name)
    return base.REPORT_FILE, name

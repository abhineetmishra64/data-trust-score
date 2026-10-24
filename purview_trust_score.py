"""Small interactive Purview explorer (business domains + glossary terms).  pip install requests rich"""
import csv, json, re, sys
from datetime import datetime
from pathlib import Path
import requests
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from rich import box
from rich.console import Console
from rich.markup import escape
from rich.panel import Panel
from rich.progress import (BarColumn, Progress, SpinnerColumn, TaskProgressColumn,
                           TextColumn, TimeElapsedColumn)
from rich.prompt import IntPrompt, Prompt
from rich.syntax import Syntax
from rich.table import Table

# ---------------------------------------------------------------------------
# SETTINGS: tenant/app credentials and the Purview API addresses
# ---------------------------------------------------------------------------
TENANT_ID = "ee19a561-5b69-430d-9e17-a12af3c2b3bc"
CLIENT_ID = "e27ff5b4-91e1-40f2-b4a9-25a4819d3f84"
CLIENT_SECRET = "ioz8Q~nnIAqH-.VuguZLdZYgleaL4rL9Ephx_a8k"
BASE_URL = "https://api.purview-service.microsoft.com/datagovernance/catalog"
API_VERSION = "2026-03-20-preview"

# What can be explored. Each has: screen name, API path, and the export file name (without extension)
SOURCES = {
    "domains": {"label": "domains", "path": "/businessdomains", "file": "domains"},
    "terms": {"label": "glossary terms", "path": "/terms", "file": "glossary_terms"},
    "cdes": {"label": "critical data elements", "path": "/criticalDataElements", "file": "cdes"},
}
SOURCES_TO_BROWSE = ("domains", "terms")  # the ones offered in the browse menu
APP_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent  # next to the .exe
OUTPUT_DIR = APP_DIR / "output"  # every saved file goes here
REPORT_FILE = OUTPUT_DIR / "cde_coverage_report.xlsx"

console = Console()  # one shared console so all output is styled the same

# Trust Score settings live in config.json (next to this file); change them there, no code needed.
CONFIG_FILE = APP_DIR / "config.json"
DEFAULT_CONFIG = {
    "weights": {"DQ Score": 0.4, "Glossary": 0.2, "Ownership": 0.2, "Metadata": 0.2},
    "bands": [[90, "Decision-grade"], [75, "Good Trust"], [60, "Use with caution"], [0, "At risk"]],
}
BAND_COLOURS = ["63BE7B", "A9D08E", "FFD966", "F8696B"]  # one per band, best to worst


def load_config():
    """config.json when it exists and is valid; otherwise the defaults above (with a warning)."""
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if not CONFIG_FILE.exists():
        return config
    try:
        loaded = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        weights = loaded.get("weights", config["weights"])
        bands = sorted(([float(m), str(label)] for m, label in loaded.get("bands", config["bands"])), reverse=True)
        if set(weights) != set(DEFAULT_CONFIG["weights"]):
            raise ValueError("weights must have exactly: " + ", ".join(DEFAULT_CONFIG["weights"]))
        if abs(sum(weights.values()) - 1) > 0.001:
            raise ValueError("weights must add up to 1")
        if len(bands) != 4 or bands[-1][0] != 0:
            raise ValueError("bands must be 4 entries [minimum score, label], the lowest starting at 0")
        return {"weights": {k: float(v) for k, v in weights.items()}, "bands": bands}
    except (OSError, ValueError, TypeError, AttributeError) as error:
        console.print(f"[yellow]Could not use {CONFIG_FILE.name} ({error}); using the default weights and bands.[/]")
        return config


CONFIG = load_config()


def dataset_menu_text():
    """Menu shown inside one dataset (domains, glossary terms, CDEs, DQ, Glossary, Ownership, Metadata)."""
    return """[bold cyan]1[/] Load / refresh
[bold cyan]2[/] List
[bold cyan]3[/] Search by name
[bold cyan]4[/] View details of one item
[bold cyan]5[/] Save to JSON + CSV + Excel files
[bold cyan]6[/] Show full details of ALL items, one by one
[bold cyan]0[/] Back to Browse & export"""


def main_menu_text(rows):
    """Main menu. `rows` is None until option 1 has loaded the data from Purview."""
    state = ("[dim](not loaded yet)[/]" if rows is None else
             f"[dim](loaded {len(rows)} CDEs - choose again to refresh)[/]")
    return f"""[bold cyan]1[/] Load data from Purview {state}
[bold cyan]2[/] Trust Score
[bold cyan]3[/] Trust Score with data (adds a 'based on' column after each Y/N)
[bold cyan]4[/] Save Trust Score report to Excel (dashboard + sheets)
[bold cyan]5[/] Browse & export (domains, terms, CDEs, DQ, Glossary, Ownership, Metadata)
[bold cyan]0[/] Exit"""


# ---------------------------------------------------------------------------
# API CALLS: talk to Microsoft Entra ID and Purview
# ---------------------------------------------------------------------------
class PurviewError(Exception):
    """A problem we already explained in plain words (shown to the user as-is)."""


# Most likely cause for common HTTP status codes
HTTP_HINTS = {
    400: "Bad request. Check tenant ID, client ID, secret and scope.",
    401: "Unauthorized. The secret may be wrong or expired.",
    403: "Forbidden. The app needs a Purview role (e.g. Data Curator).",
    404: "Not found. Check the URL and api-version.",
    429: "Too many requests. Wait a bit and retry.",
}


def explain_error(error, step):
    """Turn a requests exception into a PurviewError with a readable message."""
    if isinstance(error, requests.exceptions.Timeout):
        return PurviewError(f"{step}: timed out. Check your internet/VPN and retry.")
    if isinstance(error, requests.exceptions.SSLError):
        return PurviewError(f"{step}: SSL error. A proxy or firewall may be blocking traffic.")
    if isinstance(error, requests.exceptions.ConnectionError):
        return PurviewError(f"{step}: cannot connect. Check your internet/VPN.")
    if isinstance(error, requests.exceptions.HTTPError):
        code = error.response.status_code
        try:  # the server usually explains the problem in the JSON body
            body = error.response.json()
            detail = body.get("error_description") or body.get("error") or body
        except ValueError:
            detail = error.response.text[:300]
        return PurviewError(f"{step}: HTTP {code}. {HTTP_HINTS.get(code, 'Unexpected response.')}\nDetails: {detail}")
    return PurviewError(f"{step}: {error}")


def request_access_token():
    """Log in as the app (client credentials) and return a bearer token."""
    try:
        response = requests.post(
            f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
            data={"grant_type": "client_credentials", "client_id": CLIENT_ID,
                  "client_secret": CLIENT_SECRET, "scope": "https://purview.azure.net/.default"},
            timeout=30)
        response.raise_for_status()  # turn HTTP errors (401, 400...) into exceptions
        return response.json()["access_token"]
    except requests.exceptions.RequestException as error:
        raise explain_error(error, "Getting access token") from error
    except (ValueError, KeyError) as error:  # reply was not JSON or had no token
        raise PurviewError("Getting access token: unexpected response format.") from error


def fetch_all_items(source, report=lambda text: None):
    """Get every item of a source (domains or terms), following 'nextLink' until no more pages.
    `report(text)` is called at each step so the screen can show live progress."""
    report("Signing in to Microsoft Entra ID...")
    headers = {"Authorization": f"Bearer {request_access_token()}"}
    return fetch_pages(SOURCES[source]["path"], headers, SOURCES[source]["label"], report)


def fetch_pages(path, headers, label, report=lambda text: None, extra_params=None):
    """GET BASE_URL + path and follow 'nextLink'; returns the combined 'value' list.
    `extra_params` adds query parameters such as entityType=TERM."""
    url = BASE_URL + path
    params, items, page_number = {"api-version": API_VERSION, **(extra_params or {})}, [], 0
    try:
        while url:
            page_number += 1
            report(f"Fetching {label}, page {page_number} ({len(items)} so far)...")
            response = requests.get(url, headers=headers, params=params, timeout=60)
            response.raise_for_status()
            page = response.json()
            items += page.get("value", [])
            url, params = page.get("nextLink"), None  # nextLink is a full URL with its own params
    except requests.exceptions.RequestException as error:
        raise explain_error(error, f"Fetching {label}") from error
    except (ValueError, AttributeError) as error:  # body was not a JSON object
        raise PurviewError(f"Fetching {label}: response was not valid JSON.") from error
    return items


# ---------------------------------------------------------------------------
# HELPERS: reshape API data for display
# ---------------------------------------------------------------------------
def make_progress():
    """A live progress display: spinner, step text, bar, percentage and elapsed time."""
    return Progress(SpinnerColumn(), TextColumn("[cyan]{task.description}"), BarColumn(),
                    TaskProgressColumn(), TimeElapsedColumn(), console=console)


def item_name(item):
    """Best available name: terms/domains may use different fields."""
    return str(item.get("name") or item.get("displayName") or item.get("nickName") or item.get("CDE ColumnName") or "-")


def remove_html_tags(text):
    """Descriptions come as '<div>text</div>'; keep only the text."""
    return re.sub(r"<[^>]+>", "", text or "").strip()


def flatten_fields(value, path=""):
    """Yield (field_path, value) for every field, however deeply nested.
    Example: systemData.createdAt, domains[0].relatedCollections[1].name"""
    if isinstance(value, dict) and value:
        for key, item in value.items():
            yield from flatten_fields(item, f"{path}.{key}" if path else key)
    elif isinstance(value, list) and value:
        for index, item in enumerate(value):
            yield from flatten_fields(item, f"{path}[{index}]")
    else:
        yield path, value  # a plain value (or empty dict/list): stop here


# ---------------------------------------------------------------------------
# SCREENS: one function per menu option (they work for domains and terms alike)
# ---------------------------------------------------------------------------
def print_items_table(items, title):
    """Numbered overview table (Name / Status / Description, or the columns of per-CDE items)."""
    if items and "CDE ColumnName" in items[0]:  # per-CDE items (DQ / Glossary / Ownership / Metadata)
        table = Table(title=f"{title} ({len(items)})", box=box.ROUNDED, header_style="bold cyan")
        table.add_column("#", justify="right", style="dim")
        for column in items[0]:
            table.add_column(column, overflow="fold", max_width=60)
        for number, item in enumerate(items, 1):
            table.add_row(str(number), *[yes_no(v) for v in item.values()])
        console.print(table)
        return
    table = Table(title=f"{title} ({len(items)})", box=box.ROUNDED, header_style="bold cyan")
    table.add_column("#", justify="right", style="dim")
    table.add_column("Name", style="bold")
    table.add_column("Status")
    table.add_column("Description", max_width=50, overflow="ellipsis", no_wrap=True)
    for number, item in enumerate(items, 1):
        status = item.get("status", "-")
        color = "green" if status == "Published" else "yellow"  # live vs draft
        table.add_row(str(number), escape(item_name(item)), f"[{color}]{status}[/]",
                      remove_html_tags(item.get("description")))
    console.print(table)


def search_items(items, label):
    """Option 3: show only items whose name contains the typed word."""
    word = Prompt.ask("Search for").lower()
    matches = [item for item in items if word in item_name(item).lower()]
    if matches:
        print_items_table(matches, f"{label.capitalize()} matching '{word}'")
    else:
        console.print("[yellow]No matches.[/]")


def print_item_details(items, label):
    """Option 4: pick one item and show ALL of its fields."""
    print_items_table(items, label.capitalize())  # so the user can see which number to pick
    # Tell the user exactly what to type: a row number from the "#" column, or 0 to cancel
    choices = [str(i) for i in range(0, len(items) + 1)]
    extra = {"default": 1} if len(items) == 1 else {}  # only one item: just press Enter
    number = IntPrompt.ask(
        f"\nEnter the [bold cyan]#[/] of the item to view (1-{len(items)}), or [bold]0[/] to go back",
        choices=choices, show_choices=False, **extra)
    if number == 0:
        return
    item = items[number - 1]
    print_item_panel(item)

    if Prompt.ask("Show raw JSON?", choices=["y", "n"], default="n") == "y":
        console.print(Syntax(json.dumps(item, indent=2), "json", theme="monokai"))


def print_item_panel(item, position=None):
    """Show ALL fields of one item in a bordered panel (used by options 4 and 8)."""
    details = Table(box=box.SIMPLE, show_header=False)
    details.add_column(style="bold cyan", no_wrap=True)  # field name
    details.add_column(overflow="fold")                  # field value (wraps if long)
    for field, value in flatten_fields(item):
        text = remove_html_tags(value) if isinstance(value, str) else str(value)
        details.add_row(field, escape(text) or "-")      # escape: stop [brackets] being read as markup
    prefix = f"{position}. " if position else ""
    console.print(Panel(details, title=f"[bold cyan]{prefix}{escape(item_name(item))}[/]", border_style="cyan"))


def print_all_item_details(items):
    """Option 8: show the full details of EVERY item, one panel after another."""
    for position, item in enumerate(items, 1):
        print_item_panel(item, position)


def flatten_items(items, tick=lambda text: None):
    """One flat row per item (every field, nested ones included) and the list of columns."""
    rows = []
    for item in items:
        tick(f"Preparing row for {item_name(item)}...")
        rows.append({field: remove_html_tags(v) if isinstance(v, str) else v
                     for field, v in flatten_fields(item)})
    return rows, list(dict.fromkeys(field for row in rows for field in row))  # columns in order seen


def fill_sheet(sheet, rows, columns):
    """Header + rows with a bold header, sensible column widths and a frozen header row."""
    sheet.append(columns)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([str(row.get(c, "")) if isinstance(row.get(c), (list, dict)) else row.get(c, "")
                      for c in columns])
    for index, column in enumerate(columns, 1):
        width = max(len(str(c.value or "")) for c in sheet[get_column_letter(index)])
        sheet.column_dimensions[get_column_letter(index)].width = min(width + 2, 60)
    sheet.freeze_panes = "A2"


def write_item_files(items, file_stem, tick=lambda text: None, folder=None):
    """Write items as <file_stem>.json, .csv and .xlsx in the output folder.
    `tick(text)` is called once per step: 1 (JSON) + 1 per item + 1 (CSV) + 1 (Excel) = len(items) + 3.
    Returns (flat rows, columns) so a combined workbook can reuse them."""
    folder = folder or OUTPUT_DIR
    folder.mkdir(parents=True, exist_ok=True)
    json_file, csv_file, xlsx_file = (folder / f"{file_stem}{ext}" for ext in (".json", ".csv", ".xlsx"))

    tick(f"Writing {json_file.name}...")  # JSON keeps the original nested structure
    with open(json_file, "w", encoding="utf-8") as file:
        json.dump({"value": items}, file, indent=2)

    rows, columns = flatten_items(items, tick)

    tick(f"Writing {csv_file.name}...")
    # utf-8-sig lets Excel read non-English text; newline="" avoids blank lines on Windows
    with open(csv_file, "w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)

    tick(f"Writing {xlsx_file.name}...")
    workbook = Workbook()
    workbook.active.title = file_stem[:31]  # Excel limits sheet names to 31 characters
    fill_sheet(workbook.active, rows, columns)
    workbook.save(xlsx_file)
    return rows, columns


def save_items_to_files(items, file_stem):
    """Save one dataset (JSON + CSV + Excel) with a progress bar."""
    with make_progress() as progress:
        task = progress.add_task("Starting...", total=len(items) + 3)

        def tick(text):
            progress.update(task, description=text)
            progress.advance(task)

        write_item_files(items, file_stem, tick)
        progress.update(task, description="Done", completed=len(items) + 3)
    console.print(f"[green]✔ Saved {file_stem}.json, {file_stem}.csv and {file_stem}.xlsx in {OUTPUT_DIR}[/]")


# ---------------------------------------------------------------------------
# COVERAGE REPORT: domain -> critical data elements -> glossary terms (loaded once with option 1)
# ---------------------------------------------------------------------------
def contains_value(data, wanted):
    """True if `wanted` appears anywhere inside nested JSON (dicts/lists) as a value."""
    if isinstance(data, dict):
        return any(contains_value(v, wanted) for v in data.values())
    if isinstance(data, list):
        return any(contains_value(v, wanted) for v in data)
    return data == wanted


def owner_ids(item):
    """The owner ids found in contacts.owner (empty list = no owner)."""
    owners = (item.get("contacts") or {}).get("owner")
    if isinstance(owners, (dict, str)):
        owners = [owners]
    ids = []
    for owner in owners or []:
        owner_id = owner.get("id") if isinstance(owner, dict) else owner
        if owner_id and str(owner_id).strip():
            ids.append(str(owner_id))
    return ids


def expert_ids(item):
    """The expert ids found in contacts.expert (empty list = no expert)."""
    experts = (item.get("contacts") or {}).get("expert")
    if isinstance(experts, (dict, str)):
        experts = [experts]
    ids = []
    for expert in experts or []:
        expert_id = expert.get("id") if isinstance(expert, dict) else expert
        if expert_id and str(expert_id).strip():
            ids.append(str(expert_id))
    return ids


def related_term_ids(relationships):
    """From a relationships response, return the glossary ids whose relationship type is 'Related'."""
    ids = []
    for rel in relationships:
        rel_type = str(rel.get("relationshipType", "")).lower()
        term_id = rel.get("entityId") or rel.get("id")
        if term_id and rel_type == "related":
            ids.append(term_id)
    return ids


def fetch_raw(report):
    """Steps 1-3: fetch the real Purview responses. Returns domains, terms, cdes and relationships (by CDE id)."""
    report("Signing in to Microsoft Entra ID...")
    headers = {"Authorization": f"Bearer {request_access_token()}"}
    domains = fetch_pages(SOURCES["domains"]["path"], headers, "domains", report)
    cdes = fetch_pages(SOURCES["cdes"]["path"], headers, "critical data elements", report)
    terms = fetch_pages(SOURCES["terms"]["path"], headers, "glossary terms", report)
    pairs = [(d, c) for d in domains for c in cdes if contains_value(c, d.get("id"))]
    relationships = {}
    for number, (domain, cde) in enumerate(pairs, 1):
        report(f"Checking relationships {number}/{len(pairs)}: {item_name(cde)}")
        try:
            relationships[cde.get("id")] = fetch_pages(
                f"/criticalDataElements/{cde.get('id')}/relationships", headers, "relationships",
                extra_params={"entityType": "TERM"})
        except PurviewError:  # e.g. 404 when a CDE has no relationships
            relationships[cde.get("id")] = []
    return {"domains": domains, "terms": terms, "cdes": cdes, "relationships": relationships}


def rows_from_raw(raw):
    """Steps 1-4 on raw responses (fresh from Purview or read back from a saved export): one row per CDE."""
    terms_by_id = {t.get("id"): t for t in raw["terms"]}
    term_names = {tid: item_name(t) for tid, t in terms_by_id.items()}
    rows = []
    for domain in raw["domains"]:
        for cde in raw["cdes"]:
            if not contains_value(cde, domain.get("id")):
                continue
            glossary_ids = related_term_ids(raw["relationships"].get(cde.get("id"), []))
            found_ids = [g for g in glossary_ids if g in term_names]  # Step 4: id exists in /terms
            # Ownership rule: a glossary term is linked AND that term has contacts.expert
            owners = [e for g in found_ids for e in expert_ids(terms_by_id[g])]
            description = remove_html_tags(cde.get("description"))
            rows.append({
                "Table Name": item_name(domain),
                "CDE ColumnName": item_name(cde),
                "DQ Score": "",  # no data quality source yet: fill in by hand in Excel (Y/N)
                "Glossary": "Y" if found_ids else "N",
                "Ownership": "Y" if owners else "N",
                "Metadata": "Y" if description else "N",
                "Domain ID": domain.get("id"),
                "CDE ID": cde.get("id"),
                "Glossary ID": ", ".join(glossary_ids),
                "Glossary Name": ", ".join(term_names[g] for g in found_ids),
                "Owner ID": ", ".join(owners),
                "Description": description,
            })
    return rows


def build_coverage_rows(report):
    """Fetch from Purview and score straight away (main menu option 1)."""
    return rows_from_raw(fetch_raw(report))


TRUST_WEIGHTS = CONFIG["weights"]  # from config.json, sum = 1.0


def trust_score(rows):
    """Trust score = sum(weight * % of rows marked Y) for each dimension. Blank DQ counts as 0%."""
    return sum(weight * 100 * sum(r[name] == "Y" for r in rows) / len(rows)
               for name, weight in TRUST_WEIGHTS.items())


def coverage_pct(rows, dimension):
    """Share of rows marked Y for one dimension (0-100)."""
    return 100 * sum(r[dimension] == "Y" for r in rows) / len(rows) if rows else 0


def detail_for(row, dimension, short=True):
    """The data behind a Y/N: the linked term, the owner ids, the description, or why it is N.
    short=True trims long values so they fit next to the other columns."""
    if dimension == "DQ Score":
        return ("no source" if short else "no DQ source yet") if not row["DQ Score"] else str(row["DQ Score"])
    if dimension == "Glossary":
        if row["Glossary"] == "Y":
            return row["Glossary Name"]
        return "term not published" if row["Glossary ID"] else "no term linked"
    if dimension == "Ownership":
        owners = [o for o in row["Owner ID"].split(", ") if o]
        if not owners:
            return "no expert on term" if row["Glossary"] == "Y" else "no glossary term"
        return f"{len(owners)} expert(s)" if short else f"{len(owners)} expert(s): " + ", ".join(owners)
    if dimension == "Metadata":
        text = row["Description"]
        if not text:
            return "no description"
        return text[:22] + "..." if short and len(text) > 22 else text[:200]
    return ""


def yes_no(value):
    """Green Y / red N for the screen; anything else as plain text ('-' when empty)."""
    if value in ("Y", "N"):
        return f"[{'green' if value == 'Y' else 'red'}]{value}[/]"
    return escape(str(value)) or "-"


DASHBOARD_EXTRAS = None  # optional function(dash=..., ...) a script sets to add a section to the dashboard
# Where the "Chart data" block sits on the Dashboard (a script may move it): first column, header row, header names,
# and whether the two summary charts are drawn
CHART_DATA = {"col": 14, "row": 4, "heads": ["Dimension", "Yes", "No", "Coverage %", "Weight"], "title": "Chart data",
              "charts": True}
EXCEL_HIDE = {"Table Name"}  # row fields left out of the Excel CDE sheet (the other script shows them)
GREEN, RED, AMBER = "C6EFCE", "FFC7CE", "FFEB9C"   # soft fills for Y / N / medium
NAVY = "1F3864"


# (minimum score, label, colour) from config.json: 90 Decision-grade, 75 Trusted, 60 Use with caution, below At risk
BANDS = [(int(m) if m == int(m) else m, label, colour)
         for (m, label), colour in zip(CONFIG["bands"], BAND_COLOURS)]


def band_of(score):
    """The band label for a score."""
    for minimum, label, _ in BANDS:
        if score >= minimum:
            return label
    return BANDS[-1][1]


def formula_text(separator=" + "):
    """e.g. '40% DQ Score + 20% Glossary + 20% Ownership + 20% Metadata' (from the weights)."""
    return separator.join(f"{weight:.0%} {name}" for name, weight in TRUST_WEIGHTS.items())


def band_formula(ref):
    """Excel IF(...) that turns a score cell into its band label."""
    expr = f'"{BANDS[-1][1]}"'
    for minimum, label, _ in reversed(BANDS[:-1]):
        expr = f'IF({ref}>={minimum},"{label}",{expr})'
    return "=" + expr


def score_colours(sheet, cell):
    """Colour a score cell by band: Trusted or better green, Use with caution amber, At risk red (updates live)."""
    good, ok = BANDS[1][0], BANDS[2][0]
    for operator, formula, colour in (("greaterThanOrEqual", [str(good)], GREEN),
                                      ("between", [str(ok), str(good - 0.001)], AMBER),
                                      ("lessThan", [str(ok)], RED)):
        sheet.conditional_formatting.add(cell, CellIsRule(
            operator=operator, formula=formula, fill=PatternFill("solid", bgColor=colour)))


DIMENSION_COLOURS = {"DQ Score": "2E75B6", "Glossary": "00A3A1", "Ownership": "ED7D31", "Metadata": "8E7CC3"}


def add_dashboard(workbook, letters, last, rows_info, table_col=None, table_names=None, dq_columns=None,
                  column_col=None):
    """Add a 'Dashboard' sheet: a big Trust Score, one card per dimension, the score bands and three charts.
    Everything is a formula, so it refreshes whenever Y/N values change on the 'CDE coverage' sheet."""
    pct_row, weight_row, score_row = rows_info
    data = "'CDE coverage'"
    dash = workbook.create_sheet("Dashboard")
    dash.sheet_view.showGridLines = False
    for col in "ABCDEFGHIJKL":
        dash.column_dimensions[col].width = 13
    for col in "MNOPQRS":
        dash.column_dimensions[col].width = 12
    thin = Side(style="thin", color="D9D9D9")
    box_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    centre = Alignment(horizontal="center", vertical="center")

    # Banner
    dash.merge_cells("A1:L1")
    dash["A1"] = "DATA TRUST DASHBOARD"
    dash["A1"].font = Font(size=22, bold=True, color="FFFFFF")
    dash["A1"].fill = PatternFill("solid", fgColor=NAVY)
    dash["A1"].alignment = centre
    dash.row_dimensions[1].height = 42
    dash.merge_cells("A2:L2")
    # With a table column, a dropdown (B3) filters every number and chart below to one table (or all tables)
    pick = "$D$3"
    table_range = f"{data}!${table_col}$2:${table_col}${last}" if table_col else None
    total = (f'IF({pick}="All Domain",ROWS({data}!A2:A{last}),COUNTIF({table_range},{pick}))'
             if table_col else f"ROWS({data}!A2:A{last})")
    dash["A2"] = f'={total}&" rows checked   |   Trust Score = {formula_text()}   |   Set DQ Score (Y/N) on the CDE coverage sheet"'
    dash["A2"].font = Font(italic=True, size=9, color="FFFFFF")
    dash["A2"].fill = PatternFill("solid", fgColor="2F5597")
    dash["A2"].alignment = centre

    # Chart data (light grey): Yes / No / Coverage % / Weight
    c0, hr = CHART_DATA["col"], CHART_DATA["row"]
    cl = [get_column_letter(c0 + i) for i in range(5)]      # Dimension, Yes, No, Coverage %, Weight
    dash.cell(row=hr - 1, column=c0, value=CHART_DATA["title"]).font = Font(
        bold=bool(CHART_DATA.get("bold_title")), italic=not CHART_DATA.get("bold_title"), size=12 if CHART_DATA.get("bold_title") else 9,
        color=NAVY if CHART_DATA.get("bold_title") else "808080")
    for offset, name in enumerate(TRUST_WEIGHTS, hr + 1):
        col = letters[name]
        dash[f"{cl[0]}{offset}"] = name
        every = f'COUNTIF({data}!{col}2:{col}{last},"Y")'
        if table_col:
            one = f'COUNTIFS({table_range},{pick},{data}!{col}2:{col}{last},"Y")'
            dash[f"{cl[1]}{offset}"] = f'=IF({pick}="All Domain",{every},{one})'
        else:
            dash[f"{cl[1]}{offset}"] = f"={every}"
        dash[f"{cl[2]}{offset}"] = f"={total}-{cl[1]}{offset}"  # anything not Y (blank DQ included)
        dash[f"{cl[3]}{offset}"] = f"=IF({cl[1]}{offset}+{cl[2]}{offset}=0,0,{cl[1]}{offset}/({cl[1]}{offset}+{cl[2]}{offset})*100)"
        dash[f"{cl[4]}{offset}"] = f"={data}!{col}{weight_row}"
    # colour the chart data: navy header, a colour chip per dimension, green Yes, red No, banded rows
    table_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for index, text in enumerate(CHART_DATA["heads"]):
        head = dash.cell(row=hr, column=c0 + index, value=text)
        head.font = Font(bold=True, color="FFFFFF", size=10)
        head.fill = PatternFill("solid", fgColor=NAVY)
        head.alignment = Alignment(horizontal="center", vertical="center")
        head.border = table_border
    for index, name in enumerate(TRUST_WEIGHTS):
        row = hr + 1 + index
        band = PatternFill("solid", fgColor="F2F2F2" if index % 2 else "FFFFFF")
        for column in range(c0, c0 + 5):
            cell = dash.cell(row=row, column=column)
            cell.font = Font(size=10, color="404040")
            cell.alignment = Alignment(horizontal="center" if column > c0 else "left", vertical="center")
            cell.fill = band
            cell.border = table_border
        dash[f"{cl[0]}{row}"].font = Font(size=10, bold=True, color="FFFFFF")
        dash[f"{cl[0]}{row}"].fill = PatternFill("solid", fgColor=DIMENSION_COLOURS[name])
        dash[f"{cl[1]}{row}"].fill = PatternFill("solid", fgColor=GREEN)
        dash[f"{cl[2]}{row}"].fill = PatternFill("solid", fgColor=RED)
        dash[f"{cl[3]}{row}"].number_format = '0"%"'
        dash[f"{cl[4]}{row}"].number_format = "0%"
        score_colours(dash, f"{cl[3]}{row}")
    if c0 == 14:
        dash.column_dimensions["N"].width = 14
        dash.column_dimensions["P"].width = 17
    count = len(TRUST_WEIGHTS)

    if table_col:
        dash.merge_cells("A3:C3")  # label block
        dash["A3"] = "SELECT DOMAIN  ▶"
        dash["A3"].font = Font(bold=True, size=13, color="FFFFFF")
        dash["A3"].fill = PatternFill("solid", fgColor=NAVY)
        dash["A3"].alignment = centre
        dash.merge_cells("D3:G3")  # the dropdown cell
        dash["D3"] = "All Domain"
        dash["D3"].font = Font(bold=True, size=16, color=NAVY)
        dash["D3"].fill = PatternFill("solid", fgColor="FFE699")
        dash["D3"].alignment = centre
        heavy = Side(style="medium", color=NAVY)
        for column, cell in enumerate(("D3", "E3", "F3", "G3")):
            dash[cell].border = Border(top=heavy, bottom=heavy, left=heavy if column == 0 else None,
                                       right=heavy if column == 3 else None)
            dash[cell].fill = PatternFill("solid", fgColor="FFE699")
        dash.merge_cells("H3:L3")  # hint
        dash["H3"] = "▼  click the yellow cell and pick a domain - the whole dashboard updates"
        dash["H3"].font = Font(italic=True, size=10, color="7F6000")
        dash["H3"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
        options = ["All Domain"] + list(table_names)
        dash["V4"] = "Table list"
        dash["V4"].font = Font(italic=True, color="808080", size=9)
        for offset, name in enumerate(options, 5):
            dash[f"V{offset}"] = name
            dash[f"V{offset}"].font = Font(size=9, color="808080")
        picker = DataValidation(type="list", allow_blank=False, formula1=f"=$V$5:$V${4 + len(options)}")
        dash.add_data_validation(picker)
        picker.add("D3")
        dash.row_dimensions[3].height = 34

    # Hero tile: the Trust Score and its band
    dash.merge_cells("A4:D4")
    dash["A4"] = f'="TRUST SCORE: "&{pick}' if table_col else "TRUST SCORE"
    dash["A4"].font = Font(bold=True, color="FFFFFF", size=12)
    dash["A4"].fill = PatternFill("solid", fgColor=NAVY)
    dash["A4"].alignment = centre
    dash.merge_cells("A5:D7")
    first_row, last_row = hr + 1, hr + len(TRUST_WEIGHTS)
    dash["A5"] = f"=SUMPRODUCT({cl[3]}{first_row}:{cl[3]}{last_row},{cl[4]}{first_row}:{cl[4]}{last_row})"  # follows the selection
    dash["A5"].number_format = "0.0"
    dash["A5"].font = Font(size=44, bold=True, color="1F3864")
    dash["A5"].alignment = centre
    score_colours(dash, "A5")
    dash.merge_cells("A8:D8")
    dash["A8"] = band_formula("A5")
    dash["A8"].font = Font(bold=True, size=13, color="1F3864")
    dash["A8"].alignment = centre
    for formula, colour in ((f"$A$5>={BANDS[1][0]}", GREEN), (f"AND($A$5>={BANDS[2][0]},$A$5<{BANDS[1][0]})", AMBER),
                            (f"$A$5<{BANDS[2][0]}", RED)):
        dash.conditional_formatting.add("A8", FormulaRule(formula=[formula], fill=PatternFill("solid", bgColor=colour)))
    for row in range(4, 9):
        for column in range(1, 5):
            dash.cell(row=row, column=column).border = box_border
    for row in (5, 6, 7):
        dash.row_dimensions[row].height = 24

    # One card per dimension: label, coverage %, weight, "x of y CDEs"
    for index, name in enumerate(TRUST_WEIGHTS):
        left, right = get_column_letter(5 + index * 2), get_column_letter(6 + index * 2)
        source = hr + 1 + index
        for rows_span in ((4, 4), (5, 6), (7, 7), (8, 8)):  # label / big number / weight / count (no overlaps)
            dash.merge_cells(f"{left}{rows_span[0]}:{right}{rows_span[1]}")
        dash[f"{left}4"] = name.upper()
        dash[f"{left}4"].font = Font(bold=True, color="FFFFFF", size=11)
        dash[f"{left}4"].fill = PatternFill("solid", fgColor=DIMENSION_COLOURS[name])
        dash[f"{left}4"].alignment = centre
        dash[f"{left}5"] = f"={cl[3]}{source}"
        dash[f"{left}5"].number_format = '0"%"'
        dash[f"{left}5"].font = Font(size=26, bold=True, color="404040")
        dash[f"{left}5"].alignment = centre
        score_colours(dash, f"{left}5")
        dash[f"{left}7"] = f'="weight "&TEXT({cl[4]}{source},"0%")'
        dash[f"{left}7"].font = Font(size=9, color="595959")
        dash[f"{left}7"].alignment = centre
        dash[f"{left}8"] = f'={cl[1]}{source}&" of "&({cl[1]}{source}+{cl[2]}{source})&" CDEs"'
        dash[f"{left}8"].font = Font(size=9, color="595959")
        dash[f"{left}8"].alignment = centre
        for row in range(4, 9):
            for column in (5 + index * 2, 6 + index * 2):
                dash.cell(row=row, column=column).border = box_border

    # Band legend
    dash["A10"] = "Score bands"
    dash["A10"].font = Font(bold=True, color="595959", size=9)
    for index, (minimum, label, colour) in enumerate(BANDS):
        left, right = get_column_letter(1 + index * 3), get_column_letter(3 + index * 3)
        dash.merge_cells(f"{left}11:{right}11")
        upper = "100" if index == 0 else str(BANDS[index - 1][0] - 1)
        dash[f"{left}11"] = f"{minimum}-{upper}  {label}" if index < len(BANDS) - 1 else f"below {BANDS[index - 1][0]}  {label}"
        dash[f"{left}11"].fill = PatternFill("solid", fgColor=colour)
        dash[f"{left}11"].font = Font(bold=True, size=10, color="404040")
        dash[f"{left}11"].alignment = centre
        dash.row_dimensions[11].height = 22

    def place(chart, title, anchor, width=14):
        chart.title, chart.width, chart.height = title, width, 8
        chart.x_axis.delete = chart.y_axis.delete = False  # make sure axes are drawn
        chart.y_axis.majorGridlines = None
        dash.add_chart(chart, anchor)

    def colour_points(series, colours):
        for position, colour in enumerate(colours):
            point = DataPoint(idx=position)
            point.graphicalProperties.solidFill = colour
            series.dPt.append(point)

    dim_colours = [DIMENSION_COLOURS[n] for n in TRUST_WEIGHTS]

    if CHART_DATA["charts"]:
        # Chart 1: coverage % per dimension (one colour per dimension, value labels)
        coverage = BarChart()
        coverage.type = "bar"
        coverage.add_data(Reference(dash, min_col=c0 + 3, min_row=hr, max_row=hr + count), titles_from_data=True)
        coverage.set_categories(Reference(dash, min_col=c0, min_row=hr + 1, max_row=hr + count))
        coverage.y_axis.scaling.min, coverage.y_axis.scaling.max = 0, 100
        coverage.legend = None
        coverage.gapWidth = 60
        colour_points(coverage.series[0], dim_colours)
        coverage.dataLabels = DataLabelList()
        coverage.dataLabels.showVal = True
        coverage.dataLabels.showSerName = coverage.dataLabels.showCatName = coverage.dataLabels.showLegendKey = False
        coverage.dataLabels.showPercent = False
        coverage.dataLabels.numFmt = '0"%"'          # whole numbers only: 96%, not 96.2%
        coverage.y_axis.number_format = '0"%"'
        coverage.y_axis.majorUnit = 20
        place(coverage, "Coverage % by dimension", "A13")

        # Chart 2: share of CDEs with Y and N in each dimension
        breakdown = BarChart()
        breakdown.type, breakdown.grouping, breakdown.overlap = "col", "stacked", 100
        breakdown.add_data(Reference(dash, min_col=c0 + 1, max_col=c0 + 2, min_row=hr, max_row=hr + count), titles_from_data=True)
        breakdown.set_categories(Reference(dash, min_col=c0, min_row=hr + 1, max_row=hr + count))
        breakdown.series[0].graphicalProperties.solidFill = "63BE7B"
        breakdown.series[1].graphicalProperties.solidFill = "F8696B"
        breakdown.gapWidth = 60
        breakdown.legend.position = "r"  # to the right (does not overlap the title or axis labels)
        breakdown.legend.overlay = False
        place(breakdown, "CDEs with Y and N", "G13")

    # Extra dashboard section supplied by a script (purview_trust_by_domain.py draws its table-wise DQ grid here)
    if DASHBOARD_EXTRAS and dq_columns and table_col and table_names:
        DASHBOARD_EXTRAS(dash=dash, data=data, last=last, table_range=table_range, table_names=table_names,
                         dq_columns=dq_columns)

    workbook.move_sheet("Dashboard", offset=-1)  # make it the first tab
    workbook.active = 0


def letters_for(columns, name):
    """Excel column letter of a column name, or None when that column is not in the sheet."""
    return get_column_letter(columns.index(name) + 1) if name in columns else None


def save_report_to_excel(rows):
    """Write the CDE rows plus the Trust Score section to an .xlsx file (sheets: Dashboard, CDE coverage).
    Y cells are green, N cells red. The scores are real Excel formulas, so filling in
    the DQ Score column by hand (Y/N) updates them automatically."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CDE coverage"
    columns = [c for c in rows[0] if c not in EXCEL_HIDE]  # Table Name is hidden here: every CDE is in one domain
    sheet.append(columns)
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    for row in rows:
        sheet.append([row[c] for c in columns])
    last = len(rows) + 1  # last data row

    # Y/N cells: colours follow the value (conditional formatting), so they update when you type
    first_dim = get_column_letter(columns.index("DQ Score") + 1)
    last_dim = get_column_letter(columns.index("Metadata") + 1)
    yn_range = f"{first_dim}2:{last_dim}{last}"
    sheet.conditional_formatting.add(yn_range, CellIsRule(
        operator="equal", formula=['"Y"'], fill=PatternFill("solid", bgColor=GREEN)))
    sheet.conditional_formatting.add(yn_range, CellIsRule(
        operator="equal", formula=['"N"'], fill=PatternFill("solid", bgColor=RED)))
    for name in TRUST_WEIGHTS:
        for r in range(2, last + 1):
            sheet.cell(row=r, column=columns.index(name) + 1).alignment = Alignment(horizontal="center")
    # DQ Score is typed by hand: give it a Y/N dropdown
    dq_col = get_column_letter(columns.index("DQ Score") + 1)
    dropdown = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    sheet.add_data_validation(dropdown)
    dropdown.add(f"{dq_col}2:{dq_col}{last}")
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{last}"  # filter buttons on the header
    sheet.freeze_panes = "A2"

    # Dimension columns and their letters
    letters = {name: get_column_letter(columns.index(name) + 1) for name in TRUST_WEIGHTS}

    # Overall section: coverage % per dimension, weights, and the Trust Score
    pct_row, weight_row, score_row = last + 2, last + 3, last + 4
    sheet.cell(row=pct_row, column=1, value="Coverage %").font = Font(bold=True)
    sheet.cell(row=weight_row, column=1, value="Weight").font = Font(bold=True)
    sheet.cell(row=score_row, column=1, value="Trust Score").font = Font(bold=True)
    for name, weight in TRUST_WEIGHTS.items():
        col = letters[name]
        sheet[f"{col}{pct_row}"] = f'=COUNTIF({col}2:{col}{last},"Y")/ROWS({col}2:{col}{last})*100'
        sheet[f"{col}{weight_row}"] = weight
    first, final = letters["DQ Score"], letters["Metadata"]
    sheet[f"B{score_row}"] = f"=SUMPRODUCT({first}{pct_row}:{final}{pct_row},{first}{weight_row}:{final}{weight_row})"
    sheet[f"B{score_row}"].font = Font(bold=True)

    for index, column in enumerate(columns, 1):  # widen columns to fit their content
        width = max(len(str(c.value or "")) for c in sheet[get_column_letter(index)] if not str(c.value or "").startswith("="))
        sheet.column_dimensions[get_column_letter(index)].width = min(max(width, len(column)) + 2, 60)

    table_col = letters_for(columns, "Table Name")
    tables = list(dict.fromkeys(r["Table Name"] for r in rows)) if table_col else None
    # one 'DQ <dimension> %' column per data quality dimension (present when the rows carry them)
    dq_columns = {c[3:-2]: get_column_letter(columns.index(c) + 1) for c in columns
                  if c.startswith("DQ ") and c.endswith(" %")}
    add_dashboard(workbook, letters, last, (pct_row, weight_row, score_row), table_col, tables, dq_columns,
                  letters_for(columns, "Column Name"))
    workbook.save(REPORT_FILE)


def load_with_progress():
    """Main menu option 1: load everything once from Purview (live progress bar)."""
    with make_progress() as progress:
        task = progress.add_task("Starting...", total=None)
        rows = build_coverage_rows(lambda text: progress.update(task, description=text))
        progress.update(task, description="Done", total=1, completed=1)
    console.print(f"[green]✔ Loaded {len(rows)} CDE(s)[/]\n")
    return rows


def show_cde_table(rows, with_data):
    """One line per CDE with the Y/N of every dimension.
    with_data=True adds a 'based on' column after each Y/N showing the data behind it."""
    table = Table(title=f"CDE coverage ({len(rows)})", box=box.ROUNDED, header_style="bold cyan")
    table.add_column("Table Name")
    table.add_column("CDE ColumnName")
    for dimension in TRUST_WEIGHTS:
        table.add_column(dimension)
        if with_data:
            table.add_column(f"{dimension.split()[0]} based on", overflow="fold", style="dim")
    for row in rows:
        cells = [escape(row["Table Name"]), escape(row["CDE ColumnName"])]
        for dimension in TRUST_WEIGHTS:
            cells.append(yes_no(row[dimension]))
            if with_data:
                cells.append(escape(detail_for(row, dimension)))
        table.add_row(*cells)
    console.print(table)


def show_scores(rows):
    """The scores: per table and overall (share of Y per dimension, then the weighted sum)."""
    by_table = Table(title="Trust Score by table", box=box.ROUNDED, header_style="bold cyan")
    for column in ("Table Name", "CDEs", *[f"{d} %" for d in TRUST_WEIGHTS], "Trust Score"):
        by_table.add_column(column, justify="right" if column != "Table Name" else "left")
    groups = {name: [r for r in rows if r["Table Name"] == name] for name in dict.fromkeys(r["Table Name"] for r in rows)}
    groups["ALL"] = rows
    for name, subset in groups.items():
        by_table.add_row(escape(name), str(len(subset)),
                         *[f"{coverage_pct(subset, d):.0f}" for d in TRUST_WEIGHTS],
                         f"[bold]{trust_score(subset):.1f}[/]", style="bold" if name == "ALL" else None)
    console.print(by_table)
    console.print(Panel(
        f"[bold]Trust Score: {trust_score(rows):.1f}  ({band_of(trust_score(rows))})[/]\n"
        f"= {formula_text()} (each = % of CDEs marked Y)\n"
        "Bands: " + ", ".join(f"{m}+ {label}" if m else label for m, label, _ in BANDS) + "\n"
        + ("" if any(r["DQ Score"] for r in rows) else
           f"[yellow]DQ Score is blank for now (no data quality source yet), so it counts as 0%: "
           f"the score can reach at most {100 * (1 - TRUST_WEIGHTS['DQ Score']):.0f}.[/]"),
        border_style="cyan", expand=False))


def view_trust_score(rows):
    """Option 2: Trust Score - the Y/N per CDE (no 'based on' columns), then the scores."""
    show_cde_table(rows, with_data=False)
    show_scores(rows)


def view_trust_with_data(rows):
    """Option 3: the same, plus a 'based on' column after every Y/N showing the data behind it."""
    show_cde_table(rows, with_data=True)
    show_scores(rows)


def show_dimension(rows, dimension, note=None):
    """One line per CDE: the Y/N for one area and a 'Based on' column with the full data behind it."""
    table = Table(title=f"{dimension} per CDE ({len(rows)})", box=box.ROUNDED, header_style="bold cyan")
    table.add_column("Table Name")
    table.add_column("CDE ColumnName")
    table.add_column(dimension)
    table.add_column("Based on", overflow="fold", max_width=70)
    for row in rows:
        table.add_row(escape(row["Table Name"]), escape(row["CDE ColumnName"]), yes_no(row[dimension]),
                      escape(detail_for(row, dimension, short=False)))
    console.print(table)
    yes = sum(r[dimension] == "Y" for r in rows)
    lines = [f"[bold]{dimension} coverage: {coverage_pct(rows, dimension):.0f}%[/]  ({yes} of {len(rows)} CDEs)",
             f"Weight in the Trust Score: {TRUST_WEIGHTS[dimension]:.0%}"]
    if note:
        lines.append(note)
    console.print(Panel("\n".join(lines), border_style="cyan", expand=False))


def save_excel_report(rows):
    """Option 4: write the Excel report (CDE sheet, summary and dashboard) from the loaded data."""
    save_report_to_excel(rows)
    console.print(f"[green]✔ Saved {REPORT_FILE.name} in {OUTPUT_DIR}[/]")


def show_error(error):
    """Show any failure in a red box so the program keeps running."""
    if isinstance(error, PurviewError):  # known API/network problem, already explained
        message, title = str(error), "Error"
    elif isinstance(error, OSError):  # e.g. cannot write a file (is it open in Excel?)
        message, title = f"File problem: {error}", "Error"
    else:  # last-resort net: show it instead of crashing
        message, title = f"{type(error).__name__}: {error}", "Unexpected error"
    console.print(Panel(message, title=f"[red]{title}[/]", border_style="red"))


# ---------------------------------------------------------------------------
# BROWSE & EXPORT: every dataset can be listed, searched, opened and saved; or export them all
# ---------------------------------------------------------------------------
# key -> (screen name, kind, source key or dimension, file name without extension)
DATASETS = {
    "domains": ("Business domains", "source", "domains", "domains"),
    "terms": ("Glossary terms", "source", "terms", "glossary_terms"),
    "cdes": ("Critical data elements", "source", "cdes", "cdes"),
    "dq": ("DQ Score per CDE", "area", "DQ Score", "dq_score_per_cde"),
    "glossary": ("Glossary per CDE", "area", "Glossary", "glossary_per_cde"),
    "ownership": ("Ownership per CDE", "area", "Ownership", "ownership_per_cde"),
    "metadata": ("Metadata per CDE", "area", "Metadata", "metadata_per_cde"),
}
AREA_NOTES = {"DQ Score": "[yellow]No data quality source yet: DQ Score is blank and counts as 0% "
                          "until the DQ endpoint is added.[/]"}


def area_items(rows, dimension):
    """Per-CDE items for one area: table, CDE, the Y/N, and the full data it is based on."""
    return [{"Table Name": r["Table Name"], "CDE ColumnName": r["CDE ColumnName"], dimension: r[dimension],
             "Based on": detail_for(r, dimension, short=False)} for r in rows]


def get_items(state, key):
    """The items of one dataset as loaded so far ([] when not loaded)."""
    _, kind, ref, _ = DATASETS[key]
    if kind == "source":
        return state["items"].get(ref, [])
    return area_items(state["rows"], ref) if state["rows"] else []


def load_source(state, key):
    """Fetch domains / glossary terms / CDEs from Purview with a progress bar."""
    with make_progress() as progress:
        task = progress.add_task("Starting...", total=None)
        state["items"][key] = fetch_all_items(key, lambda text: progress.update(task, description=text))
        progress.update(task, description="Done", total=1, completed=1)
    return state["items"][key]


def show_items(state, key):
    """List one dataset: the area datasets get their coverage panel, the others a plain table."""
    label, kind, ref, _ = DATASETS[key]
    if kind == "area":
        show_dimension(state["rows"], ref, AREA_NOTES.get(ref))
    else:
        print_items_table(get_items(state, key), label)


def dataset_menu(state, key):
    """Load, list, search, open one item, save, or show all items of one dataset."""
    label, kind, ref, stem = DATASETS[key]
    while True:
        items = get_items(state, key)
        console.print(Panel(dataset_menu_text(), title=f"{label}", border_style="blue", expand=False))
        choice = Prompt.ask("Choose", choices=list("0123456"), default="1" if not items else "2")
        if choice == "0":
            return
        try:
            if choice == "1":
                if kind == "source":
                    load_source(state, ref)
                else:  # per-CDE data comes from the main data: refreshing it reloads everything
                    state["rows"] = load_with_progress()
                console.print(f"[green]✔ Loaded {len(get_items(state, key))} item(s)[/]")
                if get_items(state, key):
                    show_items(state, key)
            elif not items:
                console.print("[yellow]Nothing loaded yet. Choose 1 first.[/]")
            elif choice == "2":
                show_items(state, key)
            elif choice == "3":
                search_items(items, label)
            elif choice == "4":
                print_item_details(items, label)
            elif choice == "5":
                save_items_to_files(items, stem)
            elif choice == "6":
                print_all_item_details(items)
        except Exception as error:
            show_error(error)
        console.print()


EXPORT_PREFIX = "purview_export_"
RAW_FILES = ("domains", "terms", "cdes", "relationships")  # the real Purview responses, one JSON each


def export_all(state):
    """Fetch the real Purview responses and save them in a NEW folder (output/purview_export_<time>):
    domains, terms, cdes (JSON + CSV + Excel) and relationships (JSON). Nothing computed is saved."""
    with make_progress() as progress:
        task = progress.add_task("Starting...", total=None)
        raw = fetch_raw(lambda text: progress.update(task, description=text))
        progress.update(task, description="Saving...", total=1, completed=0)
        folder = OUTPUT_DIR / f"{EXPORT_PREFIX}{datetime.now():%Y%m%d_%H%M%S}"
        folder.mkdir(parents=True, exist_ok=True)
        for key in ("domains", "terms", "cdes"):
            write_item_files(raw[key], key, folder=folder)
        (folder / "relationships.json").write_text(json.dumps(raw["relationships"], indent=2), encoding="utf-8")
        progress.update(task, description="Done", completed=1)
    for key in ("domains", "terms", "cdes"):
        state["items"][key] = raw[key]
    counts = ", ".join(f"{len(raw[k])} {k}" for k in ("domains", "terms", "cdes"))
    console.print(f"[green]✔ Saved the real Purview responses ({counts}) in {folder}[/]")
    if Prompt.ask("Create the Trust Score from this export now?", choices=["y", "n"], default="y") == "y":
        state["rows"] = rows_from_raw(raw)
        show_trust_from_rows(state["rows"])


def load_saved_export():
    """Ask which saved export folder to use (newest first) and read its JSON files back."""
    folders = sorted((p for p in OUTPUT_DIR.glob(f"{EXPORT_PREFIX}*") if p.is_dir()), reverse=True)
    if not folders:
        raise PurviewError("No saved export found. Use 'Export ALL' first (it saves the real Purview responses).")
    for number, folder in enumerate(folders, 1):
        console.print(f"[bold cyan]{number}[/] {folder.name}")
    pick = Prompt.ask("Use which export", choices=[str(i) for i in range(1, len(folders) + 1)], default="1")
    folder = folders[int(pick) - 1]
    try:
        raw = {key: json.loads((folder / f"{key}.json").read_text(encoding="utf-8")) for key in RAW_FILES}
    except (OSError, ValueError) as error:
        raise PurviewError(f"Could not read the export in {folder}: {error}") from error
    for key in ("domains", "terms", "cdes"):  # the list files are saved as {"value": [...]}
        if isinstance(raw[key], dict):
            raw[key] = raw[key].get("value", [])
    console.print(f"[green]✔ Read {folder.name}[/]")
    return raw


def show_trust_from_rows(rows):
    """Trust Score screen (with the data behind each Y/N) for rows built from an export."""
    if not rows:
        console.print("[yellow]No critical data elements found in this data.[/]")
        return
    view_trust_with_data(rows)


def trust_from_saved_export(state):
    """Create the Trust Score from a saved export, without calling Purview."""
    state["rows"] = rows_from_raw(load_saved_export())
    show_trust_from_rows(state["rows"])


def browse_menu(state):
    """Option 5: pick a dataset to browse (and save), or export all of them."""
    keys = list(DATASETS)
    while True:
        lines = []
        for number, key in enumerate(keys, 1):
            count = len(get_items(state, key))
            status = f"loaded {count}" if count else "not loaded"
            lines.append(f"[bold cyan]{number}[/] {DATASETS[key][0]} [dim]({status})[/]")
        lines.append(f"[bold cyan]{len(keys) + 1}[/] Export ALL: fetch the real Purview responses into a new folder")
        lines.append(f"[bold cyan]{len(keys) + 2}[/] Trust Score from a saved export (no Purview call)")
        lines.append("[bold cyan]0[/] Back to main menu")
        console.print(Panel("\n".join(lines), title="Browse & export", border_style="blue", expand=False))
        pick = Prompt.ask("Choose", choices=[str(i) for i in range(0, len(keys) + 3)], default="1")
        if pick == "0":
            return
        try:
            if int(pick) == len(keys) + 1:
                export_all(state)
            elif int(pick) == len(keys) + 2:
                trust_from_saved_export(state)
            else:
                dataset_menu(state, keys[int(pick) - 1])
        except Exception as error:
            show_error(error)
        console.print()


# ---------------------------------------------------------------------------
# MAIN MENU: load first (1), then Trust Score (2, 3), Excel (4), browse & export (5)
# ---------------------------------------------------------------------------
def main():
    console.print(Panel.fit("[bold cyan]Microsoft Purview Trust Score[/]", border_style="cyan"))
    # rows: the CDE data from option 1 (Trust Score screens, per-CDE datasets); items: raw domains / terms / CDEs
    state = {"rows": None, "items": {}}
    while True:
        console.print(Panel(main_menu_text(state["rows"]), title="Main menu", border_style="blue", expand=False))
        choice = Prompt.ask("Choose", choices=list("012345"), default="1")
        if choice == "0":
            console.print("[cyan]Bye![/]")
            break
        try:
            if choice == "5":
                browse_menu(state)  # loads what it needs itself, so it works without option 1
            elif choice == "1":
                state["rows"] = load_with_progress()
            elif state["rows"] is None:
                console.print("[yellow]No data loaded yet. Choose 1 (Load data from Purview) first.[/]")
            elif not state["rows"]:
                console.print("[yellow]No critical data elements found for any domain.[/]")
            elif choice == "2":
                view_trust_score(state["rows"])
            elif choice == "3":
                view_trust_with_data(state["rows"])
            else:
                save_excel_report(state["rows"])
        except Exception as error:
            show_error(error)
        console.print()


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):  # Ctrl+C: exit politely, no traceback
        console.print("\n[yellow]Bye![/]")

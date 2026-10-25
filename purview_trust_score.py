"""Small interactive Purview explorer (business domains + glossary terms).  pip install requests rich"""
import csv, json, sys
from datetime import datetime
from rich import box
from rich.markup import escape
from rich.panel import Panel
from rich.progress import (BarColumn, Progress, SpinnerColumn, TaskProgressColumn,
                           TextColumn, TimeElapsedColumn)
from rich.prompt import IntPrompt, Prompt
from rich.syntax import Syntax
from rich.table import Table

from purview_engine.api import (HTTP_HINTS, PurviewError, explain_error, fetch_all_items, fetch_pages,
                                request_access_token)
from purview_engine.data import (band_formula, band_of, build_coverage_rows, contains_value, coverage_pct,
                                 detail_for, expert_ids, fetch_raw, formula_text, owner_ids, related_term_ids,
                                 rows_from_raw, trust_score, yes_no)
from purview_engine.helpers import flatten_fields, item_name, remove_html_tags
from purview_engine.settings import (AMBER, API_VERSION, APP_DIR, BAND_COLOURS, BANDS, BASE_URL, CLIENT_ID,
                                     CLIENT_SECRET, CONFIG, CONFIG_FILE, DEFAULT_CONFIG, GREEN, NAVY, OUTPUT_DIR,
                                     RED, REPORT_FILE, SOURCES, SOURCES_TO_BROWSE, TENANT_ID, TRUST_WEIGHTS,
                                     console, load_config)

# ---------------------------------------------------------------------------
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
# HELPERS: reshape API data for display
# ---------------------------------------------------------------------------
def make_progress():
    """A live progress display: spinner, step text, bar, percentage and elapsed time."""
    return Progress(SpinnerColumn(), TextColumn("[cyan]{task.description}"), BarColumn(),
                    TaskProgressColumn(), TimeElapsedColumn(), console=console)


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
# Dashboard hooks for companion scripts that customize report layout.
DASHBOARD_EXTRAS = None
CHART_DATA = {"col": 14, "row": 4, "heads": ["Dimension", "Yes", "No", "Coverage %", "Weight"], "title": "Chart data",
              "charts": True}
EXCEL_HIDE = {"Table Name"}


from purview_engine.report import (DIMENSION_COLOURS, add_dashboard, letters_for, score_colours,
                                    save_report_to_excel as build_excel_report)


def save_report_to_excel(rows):
    """Write the CDE and Trust Score sheets, honoring runtime dashboard customizations."""
    return build_excel_report(rows, excel_hide=EXCEL_HIDE, chart_data=CHART_DATA,
                              dashboard_extras=DASHBOARD_EXTRAS)


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

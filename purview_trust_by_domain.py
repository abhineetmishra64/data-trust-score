"""Purview Trust Score, built domain by domain (domain -> terms / CDEs -> tables and columns -> data quality).

Same screens and the same Trust Score as purview_trust_score.py, but the data comes from these endpoints:
  1  GET /businessdomains                                        domains
  2  GET /terms?domainId={id}                                    glossary terms of each domain
  3  GET /criticalDataElements                                   CDEs (matched to a domain by their "domain" id)
  4  GET /criticalDataElements/{id}/relationships?entityType=TERM         linked glossary terms
  5  GET /criticalDataElements/{id}/relationships?entityType=DATACOLUMN   linked columns, resolved with
     GET /dataColumns/{id} (-> table asset + Data Map column guid)
  6  GET /dataProducts?domainId={id}, GET /dataAssets            tables of the domain
  7  GET /datamap/api/atlas/v2/entity/guid/{guid}                real table and column names
  8  GET /datagovernance/quality/business-domains/{domainId}/asset-scores and snapshots (scored rules by column)
One row per CDE and linked column, with column-level rule scores grouped by DQ dimension.

Run:  python purview_trust_by_domain.py        (needs purview_trust_score.py in the same folder)
"""
import json
import sys
import time
from datetime import datetime

import requests
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from rich import box
from rich.markup import escape
from rich.panel import Panel
from rich.prompt import Prompt
from rich.table import Table

import purview_trust_score as base
from purview_trust_score import (BASE_URL, OUTPUT_DIR, TRUST_WEIGHTS, PurviewError, coverage_pct, console,
                                 expert_ids, explain_error, fetch_pages, item_name, make_progress,
                                 remove_html_tags, request_access_token, save_report_to_excel, show_error,
                                 show_scores, trust_score, write_item_files, yes_no)

# ---------------------------------------------------------------------------
# SETTINGS
# ---------------------------------------------------------------------------
ROOT = "https://api.purview-service.microsoft.com"
QUALITY = ROOT + "/datagovernance/quality/business-domains"
DATAMAP = ROOT + "/datamap/api/atlas/v2/entity/guid"
DQ_API_VERSION = "2026-01-12-preview"
EXPORT_PREFIX = "purview_export_domain_"
base.EXCEL_HIDE = set()  # this script has real tables: keep Table Name in the Excel CDE sheet
# Dashboard layout: the table-wise DQ grid sits under the cards (A13), the chart data beside it (I13), no summary charts
base.CHART_DATA.update({"col": 9, "row": 14, "heads": ["Dimension", "Trusted", "Not Trusted", "Coverage %", "Weight"],
                        "title": "Chart data", "bold_title": True, "charts": False})

MAIN_MENU = """[bold cyan]1[/] Load data from Purview {state}
[bold cyan]2[/] Trust Score
[bold cyan]3[/] Trust Score with data (adds a 'based on' column after each Y/N)
[bold cyan]4[/] Save Trust Score report to Excel (dashboard + sheets)
[bold cyan]5[/] Browse & export (domains, terms, CDEs, tables, columns, data quality)
[bold cyan]0[/] Exit"""


# ---------------------------------------------------------------------------
# API CALLS
# ---------------------------------------------------------------------------
def get_json(url, headers, label, params=None):
    """One GET returning the JSON body; any failure becomes a readable PurviewError."""
    try:
        response = requests.get(url, headers=headers, params=params, timeout=60)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as error:
        raise explain_error(error, label) from error
    except ValueError as error:
        raise PurviewError(f"{label}: response was not valid JSON.") from error


def api(path, extra=None):
    """Params for a catalog call: the api-version plus anything extra."""
    return {"api-version": base.API_VERSION, **(extra or {})}


def optional(call, default):
    """Run a call whose failure should not stop the load (e.g. 404 when nothing is linked)."""
    try:
        return call(), None
    except PurviewError as error:
        return default, str(error).splitlines()[0]


def find_score(body):
    """First number stored under a key that contains 'score' (None when there is none)."""
    if isinstance(body, dict):
        for key, value in body.items():
            if "score" in key.lower() and isinstance(value, (int, float)) and not isinstance(value, bool):
                return float(value)
        for value in body.values():
            found = find_score(value)
            if found is not None:
                return found
    elif isinstance(body, list):
        for value in body:
            found = find_score(value)
            if found is not None:
                return found
    return None


def fetch_raw(report):
    """Steps 1-8: fetch the real Purview responses (everything the Trust Score needs)."""
    report("Signing in to Microsoft Entra ID...")
    headers = {"Authorization": f"Bearer {request_access_token()}"}
    problems = []

    domains = fetch_pages("/businessdomains", headers, "domains", report)                       # step 1
    cdes = fetch_pages("/criticalDataElements", headers, "critical data elements", report)      # step 3
    terms, products = {}, {}
    for domain in domains:
        report(f"Glossary terms of {item_name(domain)}...")                                      # step 2
        for term in fetch_pages("/terms", headers, "glossary terms", extra_params={"domainId": domain["id"]}):
            terms[term["id"]] = term
        found, problem = optional(lambda: fetch_pages(                                           # step 6
            "/dataProducts", headers, "data products", extra_params={"domainId": domain["id"]}), [])
        products[domain["id"]] = found
        problems += [problem] if problem else []
    assets = fetch_pages("/dataAssets", headers, "data assets", report)                          # step 6

    cde_ids = [c["id"] for c in cdes if any(c.get("domain") == d["id"] for d in domains)]
    rel_term, rel_column = {}, {}
    for number, cde_id in enumerate(cde_ids, 1):
        report(f"Reading links of CDE {number}/{len(cde_ids)}")                                  # steps 4, 5
        for store, entity_type in ((rel_term, "TERM"), (rel_column, "DATACOLUMN")):
            store[cde_id], problem = optional(lambda: fetch_pages(
                f"/criticalDataElements/{cde_id}/relationships", headers, "relationships",
                extra_params={"entityType": entity_type}), [])
            problems += [f"{entity_type} links of {cde_id}: {problem}"] if problem else []

    # a CDE may be linked to a term that the domain list does not return (e.g. a Draft one): fetch it by id
    wanted = {link["entityId"] for links in rel_term.values() for link in links
              if str(link.get("relationshipType", "")).lower() == "related"}
    for term_id in sorted(wanted - set(terms)):
        report("Looking up a linked glossary term not in the list...")
        found, problem = optional(lambda: get_json(f"{BASE_URL}/terms/{term_id}", headers, "term",
                                                   api("")), None)
        if found:
            terms[term_id] = found
        else:
            problems.append(f"glossary term {term_id}: {problem}")

    columns, tables = resolve_columns(headers, assets, rel_column, report, problems)             # steps 5, 7
    quality = fetch_quality(headers, domains, cdes, rel_column, columns, assets, report, problems)  # step 8
    return {"domains": domains, "terms": list(terms.values()), "cdes": cdes, "dataProducts": products,
            "dataAssets": assets, "relationshipsTerm": rel_term, "relationshipsColumn": rel_column,
            "columns": columns, "dataMapTables": tables, "quality": quality, "problems": problems}


def resolve_columns(headers, assets, rel_column, report, problems):
    """Each linked column id -> its table (data asset) and real column name (Data Map)."""
    by_guid = {a.get("source", {}).get("assetId"): a for a in assets}
    columns, tables = {}, {}
    ids = sorted({link["entityId"] for links in rel_column.values() for link in links})
    for number, column_id in enumerate(ids, 1):
        report(f"Resolving columns {number}/{len(ids)}")
        info, problem = optional(lambda: get_json(f"{BASE_URL}/dataColumns/{column_id}", headers,
                                                  "data column", api("")), None)
        if not info:
            problems.append(f"column {column_id}: {problem}")
            continue
        column_info = info.get("value") or info
        source = column_info.get("source") or {}
        guid, column_guid = source.get("assetId"), source.get("columnId")
        asset = by_guid.get(guid, {})
        if guid and guid not in tables:  # step 7: the Data Map entity of the table lists its columns
            tables[guid], problem = optional(lambda: get_json(f"{DATAMAP}/{guid}", headers, "Data Map entity"), {})
            if problem:
                problems.append(f"Data Map {guid}: {problem}")
        entity = tables.get(guid) or {}
        column_entity = (entity.get("referredEntities") or {}).get(column_guid, {})
        columns[column_id] = {
            "columnId": column_id, "tableAssetId": (asset.get("source") or {}).get("assetId"), "dataMapGuid": guid,
            "table": asset.get("name") or (entity.get("entity", {}).get("attributes", {}).get("name")) or "(unknown table)",
            "column": (column_entity.get("attributes") or {}).get("name") or "(unknown column)",
            "qualifiedName": (column_entity.get("attributes") or {}).get("qualifiedName", "")}
    return columns, tables


def timeliness_map():
    """config.json "timeliness": {"<freshness view>": {"table": "<table>", "column": "<column>"}}.
    A freshness view (for example VW_CLAIM_FRESHNESS) carries a Timeliness rule; this says which table column
    that rule is about, because Purview has no lineage between the view and the table."""
    try:
        found = json.loads(base.CONFIG_FILE.read_text(encoding="utf-8")).get("timeliness", {})
    except (OSError, ValueError, AttributeError):
        return {}
    return {str(view).casefold(): spec for view, spec in found.items()
            if isinstance(spec, dict) and spec.get("table") and spec.get("column")}


def fetch_quality(headers, domains, cdes, rel_column, columns, assets, report, problems):
    """Fetch table scores and latest snapshots; snapshots contain the scored rules targeted at columns."""
    quality = {"available": False, "assets": {}}
    for domain in domains:
        asset_ids = sorted({columns[link["entityId"]]["tableAssetId"]
                            for c in cdes if c.get("domain") == domain["id"]
                            for link in rel_column.get(c["id"], []) if link["entityId"] in columns
                            and columns[link["entityId"]]["tableAssetId"]})
        # freshness views named in config.json are not linked from any CDE, but their Timeliness rule is needed
        views = {a["source"]["assetId"] for a in assets
                 if str(a.get("name", "")).casefold() in timeliness_map() and (a.get("source") or {}).get("assetId")}
        asset_ids = sorted(set(asset_ids) | views)
        for number, asset_id in enumerate(asset_ids, 1):
            report(f"Data quality {number}/{len(asset_ids)} of {item_name(domain)}")
            prefix = f"{QUALITY}/{domain['id']}"
            scores, problem = optional(lambda: get_json(f"{prefix}/asset-scores", headers, "asset scores",
                                                        {"api-version": DQ_API_VERSION,
                                                         "assetId": asset_id}), None)
            entry = {"domainId": domain["id"], "assetId": asset_id, "assetScores": scores,
                     "score": find_score(scores) if scores else None}
            if scores is None:
                entry["note"] = problem
            else:
                quality["available"] = True
                for part in ("rules", "runs"):
                    entry[part], _ = optional(lambda: get_json(
                        f"{prefix}/data-assets/{asset_id}/{part}", headers, part,
                        {"api-version": DQ_API_VERSION}), None)
            entry["snapshots"], snapshot_problem = optional(lambda: get_json(
                f"{prefix}/data-assets/{asset_id}/snapshots", headers, "snapshots",
                {"api-version": DQ_API_VERSION}), None)
            if entry["snapshots"] is not None:
                quality["available"] = True
            elif snapshot_problem:
                problems.append(f"snapshot for data asset {asset_id}: {snapshot_problem}")
            quality["assets"][f"{domain['id']}:{asset_id}"] = entry
    if not quality["available"] and quality["assets"]:
        problems.append("Data quality endpoints returned nothing (not available on this tenant): no column-level DQ results.")
    return quality


# ---------------------------------------------------------------------------
# SCORING: one row per CDE and linked column
# ---------------------------------------------------------------------------
def rows_from_raw(raw):
    """Rows for the Trust Score from raw responses (fresh from Purview or read back from a saved export)."""
    terms = {t["id"]: t for t in raw["terms"]}
    quality_by_asset = {e["assetId"]: e for e in raw["quality"]["assets"].values()}
    # every DQ dimension found in the rules (Completeness, Uniqueness, ...): one numeric column each
    dimensions = sorted({str(rule["dimension"]) for e in quality_by_asset.values()
                         for rule in ((e.get("snapshots") or {}).get("rules") or []) if rule.get("dimension")})
    rows = []
    for domain in raw["domains"]:
        for cde in raw["cdes"]:
            if cde.get("domain") != domain["id"]:
                continue
            related = [link["entityId"] for link in raw["relationshipsTerm"].get(cde["id"], [])
                       if str(link.get("relationshipType", "")).lower() == "related"]
            published = [terms[t] for t in related if t in terms
                         and str(terms[t].get("status", "")).lower() == "published"]
            experts = [e for t in published for e in expert_ids(t)]  # Ownership: published term has an expert
            people = raw.get("people", {})
            expert_names = list(dict.fromkeys(
                name for term in published for name in expert_names_from_term(term)))
            if not expert_names:  # names looked up from the directory (see resolve_people)
                expert_names = [people[e] for e in dict.fromkeys(experts) if e in people]
            description = remove_html_tags(cde.get("description"))
            linked = [raw["columns"][link["entityId"]] for link in raw["relationshipsColumn"].get(cde["id"], [])
                      if link["entityId"] in raw["columns"]] or [None]
            for column in linked:
                asset_quality = quality_by_asset.get(column["tableAssetId"], {}) if column else {}
                snapshot = asset_quality.get("snapshots") or {}
                column_name = str(column["column"]).strip().casefold() if column else ""
                column_rules = []
                snapshot_rules = list(snapshot.get("rules", []) or [])
                # Timeliness: the rule of a freshness view counts for the column named in config.json
                for view_name, spec in timeliness_map().items():
                    if column and column["table"].casefold() == str(spec["table"]).casefold() \
                            and column_name == str(spec["column"]).strip().casefold():
                        view_guid = next((a["source"]["assetId"] for a in raw["dataAssets"]
                                          if str(a.get("name", "")).casefold() == view_name), None)
                        view_snapshot = (quality_by_asset.get(view_guid) or {}).get("snapshots") or {}
                        for view_rule in view_snapshot.get("rules", []) or []:
                            if str(view_rule.get("dimension")) == "Timeliness":
                                # the rule has no column of its own here: mark it as belonging to this column
                                snapshot_rules.append({**view_rule, "typeProperties": {
                                    **(view_rule.get("typeProperties") or {}),
                                    "column": {"type": "Column", "value": column["column"]}}})
                for rule in snapshot_rules:
                    score = rule.get("score")
                    if not isinstance(score, (int, float)) or isinstance(score, bool) or score < 0:
                        continue
                    rule_columns = []
                    type_properties = rule.get("typeProperties") or {}
                    column_spec = type_properties.get("column") or {}
                    if isinstance(column_spec, dict) and column_spec.get("value") is not None:
                        rule_columns.append(str(column_spec["value"]))
                    for value in type_properties.get("columns", []) or []:
                        if isinstance(value, str):
                            rule_columns.append(value)
                        elif isinstance(value, dict):
                            name = value.get("value") or value.get("name") or value.get("columnName")
                            if name is not None:
                                rule_columns.append(str(name))
                    for value in rule.get("columns", []) or []:
                        if isinstance(value, str):
                            rule_columns.append(value)
                        elif isinstance(value, dict):
                            name = value.get("value") or value.get("name") or value.get("columnName")
                            if name is not None:
                                rule_columns.append(str(name))
                    if column_name and any(name.strip().casefold() == column_name for name in rule_columns):
                        column_rules.append(rule)
                row = {
                    "Domain": item_name(domain),
                    "Table Name": column["table"] if column else "(no column linked)",
                    "CDE ColumnName": item_name(cde),
                    "Column Name": column["column"] if column else "-",
                    "DQ Score": "Y" if column_rules else "N",
                    "Glossary": "Y" if published else "N",
                    "Ownership": "Y" if experts else "N",
                    "Metadata": "Y" if description else "N",
                    "DQ Rule Name": format_dq_rule_names(column_rules),
                    "DQ Dimension": format_dq_dimensions(column_rules),
                    "DQ Rule Score": format_dq_scores(column_rules),
                    "Domain ID": domain["id"], "CDE ID": cde["id"], "Glossary ID": ", ".join(related),
                    "Glossary Name": ", ".join(item_name(t) for t in published),
                    "Glossary Status": ", ".join(str(terms[t].get("status", "?")) for t in related if t in terms),
                    "Owner ID": ", ".join(experts), "Owner Name": ", ".join(expert_names),
                    "Description": description}
                for dimension in dimensions:  # average rule score (in %) of this column's rules in that dimension
                    scores = [rule["score"] * 100 for rule in column_rules if str(rule.get("dimension")) == dimension]
                    row[f"DQ {dimension} %"] = round(sum(scores) / len(scores), 1) if scores else ""
                rows.append(row)
    rows.sort(key=lambda r: (r["Domain"].casefold(), r["Table Name"].casefold()))  # stable: keeps CDE order per table
    return rows


def format_dq_rule_names(rules):
    """Names of the matched rules (for example 'Claim Number is unique'), in the same order as the dimensions."""
    return "; ".join(str(rule.get("name") or "Unnamed rule") for rule in rules)


def format_dq_dimensions(rules):
    """Format matched rule dimensions as a scalar, preserving rule order."""
    return "; ".join(str(rule.get("dimension") or "Unspecified") for rule in rules)


def format_dq_scores(rules):
    """Format matched rule scores as a scalar, aligned with the dimension order."""
    return "; ".join(f"{rule['score']:.1%}" for rule in rules)


def expert_names_from_term(term):
    """Return expert names included in Purview contact objects, without treating IDs as names."""
    experts = (term.get("contacts") or {}).get("expert")
    if isinstance(experts, (dict, str)):
        experts = [experts]
    names = []
    for expert in experts or []:
        if not isinstance(expert, dict):
            continue
        name = next((expert.get(key) for key in ("displayName", "name", "fullName", "userName")
                     if isinstance(expert.get(key), str) and expert[key].strip()), None)
        if name:
            names.append(name.strip())
    return names


def detail_for(row, dimension, short=True):
    """The data behind a Y/N: DQ score, linked term, experts, description, or why it is N."""
    if dimension == "DQ Score":
        names = row["DQ Rule Name"].split("; ")
        dimensions = row["DQ Dimension"].split("; ")
        scores = row["DQ Rule Score"].split("; ")
        if not row["DQ Rule Score"]:
            return "no column-level DQ rule result"
        # rule name (dimension): score
        return "; ".join(f"{rule} ({dimension}): {score}" for rule, dimension, score in zip(names, dimensions, scores))
    if dimension == "Glossary":
        if row["Glossary"] == "Y":
            return row["Glossary Name"]
        return (f"term {row['Glossary Status'] or 'not published'}" if row["Glossary ID"] else "no term linked")
    if dimension == "Ownership":
        experts = [e for e in row["Owner ID"].split(", ") if e]
        if not experts:
            return "no expert on term" if row["Glossary"] == "Y" else "no glossary term"
        names = [name for name in row.get("Owner Name", "").split(", ") if name]
        return ", ".join(names) if names else f"{len(experts)} expert(s); names not provided"
    text = row["Description"]
    if not text:
        return "no description"
    return text[:22] + "..." if short and len(text) > 22 else text[:200]


def dq_note(rows):
    """Reminder under the score when no column has a DQ rule result."""
    if any(r["DQ Score"] for r in rows):
        return "DQ Score: Y when a column has one or more scored DQ rules; details are grouped by DQ dimension."
    return ("[yellow]No column-level data quality results found (endpoint not available or no scored rules): "
            "DQ Score is N for all rows and "
            f"counts as 0%, so the score can reach at most {100 * (1 - TRUST_WEIGHTS['DQ Score']):.0f}.[/]")


# ---------------------------------------------------------------------------
# SCREENS
# ---------------------------------------------------------------------------
def show_cde_table(rows, with_data):
    """One line per CDE and linked column; with_data adds a 'based on' column after each Y/N."""
    table = Table(title=f"CDE coverage ({len(rows)})", box=box.ROUNDED, header_style="bold cyan")
    for column in ("Table Name", "CDE ColumnName", "Column Name"):
        table.add_column(column, overflow="fold")
    for dimension in TRUST_WEIGHTS:
        table.add_column(dimension)
        if with_data:
            table.add_column(f"{dimension.split()[0]} based on", overflow="fold", style="dim")
    for row in rows:
        cells = [escape(row["Table Name"]), escape(row["CDE ColumnName"]), escape(row["Column Name"])]
        for dimension in TRUST_WEIGHTS:
            cells.append(yes_no(row[dimension]))
            if with_data:
                cells.append(escape(detail_for(row, dimension)))
        table.add_row(*cells)
    console.print(table)


def show_dq_dimensions(rows):
    """Data quality by dimension: columns covered by a rule, average and lowest score."""
    dimensions = [k for k in (rows[0] if rows else {}) if k.startswith("DQ ") and k.endswith(" %")]
    if not dimensions:
        return
    table = Table(title="Data quality by dimension", box=box.ROUNDED, header_style="bold cyan")
    for column in ("Dimension", "Columns covered", "Average score %", "Lowest %"):
        table.add_column(column, justify="left" if column == "Dimension" else "right")
    for key in dimensions:
        values = [r[key] for r in rows if r[key] != ""]
        table.add_row(key[3:-2], str(len(values)), f"{sum(values) / len(values):.0f}" if values else "-",
                      f"{min(values):.0f}" if values else "-")
    console.print(table)


def show_result(rows, with_data):
    show_cde_table(rows, with_data)
    show_dq_dimensions(rows)
    show_scores(rows)  # scores per table and overall, plus the formula and bands
    console.print(dq_note(rows))


# --- owner names -----------------------------------------------------------------------------------
# Purview stores only ids for owners and experts, and the app itself may not read the directory. A signed-in user
# can read basic profiles (delegated User.ReadBasic.All): the first load offers a one-time sign-in with a code,
# then the names are kept in output/people_cache.json (names only, no token), so later loads do not ask again.
GRAPH_CLIENT = "14d82eec-204b-4c2f-b7e8-296a70dab67e"  # Microsoft Graph Command Line Tools (public client)
GRAPH_USERS = "https://graph.microsoft.com/v1.0/users"
PEOPLE_CACHE = OUTPUT_DIR / "people_cache.json"


def contact_ids(item):
    """Every person id in an object's contacts (owner, expert, ...): contacts may hold lists, dicts or plain ids."""
    found = set()
    for people in (item.get("contacts") or {}).values():
        for person in people if isinstance(people, list) else [people]:
            person_id = person.get("id") if isinstance(person, dict) else person
            if isinstance(person_id, str) and person_id.strip():
                found.add(person_id.strip())
    return found


def read_people_cache():
    try:
        return json.loads(PEOPLE_CACHE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def sign_in_for_names():
    """Device-code sign-in of the person running the tool; returns an access token for reading basic profiles."""
    endpoint = f"https://login.microsoftonline.com/{base.TENANT_ID}/oauth2/v2.0"
    try:
        start = requests.post(f"{endpoint}/devicecode", timeout=30, data={
            "client_id": GRAPH_CLIENT, "scope": "https://graph.microsoft.com/User.ReadBasic.All"})
        start.raise_for_status()
        flow = start.json()
        console.print(Panel(f"Open [bold]{flow['verification_uri']}[/] and enter the code [bold cyan]{flow['user_code']}"
                            f"[/]\nSign in with your work account (once). Waiting...", title="Sign in to read owner names",
                            border_style="cyan", expand=False))
        deadline = time.time() + int(flow.get("expires_in", 900))
        while time.time() < deadline:
            time.sleep(int(flow.get("interval", 5)))
            reply = requests.post(f"{endpoint}/token", timeout=30, data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code", "client_id": GRAPH_CLIENT,
                "device_code": flow["device_code"]}).json()
            if "access_token" in reply:
                return reply["access_token"]
            if reply.get("error") not in ("authorization_pending", "slow_down"):
                raise PurviewError(f"Sign-in failed: {str(reply.get('error_description', reply.get('error')))[:200]}")
    except requests.exceptions.RequestException as error:
        raise explain_error(error, "Sign-in for owner names") from error
    raise PurviewError("Sign-in for owner names timed out (the code lasts 15 minutes).")


def resolve_people(raw, ask=True):
    """id -> display name for every owner / expert in the data. Every load (menu option 1) shows the sign-in code and
    looks all names up again, even when there are no ids yet; if the sign-in fails the cached names are used."""
    ids = {i for group in (raw["domains"], raw["cdes"], raw["terms"], raw["dataAssets"])
           for item in group for i in contact_ids(item)}
    people = {i: name for i, name in read_people_cache().items() if i in ids}
    if not ask:
        return people
    console.print(f"[yellow]{len(ids)} owner/expert id(s) found ({len(people)} cached; Purview stores only ids).[/]")
    try:
        headers = {"Authorization": f"Bearer {sign_in_for_names()}"}
        for person_id in sorted(ids):
            response = requests.get(f"{GRAPH_USERS}/{person_id}", headers=headers, timeout=30,
                                    params={"$select": "displayName"})
            if response.ok and response.json().get("displayName"):
                people[person_id] = response.json()["displayName"]
        if people:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            PEOPLE_CACHE.write_text(json.dumps({**read_people_cache(), **people}, indent=2), encoding="utf-8")
        console.print(f"[green]✔ Found {len(people)} of {len(ids)} name(s)[/]")
    except PurviewError as error:
        show_error(error)
    return people


# --- Excel dashboard: data quality by dimension, table by table ---------------------------------------------
def draw_dq_grid(dash, data, last, table_range, table_names, dq_columns):
    """Dashboard section to the right of the two charts. One row per table, one column per DQ dimension:
    the average rule score of that table's columns (in %). Timeliness is always the last column and is a table
    check: 100% when the freshness rule passes (IS_FRESH is true), otherwise 0%; "-" when the table has no check."""
    navy, centre = "1F3864", Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    order = [n for n in dq_columns if n != "Timeliness"] + (["Timeliness"] if "Timeliness" in dq_columns else [])
    start_col = 1                                    # column A: under the score cards, left of the chart data (I-M)
    name_col = get_column_letter(start_col)
    title_row, header_row = 13, 14
    title = dash.cell(row=title_row, column=start_col, value="Data quality by dimension - table wise")
    title.font = Font(bold=True, size=12, color=navy)
    palette = ["2E75B6", "00A3A1", "8E7CC3", "C00000", "7F7F7F"]
    dash.merge_cells(start_row=header_row, start_column=start_col, end_row=header_row, end_column=start_col + 1)
    head = dash.cell(row=header_row, column=start_col, value="Domain")
    head.font, head.fill, head.alignment = Font(bold=True, color="FFFFFF", size=11), PatternFill("solid", fgColor=navy), centre
    for offset, name in enumerate(order):
        cell = dash.cell(row=header_row, column=start_col + 2 + offset, value=name)
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill("solid", fgColor="ED7D31" if name == "Timeliness" else palette[offset % len(palette)])
        cell.alignment = centre

    body = list(table_names) + ["All Domain"]
    for row_number, table_name in enumerate(body, header_row + 1):
        dash.merge_cells(start_row=row_number, start_column=start_col, end_row=row_number, end_column=start_col + 1)
        label = dash.cell(row=row_number, column=start_col, value=table_name)
        label.font = Font(bold=True, size=10, color="404040")
        label.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        everyone = table_name == "All Domain"
        for offset, name in enumerate(order):
            values = f"{data}!${dq_columns[name]}$2:${dq_columns[name]}${last}"
            mean = f"AVERAGE({values})" if everyone else f"AVERAGEIFS({values},{table_range},${name_col}{row_number})"
            cell = dash.cell(row=row_number, column=start_col + 2 + offset)
            if name == "Timeliness":                  # a table check: 100% or 0%
                cell.value = f'=IFERROR(IF({mean}>=100,100,0),"-")'
            else:
                cell.value = f'=IFERROR({mean},"-")'
            cell.number_format = '0"%"'
            cell.alignment = centre
            cell.font = Font(size=10, bold=everyone, color="404040")
        for column in range(start_col, start_col + 2 + len(order)):
            dash.cell(row=row_number, column=column).border = border
            dash.cell(row=row_number, column=column).fill = PatternFill("solid", fgColor="F2F2F2" if everyone else "FFFFFF")

    # colours follow the score bands (so 100% is green, 0% is red); a "-" stays plain
    good, ok = base.BANDS[1][0], base.BANDS[2][0]
    for offset in range(len(order)):
        letter = get_column_letter(start_col + 2 + offset)
        first_cell = f"{letter}{header_row + 1}"
        span = f"{first_cell}:{letter}{header_row + len(body)}"
        for formula, colour in ((f"AND(ISNUMBER({first_cell}),{first_cell}>={good})", base.GREEN),
                                (f"AND(ISNUMBER({first_cell}),{first_cell}>={ok},{first_cell}<{good})", base.AMBER),
                                (f"AND(ISNUMBER({first_cell}),{first_cell}<{ok})", base.RED)):
            dash.conditional_formatting.add(span, FormulaRule(formula=[formula], fill=PatternFill("solid", bgColor=colour)))
    note = dash.cell(row=header_row + len(body) + 2, column=start_col,
                     value="Timeliness is 100% when the freshness check passes (IS_FRESH = true), 0% when it fails, "
                           "- when the table has no check.")
    note.font = Font(italic=True, size=9, color="595959")


def draw_rule_chart(dash, data):
    """Bar chart under the two charts: one bar per CDE column and DQ rule (score in %). The pairs are listed on a
    'DQ Rules' sheet, read from the CDE coverage sheet (where several rules of a row are joined with '; ')."""
    source = dash.parent[data.strip("'").replace("''", "'")]  # `data` is quoted for formulas: 'CDE coverage'
    header = [c.value for c in source[1]]
    if not {"CDE ColumnName", "DQ Dimension", "DQ Rule Score"} <= set(header):
        return
    name_i, dim_i, score_i = (header.index(h) for h in ("CDE ColumnName", "DQ Dimension", "DQ Rule Score"))
    scores_by_label = {}
    for row in source.iter_rows(min_row=2, values_only=True):
        if not row[dim_i] or str(row[dim_i]).startswith("="):
            continue
        for dimension, score in zip(str(row[dim_i]).split("; "), str(row[score_i]).split("; ")):
            try:
                scores_by_label.setdefault(f"{row[name_i]} - {dimension}", set()).add(float(score.rstrip("%")))
            except ValueError:
                continue
    # one bar per CDE column and DQ dimension (average when several rules share a dimension)
    pairs = sorted(((label, round(sum(v) / len(v), 1)) for label, v in scores_by_label.items()),
                   key=lambda p: p[1])  # lowest score first: problems at the top
    if not pairs:
        return
    sheet = dash.parent.create_sheet("DQ Rules")
    sheet.append(["CDE ColumnName - DQ Dimension", "Score %"])
    for pair in pairs:
        sheet.append(list(pair))
    sheet.column_dimensions["A"].width = 60
    label_col, score_col, top = 1, 2, 1
    chart = BarChart()
    chart.type = "bar"
    chart.add_data(Reference(sheet, min_col=score_col, min_row=top, max_row=top + len(pairs)), titles_from_data=True)
    chart.set_categories(Reference(sheet, min_col=label_col, min_row=top + 1, max_row=top + len(pairs)))
    chart.title, chart.legend, chart.gapWidth = "DQ Dimension by CDE ColumnName", None, 50
    chart.title.overlay = False
    chart.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=CharacterProperties(sz=1400, b=True, solidFill="1F3864"))
    chart.x_axis.delete = chart.y_axis.delete = False
    chart.x_axis.scaling.orientation = "maxMin"      # first pair at the top
    chart.y_axis.crosses = "max"                     # ...with the % axis kept at the bottom
    chart.y_axis.scaling.min, chart.y_axis.scaling.max, chart.y_axis.majorUnit = 0, 100, 20
    chart.y_axis.number_format = '0'
    chart.y_axis.majorGridlines.spPr = GraphicalProperties(ln=None)
    chart.y_axis.majorGridlines.spPr.line.solidFill = "E7E6E6"
    chart.x_axis.graphicalProperties = GraphicalProperties()
    chart.x_axis.graphicalProperties.line.solidFill = "BFBFBF"
    small = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1000, solidFill="404040")),
                                  endParaRPr=CharacterProperties(sz=1000))])
    chart.x_axis.txPr = chart.y_axis.txPr = small
    chart.series[0].graphicalProperties.solidFill = "63BE7B"   # one green for every bar
    chart.series[0].graphicalProperties.line.solidFill = "63BE7B"
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showSerName = chart.dataLabels.showCatName = chart.dataLabels.showLegendKey = False
    chart.dataLabels.showPercent = False
    chart.dataLabels.numFmt = 'General'
    chart.dataLabels.position = "outEnd"
    chart.width, chart.height = 30, max(8, 3 + 0.75 * len(pairs))
    dash.add_chart(chart, "A24")


def draw_refresh_stamp(dash):
    """Top right (N1:P2): a small card with the time of the last refresh and its UTC offset."""
    now = datetime.now().astimezone()
    offset = now.strftime("%z")
    stamp = f"{now:%d %b %Y  %H:%M}  (UTC{offset[:3]}:{offset[3:]})"
    side = Side(style="thin", color="BFBFBF")
    for column in "NOP":
        dash.column_dimensions[column].width = 14
    dash.merge_cells("N1:P1")
    dash.merge_cells("N2:P2")
    dash["N1"], dash["N2"] = "Timezone (Last Refresh date)", stamp
    dash["N1"].font = Font(bold=True, size=9, color="FFFFFF")
    dash["N2"].font = Font(bold=True, size=10, color="1F3864")
    dash["N1"].fill = PatternFill("solid", fgColor="2F5597")
    dash["N2"].fill = PatternFill("solid", fgColor="F2F2F2")
    for row in (1, 2):
        for column in "NOP":
            cell = dash[f"{column}{row}"]
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=side if column == "N" else None, right=side if column == "P" else None,
                                 top=side if row == 1 else None, bottom=side if row == 2 else None)


def draw_dashboard_extras(dash, data, last, table_range, table_names, dq_columns):
    draw_refresh_stamp(dash)
    draw_dq_grid(dash, data, last, table_range, table_names, dq_columns)
    draw_rule_chart(dash, data)


base.DASHBOARD_EXTRAS = draw_dashboard_extras  # the shared Excel dashboard calls this to add the sections


def load_with_progress():
    """Main menu option 1: fetch everything once (live progress bar)."""
    with make_progress() as progress:
        task = progress.add_task("Starting...", total=None)
        raw = fetch_raw(lambda text: progress.update(task, description=text))
        progress.update(task, description="Done", total=1, completed=1)
    raw["people"] = resolve_people(raw)
    rows = rows_from_raw(raw)
    console.print(f"[green]✔ Loaded {len(raw['domains'])} domain(s), {len(raw['cdes'])} CDE(s), "
                  f"{len(raw['dataAssets'])} table(s) - {len(rows)} CDE/column row(s)[/]")
    for problem in raw["problems"][:5]:
        console.print(f"[yellow]  note: {escape(problem)}[/]")
    if len(raw["problems"]) > 5:
        console.print(f"[yellow]  ... and {len(raw['problems']) - 5} more[/]")
    console.print()
    return raw


# --- browse & export ------------------------------------------------------------------------------
def datasets(raw):
    """name -> (label, items): the lists that can be browsed and saved."""
    mapping = [{"name": f"{c['table']}.{c['column']}", "status": "", "columnId": c["columnId"], **c}
               for c in raw["columns"].values()]
    quality = [{"name": key, "status": "available" if e.get("assetScores") is not None
                or e.get("snapshots") is not None else "no result", **e}
               for key, e in raw["quality"]["assets"].items()]
    return {"domains": ("Business domains", raw["domains"]), "terms": ("Glossary terms", raw["terms"]),
            "cdes": ("Critical data elements", raw["cdes"]), "tables": ("Tables (data assets)", raw["dataAssets"]),
            "columns": ("CDE columns", mapping), "dq": ("Data quality results", quality)}


def dataset_menu(label, items, stem):
    while True:
        console.print(Panel("[bold cyan]1[/] List\n[bold cyan]2[/] Search by name\n[bold cyan]3[/] View details of one item\n"
                            "[bold cyan]4[/] Show full details of ALL items\n[bold cyan]5[/] Save to JSON + CSV + Excel\n"
                            "[bold cyan]0[/] Back", title=label, border_style="blue", expand=False))
        choice = Prompt.ask("Choose", choices=list("012345"), default="1")
        if choice == "0":
            return
        try:
            if not items:
                console.print("[yellow]Nothing here.[/]")
            elif choice == "1":
                base.print_items_table(items, label)
            elif choice == "2":
                base.search_items(items, label)
            elif choice == "3":
                base.print_item_details(items, label)
            elif choice == "4":
                base.print_all_item_details(items)
            else:
                base.save_items_to_files(items, stem)
        except Exception as error:
            show_error(error)
        console.print()


def export_all(state):
    """Save the real Purview responses in a NEW folder: raw.json (everything) + each list as JSON/CSV/Excel."""
    raw = state["raw"]
    if raw is None:
        raw = state["raw"] = load_with_progress()
    folder = OUTPUT_DIR / f"{EXPORT_PREFIX}{datetime.now():%Y%m%d_%H%M%S}"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "raw.json").write_text(json.dumps(raw, indent=2), encoding="utf-8")
    for key, (label, items) in datasets(raw).items():
        write_item_files(items, key, folder=folder)
    console.print(f"[green]✔ Saved the real Purview responses in {folder}[/]")


def trust_from_saved_export(state):
    """Build the Trust Score from a saved export (no Purview call)."""
    folders = sorted((p for p in OUTPUT_DIR.glob(f"{EXPORT_PREFIX}*") if (p / "raw.json").exists()), reverse=True)
    if not folders:
        raise PurviewError("No saved export found. Use 'Export ALL' first.")
    for number, folder in enumerate(folders, 1):
        console.print(f"[bold cyan]{number}[/] {folder.name}")
    pick = Prompt.ask("Use which export", choices=[str(i) for i in range(1, len(folders) + 1)], default="1")
    try:
        state["raw"] = json.loads((folders[int(pick) - 1] / "raw.json").read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise PurviewError(f"Could not read the export: {error}") from error
    show_result(rows_from_raw(state["raw"]), with_data=True)


def browse_menu(state):
    """Option 5: look at / save each dataset, export everything, or score a saved export."""
    while True:
        raw = state["raw"]
        sets = datasets(raw) if raw else {}
        names = list(sets) or ["domains", "terms", "cdes", "tables", "columns", "dq"]
        lines = [f"[bold cyan]{i}[/] {sets[k][0] if k in sets else k} [dim]({len(sets[k][1]) if k in sets else 'not loaded'})[/]"
                 for i, k in enumerate(names, 1)]
        lines += [f"[bold cyan]{len(names) + 1}[/] Export ALL: save the real Purview responses in a new folder",
                  f"[bold cyan]{len(names) + 2}[/] Trust Score from a saved export (no Purview call)",
                  "[bold cyan]0[/] Back to main menu"]
        console.print(Panel("\n".join(lines), title="Browse & export", border_style="blue", expand=False))
        pick = Prompt.ask("Choose", choices=[str(i) for i in range(0, len(names) + 3)], default="1")
        if pick == "0":
            return
        try:
            if int(pick) == len(names) + 1:
                export_all(state)
            elif int(pick) == len(names) + 2:
                trust_from_saved_export(state)
            elif raw is None:
                console.print("[yellow]No data loaded yet. Choose 1 (Load data) from the main menu first.[/]")
            else:
                key = names[int(pick) - 1]
                dataset_menu(sets[key][0], sets[key][1], key)
        except Exception as error:
            show_error(error)
        console.print()


# ---------------------------------------------------------------------------
# MAIN MENU
# ---------------------------------------------------------------------------
def main():
    console.print(Panel.fit("[bold cyan]Microsoft Purview Trust Score (by domain)[/]", border_style="cyan"))
    state = {"raw": None}
    while True:
        raw = state["raw"]
        status = "[dim](not loaded yet)[/]" if raw is None else (
            f"[dim](loaded {len(raw['cdes'])} CDEs, {len(raw['dataAssets'])} tables - choose again to refresh)[/]")
        console.print(Panel(MAIN_MENU.format(state=status), title="Main menu", border_style="blue", expand=False))
        choice = Prompt.ask("Choose", choices=list("012345"), default="1")
        if choice == "0":
            console.print("[cyan]Bye![/]")
            break
        try:
            if choice == "5":
                browse_menu(state)
            elif choice == "1":
                state["raw"] = load_with_progress()
            elif raw is None:
                console.print("[yellow]No data loaded yet. Choose 1 (Load data from Purview) first.[/]")
            else:
                rows = rows_from_raw(raw)
                if not rows:
                    console.print("[yellow]No critical data elements found for any domain.[/]")
                elif choice in ("2", "3"):
                    show_result(rows, with_data=choice == "3")
                else:
                    save_report_to_excel(rows)
                    console.print(f"[green]✔ Saved {base.REPORT_FILE.name} in {OUTPUT_DIR}[/]")
        except Exception as error:
            show_error(error)
        console.print()


if __name__ == "__main__":
    try:
        main()
    except (KeyboardInterrupt, EOFError):  # Ctrl+C: exit politely, no traceback
        console.print("\n[yellow]Bye...![/]")
        sys.exit(0)

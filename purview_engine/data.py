"""Transform Purview responses into Trust Score rows and scores."""
from rich.markup import escape

from .api import PurviewError, fetch_pages
from .helpers import item_name, remove_html_tags
from .settings import BANDS, SOURCES, TRUST_WEIGHTS


def contains_value(data, wanted):
    """True if wanted appears anywhere inside nested JSON as a value."""
    if isinstance(data, dict):
        return any(contains_value(value, wanted) for value in data.values())
    if isinstance(data, list):
        return any(contains_value(value, wanted) for value in data)
    return data == wanted


def owner_ids(item):
    """Return owner ids found in contacts.owner."""
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
    """Return expert ids found in contacts.expert."""
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
    """Return glossary ids whose relationship type is 'Related'."""
    ids = []
    for relationship in relationships:
        rel_type = str(relationship.get("relationshipType", "")).lower()
        term_id = relationship.get("entityId") or relationship.get("id")
        if term_id and rel_type == "related":
            ids.append(term_id)
    return ids


def fetch_raw(report):
    """Fetch raw domains, terms, CDEs, and relationships keyed by CDE id."""
    report("Signing in to Microsoft Entra ID...")
    headers = {"Authorization": f"******"}
    domains = fetch_pages(SOURCES["domains"]["path"], headers, "domains", report)
    cdes = fetch_pages(SOURCES["cdes"]["path"], headers, "critical data elements", report)
    terms = fetch_pages(SOURCES["terms"]["path"], headers, "glossary terms", report)
    pairs = [(domain, cde) for domain in domains for cde in cdes if contains_value(cde, domain.get("id"))]
    relationships = {}
    for number, (domain, cde) in enumerate(pairs, 1):
        report(f"Checking relationships {number}/{len(pairs)}: {item_name(cde)}")
        try:
            relationships[cde.get("id")] = fetch_pages(
                f"/criticalDataElements/{cde.get('id')}/relationships", headers, "relationships",
                extra_params={"entityType": "TERM"})
        except PurviewError:
            relationships[cde.get("id")] = []
    return {"domains": domains, "terms": terms, "cdes": cdes, "relationships": relationships}


def rows_from_raw(raw):
    """Convert raw responses into one Trust Score row per CDE."""
    terms_by_id = {term.get("id"): term for term in raw["terms"]}
    term_names = {term_id: item_name(term) for term_id, term in terms_by_id.items()}
    rows = []
    for domain in raw["domains"]:
        for cde in raw["cdes"]:
            if not contains_value(cde, domain.get("id")):
                continue
            glossary_ids = related_term_ids(raw["relationships"].get(cde.get("id"), []))
            found_ids = [term_id for term_id in glossary_ids if term_id in term_names]
            owners = [expert for term_id in found_ids for expert in expert_ids(terms_by_id[term_id])]
            description = remove_html_tags(cde.get("description"))
            rows.append({
                "Table Name": item_name(domain),
                "CDE ColumnName": item_name(cde),
                "DQ Score": "",
                "Glossary": "Y" if found_ids else "N",
                "Ownership": "Y" if owners else "N",
                "Metadata": "Y" if description else "N",
                "Domain ID": domain.get("id"),
                "CDE ID": cde.get("id"),
                "Glossary ID": ", ".join(glossary_ids),
                "Glossary Name": ", ".join(term_names[term_id] for term_id in found_ids),
                "Owner ID": ", ".join(owners),
                "Description": description,
            })
    return rows


def build_coverage_rows(report):
    """Fetch Purview data and convert it to Trust Score rows."""
    return rows_from_raw(fetch_raw(report))


def trust_score(rows):
    """Calculate the weighted share of CDEs marked Y; blank DQ counts as 0%."""
    return sum(weight * 100 * sum(row[name] == "Y" for row in rows) / len(rows)
               for name, weight in TRUST_WEIGHTS.items())


def coverage_pct(rows, dimension):
    """Return the share of rows marked Y for one dimension (0-100)."""
    return 100 * sum(row[dimension] == "Y" for row in rows) / len(rows) if rows else 0


def detail_for(row, dimension, short=True):
    """Describe the data behind a Y/N result."""
    if dimension == "DQ Score":
        return ("no source" if short else "no DQ source yet") if not row["DQ Score"] else str(row["DQ Score"])
    if dimension == "Glossary":
        if row["Glossary"] == "Y":
            return row["Glossary Name"]
        return "term not published" if row["Glossary ID"] else "no term linked"
    if dimension == "Ownership":
        owners = [owner for owner in row["Owner ID"].split(", ") if owner]
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
    """Format Y/N for Rich output and display empty values as a dash."""
    if value in ("Y", "N"):
        return f"[{'green' if value == 'Y' else 'red'}]{value}[/]"
    return escape(str(value)) or "-"


def band_of(score):
    """Return the score band label."""
    for minimum, label, _ in BANDS:
        if score >= minimum:
            return label
    return BANDS[-1][1]


def formula_text(separator=" + "):
    """Format the weighted Trust Score formula."""
    return separator.join(f"{weight:.0%} {name}" for name, weight in TRUST_WEIGHTS.items())


def band_formula(ref):
    """Build an Excel IF formula that turns a score cell into its band label."""
    expression = f'"{BANDS[-1][1]}"'
    for minimum, label, _ in reversed(BANDS[:-1]):
        expression = f'IF({ref}>={minimum},"{label}",{expression})'
    return "=" + expression

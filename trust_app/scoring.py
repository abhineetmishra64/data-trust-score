"""Every number the page shows, computed from the CDE rows (same rules as the Excel dashboard)."""
import purview_trust_score as base

ALL = "All Domain"
ROW_FIELDS = ("Table Name", "CDE ColumnName", "Column Name", "DQ Score", "Glossary", "Ownership", "Metadata",
              "DQ Rule Name", "DQ Dimension", "DQ Rule Score", "Glossary Name", "Owner Name", "Description")


def mean(values):
    return round(sum(values) / len(values), 1) if values else None


def dimension_cards(chosen):
    cards = []
    for name, weight in base.TRUST_WEIGHTS.items():
        yes = sum(r[name] == "Y" for r in chosen)
        cards.append({"name": name, "yes": yes, "no": len(chosen) - yes, "weight": weight,
                      "pct": round(100 * yes / len(chosen), 1) if chosen else 0})
    return cards


def quality_grid(rows, tables):
    """One row per table (plus the total): average DQ score per dimension. Timeliness is a 100% / 0% table check."""
    dims = sorted({k[3:-2] for k in (rows[0] if rows else {}) if k.startswith("DQ ") and k.endswith(" %")},
                  key=lambda d: (d == "Timeliness", d))

    def cell(group, dim):
        value = mean([r[f"DQ {dim} %"] for r in group if r[f"DQ {dim} %"] != ""])
        return (100 if value >= 100 else 0) if dim == "Timeliness" and value is not None else value

    grid = [{"table": t, "values": [cell([r for r in rows if r["Table Name"] == t], d) for d in dims]} for t in tables]
    grid.append({"table": ALL, "values": [cell(rows, d) for d in dims]})
    return dims, grid


def rule_scores(chosen):
    """One bar per CDE column and DQ dimension, lowest first (average when several rules share a dimension)."""
    found = {}
    for r in chosen:
        for dim, value in zip(r["DQ Dimension"].split("; "), r["DQ Rule Score"].split("; ")):
            if dim and value:
                found.setdefault(f"{r['CDE ColumnName']} - {dim}", set()).add(float(value.rstrip("%")))
    return sorted(({"label": k, "score": mean(list(v))} for k, v in found.items()), key=lambda x: x["score"])


def summarise(rows, table=ALL):
    tables = list(dict.fromkeys(r["Table Name"] for r in rows))
    chosen = rows if table == ALL else [r for r in rows if r["Table Name"] == table]
    cards = dimension_cards(chosen)
    score = round(sum(100 * c["yes"] / len(chosen) * c["weight"] for c in cards), 1) if chosen else 0  # unrounded inputs
    dims, grid = quality_grid(rows, tables)
    return {"tables": tables, "table": table, "score": score, "band": base.band_of(score), "cards": cards,
            "dims": dims, "grid": grid, "rules": rule_scores(chosen), "count": len(chosen),
            "bands": [{"min": m, "label": label, "colour": colour} for m, label, colour in base.BANDS],
            "rows": [{k: r[k] for k in ROW_FIELDS} for r in chosen]}

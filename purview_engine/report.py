"""Excel report generation for Trust Score workbooks."""
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .data import band_formula, formula_text
from .settings import AMBER, BANDS, GREEN, NAVY, OUTPUT_DIR, RED, REPORT_FILE, TRUST_WEIGHTS

DEFAULT_CHART_DATA = {"col": 14, "row": 4, "heads": ["Dimension", "Yes", "No", "Coverage %", "Weight"],
                      "title": "Chart data", "charts": True}
DIMENSION_COLOURS = {"DQ Score": "2E75B6", "Glossary": "00A3A1", "Ownership": "ED7D31", "Metadata": "8E7CC3"}


def score_colours(sheet, cell):
    """Colour a score cell by band: Trusted or better green, Use with caution amber, At risk red (updates live)."""
    good, ok = BANDS[1][0], BANDS[2][0]
    for operator, formula, colour in (("greaterThanOrEqual", [str(good)], GREEN),
                                      ("between", [str(ok), str(good - 0.001)], AMBER),
                                      ("lessThan", [str(ok)], RED)):
        sheet.conditional_formatting.add(cell, CellIsRule(
            operator=operator, formula=formula, fill=PatternFill("solid", bgColor=colour)))


def add_dashboard(workbook, letters, last, rows_info, table_col=None, table_names=None, dq_columns=None,
                  column_col=None, chart_data=None, dashboard_extras=None):
    """Add a 'Dashboard' sheet: a big Trust Score, one card per dimension, the score bands and three charts.
    Everything is a formula, so it refreshes whenever Y/N values change on the 'CDE coverage' sheet."""
    pct_row, weight_row, score_row = rows_info
    chart_data = DEFAULT_CHART_DATA if chart_data is None else chart_data
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
    c0, hr = chart_data["col"], chart_data["row"]
    cl = [get_column_letter(c0 + i) for i in range(5)]      # Dimension, Yes, No, Coverage %, Weight
    dash.cell(row=hr - 1, column=c0, value=chart_data["title"]).font = Font(
        bold=bool(chart_data.get("bold_title")), italic=not chart_data.get("bold_title"), size=12 if chart_data.get("bold_title") else 9,
        color=NAVY if chart_data.get("bold_title") else "808080")
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
    for index, text in enumerate(chart_data["heads"]):
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

    if chart_data["charts"]:
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
    if dashboard_extras and dq_columns and table_col and table_names:
        dashboard_extras(dash=dash, data=data, last=last, table_range=table_range, table_names=table_names,
                         dq_columns=dq_columns)

    workbook.move_sheet("Dashboard", offset=-1)  # make it the first tab
    workbook.active = 0


def letters_for(columns, name):
    """Excel column letter of a column name, or None when that column is not in the sheet."""
    return get_column_letter(columns.index(name) + 1) if name in columns else None


def save_report_to_excel(rows, excel_hide=None, chart_data=None, dashboard_extras=None):
    """Write the CDE rows plus the Trust Score section to an .xlsx file (sheets: Dashboard, CDE coverage).
    Y cells are green, N cells red. The scores are real Excel formulas, so filling in
    the DQ Score column by hand (Y/N) updates them automatically."""
    excel_hide = {"Table Name"} if excel_hide is None else excel_hide
    chart_data = DEFAULT_CHART_DATA if chart_data is None else chart_data
    OUTPUT_DIR.mkdir(exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "CDE coverage"
    columns = [c for c in rows[0] if c not in excel_hide]  # Table Name is hidden here: every CDE is in one domain
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
                  letters_for(columns, "Column Name"), chart_data, dashboard_extras)
    workbook.save(REPORT_FILE)

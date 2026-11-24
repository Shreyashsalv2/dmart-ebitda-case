"""Build the Excel model (model/DMart_EBITDA_Case.xlsx) with live formulas.

Inputs are the figures as filed (blue). Everything else is a formula, so a reviewer can trace
any number back to the filing and change an assumption to see the effect.
Run: python scripts/build_model.py
"""
from __future__ import annotations

import csv
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from case import DATA, ROOT, SCENARIOS, SENSITIVITY_BPS  # noqa: E402

OUT = ROOT / "model" / "DMart_EBITDA_Case.xlsx"

FONT = "Arial"
BLUE = Font(name=FONT, color="0000FF")
BLACK = Font(name=FONT, color="000000")
GREEN = Font(name=FONT, color="008000")
BOLD = Font(name=FONT, bold=True)
TITLE = Font(name=FONT, bold=True, size=14, color="1B4332")
SUB = Font(name=FONT, italic=True, size=9, color="555555")
HEAD = Font(name=FONT, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor="1B4332")
KEY_FILL = PatternFill("solid", fgColor="FFFF00")
TOTAL_FILL = PatternFill("solid", fgColor="E8F3EC")
THIN = Side(style="thin", color="999999")
TOP = Border(top=THIN)

CR = '#,##0.0;(#,##0.0);"-"'
CR2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.00%;(0.00%);"-"'
PCT1 = '0.0%;(0.0%);"-"'
BPS = '0;(0);"-"'
INT = '#,##0;(#,##0);"-"'
RAW = '#,##0;(#,##0);"-"'

SOURCES = {
    "nse_fy26": ("NSE integrated filing, standalone, year ended 31 Mar 2026 (INR lakh)",
                 "https://nsearchives.nseindia.com/corporate/ixbrl/INTEGRATED_FILING_INDAS_155301_02052026191318_iXBRL_WEB.html"),
    "nse_fy25": ("NSE integrated filing, standalone, year ended 31 Mar 2025 (INR)",
                 "https://nsearchives.nseindia.com/corporate/ixbrl/INTEGRATED_FILING_INDAS_89263_05052025110244_iXBRL_WEB.html"),
    "press_release_fy26": ("DMart press release on FY26 results, 2 May 2026",
                           "https://nsearchives.nseindia.com/corporate/DMART_02052026173007_ASLPressRelease02052026.pdf"),
    "nse_update_q4fy25": ("DMart company update, Q4 FY25, 3 Apr 2025",
                          "https://nsearchives.nseindia.com/corporate/DMART_03042025171933_ASLCompanyIntimationMarch2025.pdf"),
    "annual_report_fy25": ("DMart Annual Report FY2024-25, corporate overview",
                           "https://files.tijorifinance.com/insight/india/6857/Annual%20Report/AR-25.pdf"),
}
LABELS = {
    "revenue_from_operations": "Revenue from operations",
    "other_income": "Other income",
    "purchases_of_stock_in_trade": "Purchases of stock-in-trade",
    "changes_in_inventories": "Changes in inventories",
    "employee_benefit_expense": "Employee benefit expense",
    "finance_costs": "Finance costs",
    "depreciation_and_amortisation": "Depreciation and amortisation",
    "other_expenses": "Other expenses",
    "profit_before_tax": "Profit before tax (as reported)",
    "total_tax_expense": "Total tax expense",
    "profit_after_tax": "Profit after tax (as reported)",
}


def header(ws, row: int, values: list[str], start_col: int = 1) -> None:
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=start_col + i, value=v)
        c.font, c.fill = HEAD, HEAD_FILL
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)


def title(ws, text: str, sub: str) -> None:
    ws["A1"] = text
    ws["A1"].font = TITLE
    ws["A2"] = sub
    ws["A2"].font = SUB
    ws.sheet_view.showGridLines = False


def put(ws, ref: str, value, fmt: str | None = None, font: Font = BLACK, fill=None, bold=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=FONT, color=font.color, bold=bold or font.bold, italic=font.italic, size=font.size)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def widths(ws, cols: dict[str, float]) -> None:
    for k, v in cols.items():
        ws.column_dimensions[k].width = v


def build() -> Path:
    wb = Workbook()

    # ---------------- Cover ----------------
    cv = wb.active
    cv.title = "Cover"
    title(cv, "DMart (Avenue Supermarts) — EBITDA improvement case",
          "Standalone FY25 vs FY26 · INR crore unless stated · independent student case from public filings, not investment advice")
    rows = [
        ("Question", "Revenue grew 15.9% in FY26 but EBITDA margin stayed flat at about 7.85%. What held it back, and which levers could lift it?"),
        ("Answer (base case)", "='Levers'!D29"),
        ("Author", "Shreyash Singh Salv · github.com/Shreyashsalv2/dmart-ebitda-case"),
        ("Scope", "Standalone company (DMart stores). The DMart Ready e-commerce subsidiary is outside the standalone numbers."),
        ("Method", "Figures as filed (Inputs) → P&L and margins → EBITDA bridge → operating metrics → levers on the FY26 base → sensitivity → checks."),
        ("Caveat", "Lever sizes are the author's assumptions (yellow cells on Levers). They are not DMart guidance or forecasts."),
    ]
    for i, (k, v) in enumerate(rows, start=4):
        put(cv, f"A{i}", k, bold=True)
        c = put(cv, f"B{i}", v, font=GREEN if str(v).startswith("=") else BLACK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    put(cv, "A11", "Colour legend", bold=True)
    legend = [("Blue text", "figure as filed or an assumption you can change", BLUE, None),
              ("Black text", "formula", BLACK, None),
              ("Green text", "link to another sheet", GREEN, None),
              ("Yellow fill", "key assumption: change these on the Levers sheet", BLACK, KEY_FILL)]
    for i, (k, v, f, fill) in enumerate(legend, start=12):
        put(cv, f"A{i}", k, font=f, fill=fill)
        put(cv, f"B{i}", v)
    put(cv, "A17", "Sheets", bold=True)
    sheets = [("Inputs", "every figure as filed, with its source and the conversion to INR crore"),
              ("P&L", "profit and loss FY25 vs FY26, cost lines as % of revenue, change in bps"),
              ("Bridge", "what moved EBITDA from FY25 to FY26, with a waterfall chart"),
              ("Operations", "stores, retail area, revenue per sq ft and per store"),
              ("Levers", "three improvement levers under Cautious / Base / Stretch assumptions"),
              ("Sensitivity", "FY26 EBITDA margin if gross margin and employee cost move"),
              ("Checks", "the model reconciles to the reported figures")]
    for i, (k, v) in enumerate(sheets, start=18):
        put(cv, f"A{i}", k)
        put(cv, f"B{i}", v)
    put(cv, "A26", "Sources", bold=True)
    for i, (key, (desc, url)) in enumerate(SOURCES.items(), start=27):
        put(cv, f"A{i}", key)
        c = put(cv, f"B{i}", f"{desc}: {url}")
        c.hyperlink = url
    put(cv, "A33", "Overall check", bold=True)
    put(cv, "B33", '=IF(Checks!C15,"ALL CHECKS PASS","CHECK FAILED - see Checks sheet")', font=GREEN, bold=True)
    widths(cv, {"A": 22, "B": 120})
    cv["B5"].number_format = "@"

    # ---------------- Inputs ----------------
    ip = wb.create_sheet("Inputs")
    title(ip, "Inputs — figures as filed", "Blue = copied from the filing. Columns D-E convert to INR crore with the constants below.")
    put(ip, "A3", "Rupees per crore")
    put(ip, "B3", 10_000_000, INT, BLUE)
    put(ip, "A4", "Lakh per crore")
    put(ip, "B4", 100, INT, BLUE)
    header(ip, 6, ["Line item", "FY25 as filed (INR)", "FY26 as filed (INR lakh)", "FY25 (INR cr)", "FY26 (INR cr)", "Source FY25", "Source FY26"])
    fin_rows = {}
    with open(DATA / "financials.csv", newline="") as f:
        for r, row in enumerate(csv.DictReader(f), start=7):
            assert row["fy25_unit"] == "INR" and row["fy26_unit"] == "INR lakh"
            fin_rows[row["line"]] = r
            put(ip, f"A{r}", LABELS[row["line"]])
            put(ip, f"B{r}", float(row["fy25_reported"]), RAW, BLUE)
            put(ip, f"C{r}", float(row["fy26_reported"]), RAW, BLUE)
            put(ip, f"D{r}", f"=B{r}/$B$3", CR2)
            put(ip, f"E{r}", f"=C{r}/$B$4", CR2)
            put(ip, f"F{r}", row["fy25_source"])
            put(ip, f"G{r}", row["fy26_source"])
    ip[f"A{fin_rows['revenue_from_operations']}"].comment = Comment("Statement of profit and loss line 1. Standalone.", "Shreyash")
    op0 = max(fin_rows.values()) + 3
    header(ip, op0, ["Operating and reported figures", "FY25", "FY26", "Unit", "", "Source FY25", "Source FY26"])
    op_rows = {}
    with open(DATA / "operating.csv", newline="") as f:
        for r, row in enumerate(csv.DictReader(f), start=op0 + 1):
            op_rows[row["metric"]] = r
            scale = 0.01 if row["unit"] == "percent" else 1
            unit = "%" if row["unit"] == "percent" else row["unit"]
            fmt = PCT1 if row["unit"] == "percent" else ('0.0' if "sq ft" in row["unit"] else INT)
            put(ip, f"A{r}", row["metric"].replace("_", " ").capitalize())
            if row["fy25"]:
                put(ip, f"B{r}", float(row["fy25"]) * scale, fmt, BLUE)
            if row["fy26"]:
                put(ip, f"C{r}", float(row["fy26"]) * scale, fmt, BLUE)
            put(ip, f"D{r}", unit)
            put(ip, f"F{r}", row["fy25_source"])
            put(ip, f"G{r}", row["fy26_source"])
    note = op0 + len(op_rows) + 2
    put(ip, f"A{note}", "Source keys are explained on the Cover sheet and in data/SOURCES.md.", font=SUB)
    widths(ip, {"A": 38, "B": 22, "C": 22, "D": 16, "E": 16, "F": 20, "G": 20})

    def inp(line: str, year: str) -> str:
        return f"Inputs!{'D' if year == 'FY25' else 'E'}{fin_rows[line]}"

    def opn(metric: str, year: str) -> str:
        return f"Inputs!{'B' if year == 'FY25' else 'C'}{op_rows[metric]}"

    # ---------------- P&L ----------------
    pl = wb.create_sheet("P&L")
    title(pl, "Profit and loss — standalone", "INR crore. EBITDA = revenue − cost of goods sold − employee − other expenses (excludes other income), matching DMart's reported EBITDA.")
    header(pl, 4, ["INR crore", "FY25", "FY26", "Change", "Growth"])
    lines = [
        ("Revenue from operations", lambda y: f"={inp('revenue_from_operations', y)}", GREEN),
        ("Cost of goods sold", lambda y: f"={inp('purchases_of_stock_in_trade', y)}+{inp('changes_in_inventories', y)}", BLACK),
        ("Gross margin", lambda y: f"={y}5-{y}6", BLACK),
        ("Employee benefit expense", lambda y: f"={inp('employee_benefit_expense', y)}", GREEN),
        ("Other expenses", lambda y: f"={inp('other_expenses', y)}", GREEN),
        ("EBITDA", lambda y: f"={y}7-{y}8-{y}9", BLACK),
        ("Other income", lambda y: f"={inp('other_income', y)}", GREEN),
        ("Depreciation and amortisation", lambda y: f"={inp('depreciation_and_amortisation', y)}", GREEN),
        ("Finance costs", lambda y: f"={inp('finance_costs', y)}", GREEN),
        ("Profit before tax", lambda y: f"={y}10+{y}11-{y}12-{y}13", BLACK),
        ("Tax expense", lambda y: f"={inp('total_tax_expense', y)}", GREEN),
        ("Profit after tax", lambda y: f"={y}14-{y}15", BLACK),
    ]
    for i, (label, fn, colour) in enumerate(lines, start=5):
        bold = label in ("Gross margin", "EBITDA", "Profit after tax")
        put(pl, f"A{i}", label, bold=bold)
        for col, year in (("B", "FY25"), ("C", "FY26")):
            formula = fn(year).replace("FY25", "B").replace("FY26", "C") if "Inputs" not in fn(year) else fn(year)
            if "Inputs" not in formula:
                formula = fn(col)
            put(pl, f"{col}{i}", formula, CR, colour, bold=bold)
        put(pl, f"D{i}", f"=C{i}-B{i}", CR, bold=bold)
        put(pl, f"E{i}", f"=IF(B{i}=0,0,C{i}/B{i}-1)", PCT1, bold=bold)
        if bold:
            for col in "ABCDE":
                pl[f"{col}{i}"].fill = TOTAL_FILL
                pl[f"{col}{i}"].border = TOP
    header(pl, 19, ["% of revenue", "FY25", "FY26", "Change (bps)", ""])
    ratios = [("Gross margin %", 7), ("Employee cost %", 8), ("Other expenses %", 9), ("EBITDA margin %", 10), ("PAT margin %", 16)]
    for i, (label, src) in enumerate(ratios, start=20):
        put(pl, f"A{i}", label, bold=label.startswith("EBITDA"))
        for col in "BC":
            put(pl, f"{col}{i}", f"=IF({col}$5=0,0,{col}{src}/{col}$5)", PCT)
        put(pl, f"D{i}", f"=(C{i}-B{i})*10000", BPS)
    widths(pl, {"A": 32, "B": 14, "C": 14, "D": 14, "E": 12})

    # ---------------- Bridge ----------------
    br = wb.create_sheet("Bridge")
    title(br, "EBITDA bridge — FY25 to FY26", "INR crore. Revenue growth is valued at the FY25 margin; each rate effect is the change in that line's % of revenue × FY26 revenue. The steps add up exactly.")
    header(br, 4, ["Step", "INR crore", "Running total", "Chart: base", "Chart: increase", "Chart: decrease", "Chart: total"])
    steps = [
        ("FY25 EBITDA", "='P&L'!B10"),
        ("Revenue growth at FY25 margin", "=('P&L'!C5-'P&L'!B5)*'P&L'!B23"),
        ("Gross margin rate", "=('P&L'!C20-'P&L'!B20)*'P&L'!C5"),
        ("Employee cost rate", "=-('P&L'!C21-'P&L'!B21)*'P&L'!C5"),
        ("Other expenses rate", "=-('P&L'!C22-'P&L'!B22)*'P&L'!C5"),
        ("FY26 EBITDA", "=SUM(B5:B9)"),
    ]
    for i, (label, formula) in enumerate(steps, start=5):
        total = label.endswith("EBITDA")
        put(br, f"A{i}", label, bold=total)
        put(br, f"B{i}", formula, CR, GREEN if "P&L" in formula and total else BLACK, bold=total)
        put(br, f"C{i}", "=B5" if i == 5 else (f"=B{i}" if i == 10 else f"=C{i-1}+B{i}"), CR)
        if total:
            put(br, f"D{i}", 0, CR)
            put(br, f"E{i}", 0, CR)
            put(br, f"F{i}", 0, CR)
            put(br, f"G{i}", f"=B{i}", CR)
        else:
            put(br, f"D{i}", f"=MIN(C{i-1},C{i})", CR)
            put(br, f"E{i}", f"=MAX(B{i},0)", CR)
            put(br, f"F{i}", f"=MAX(-B{i},0)", CR)
            put(br, f"G{i}", 0, CR)
    put(br, "A12", "Check: bridge total equals FY26 EBITDA on P&L", bold=True)
    put(br, "B12", "=ABS(B10-'P&L'!C10)<0.01")
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "EBITDA bridge FY25 → FY26 (INR crore)"
    chart.y_axis.title = "INR crore (axis starts at 4,000)"
    data = Reference(br, min_col=4, max_col=7, min_row=4, max_row=10)
    cats = Reference(br, min_col=1, min_row=5, max_row=10)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    colours = [None, "2D6A4F", "C0392B", "1B4332"]
    for s, colour in zip(chart.series, colours):
        if colour is None:
            s.graphicalProperties.noFill = True
            s.graphicalProperties.line.noFill = True
        else:
            s.graphicalProperties.solidFill = colour
            s.graphicalProperties.line.solidFill = colour
    chart.y_axis.scaling.min = 4000
    chart.legend = None
    chart.roundedCorners = False
    chart.title.overlay = False
    chart.y_axis.title.overlay = False
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.height, chart.width = 7.5, 20
    chart.y_axis.number_format = "#,##0"
    br.add_chart(chart, "A15")
    widths(br, {"A": 46, "B": 14, "C": 14, "D": 13, "E": 14, "F": 14, "G": 12})

    # ---------------- Operations ----------------
    ops = wb.create_sheet("Operations")
    title(ops, "Operations — stores and space", "Revenue per sq ft here uses YEAR-END retail area (the author's measure); DMart reports its own figure on weighted-average area.")
    header(ops, 4, ["Metric", "FY25", "FY26", "Change", "Unit"])
    orows = [
        ("Stores at year end", f"={opn('stores_at_year_end', 'FY25')}", f"={opn('stores_at_year_end', 'FY26')}", INT, "stores", "pct"),
        ("Retail business area at year end", f"={opn('retail_business_area_at_year_end', 'FY25')}", f"={opn('retail_business_area_at_year_end', 'FY26')}", "0.0", "million sq ft", "pct"),
        ("Stores added in FY26", None, f"={opn('stores_added_in_year', 'FY26')}", INT, "stores", None),
        ("  of which opened in Q4 FY26", None, f"={opn('stores_added_in_q4', 'FY26')}", INT, "stores", None),
        ("  share of FY26 openings in Q4", None, "=IF(C7=0,0,C8/C7)", PCT1, "%", None),
        ("Revenue per sq ft of year-end area", "='P&L'!B5*Inputs!$B$3/(B6*1000000)", "='P&L'!C5*Inputs!$B$3/(C6*1000000)", INT, "INR per sq ft", "pct"),
        ("Revenue per store (year-end count)", "='P&L'!B5/B5", "='P&L'!C5/C5", CR2, "INR crore", "pct"),
        ("Employee cost per store (year-end count)", "='P&L'!B8/B5", "='P&L'!C8/C5", CR2, "INR crore", "pct"),
        ("Revenue growth", None, "='P&L'!E5", PCT1, "%", None),
        ("Employee cost growth", None, "='P&L'!E8", PCT1, "%", None),
    ]
    for i, (label, f25, f26, fmt, unit, chg) in enumerate(orows, start=5):
        put(ops, f"A{i}", label)
        if f25:
            put(ops, f"B{i}", f25, fmt, GREEN if "!" in f25 else BLACK)
        put(ops, f"C{i}", f26, fmt, GREEN if "!" in f26 and "P&L'!C5*" not in f26 else BLACK)
        if chg:
            put(ops, f"D{i}", f"=IF(B{i}=0,0,C{i}/B{i}-1)", PCT1)
        put(ops, f"E{i}", unit)
    widths(ops, {"A": 42, "B": 14, "C": 14, "D": 12, "E": 16})

    # ---------------- Levers ----------------
    lv = wb.create_sheet("Levers")
    title(lv, "Improvement levers on the FY26 base", "Yellow cells are the author's assumptions. Results are FY26 pro-forma (what FY26 would have looked like), not a forecast.")
    header(lv, 4, ["Assumption", "Unit", "Cautious", "Base", "Stretch", "Rationale"])
    assumptions = [
        ("Employee cost recovered", "bps of revenue", "emp_bps", "Employee cost rose 27 bps of revenue in FY26 (+33.5% vs revenue +15.9%), as new stores were staffed ahead of their sales."),
        ("Gross margin gain from mix", "bps of revenue", "gm_bps", "Shift in mix toward higher-margin general merchandise and apparel. Must not break everyday-low-price positioning."),
        ("Share of new-store density gap closed", "share", "density_recovery", "Revenue per sq ft of year-end area fell 3.2%: 58 of 85 new stores opened in Q4 and had little time to trade."),
        ("Variable share of employee + other costs", "share", "variable_opex_share", "Extra revenue from existing space carries only the variable part of operating costs."),
    ]
    for i, (label, unit, key, why) in enumerate(assumptions, start=5):
        put(lv, f"A{i}", label)
        put(lv, f"B{i}", unit)
        for col, name in zip("CDE", SCENARIOS):
            val = SCENARIOS[name][key]
            put(lv, f"{col}{i}", val, BPS if "bps" in unit else '0%', BLUE, fill=KEY_FILL)
        c = put(lv, f"F{i}", why)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        lv.row_dimensions[i].height = 30
    header(lv, 10, ["Calculation (INR crore unless stated)", "Unit", "Cautious", "Base", "Stretch", "How"])
    calc = [
        (11, "FY26 revenue", "INR crore", "='P&L'!$C$5", CR, "P&L"),
        (12, "FY26 EBITDA", "INR crore", "='P&L'!$C$10", CR, "P&L"),
        (13, "FY26 EBITDA margin", "%", "='P&L'!$C$23", PCT, "P&L"),
        (14, "Lever 1: employee productivity", "INR crore", "={c}5/10000*{c}11", CR, "bps recovered × FY26 revenue"),
        (15, "Lever 2: gross margin mix", "INR crore", "={c}6/10000*{c}11", CR, "bps gained × FY26 revenue"),
        (16, "Revenue per sq ft FY25 (year-end area)", "INR per sq ft", "=Operations!$B$10", INT, "Operations"),
        (17, "Revenue per sq ft FY26 (year-end area)", "INR per sq ft", "=Operations!$C$10", INT, "Operations"),
        (18, "Density gap", "INR per sq ft", "={c}16-{c}17", INT, "FY25 − FY26"),
        (19, "FY26 retail area", "million sq ft", "=Operations!$C$6", "0.0", "Operations"),
        (20, "Extra revenue if gap partly closed", "INR crore", "={c}7*{c}18*{c}19*1000000/Inputs!$B$3", CR, "share closed × gap × area"),
        (21, "Flow-through to EBITDA", "%", "='P&L'!$C$20-{c}8*('P&L'!$C$21+'P&L'!$C$22)", PCT, "gross margin % − variable share × (employee % + other %)"),
        (22, "Lever 3: new-store ramp", "INR crore", "={c}20*{c}21", CR, "extra revenue × flow-through"),
        (23, "Total lever impact", "INR crore", "={c}14+{c}15+{c}22", CR, "sum of levers"),
        (24, "Pro-forma FY26 revenue", "INR crore", "={c}11+{c}20", CR, ""),
        (25, "Pro-forma FY26 EBITDA", "INR crore", "={c}12+{c}23", CR, ""),
        (26, "Pro-forma EBITDA margin", "%", "=IF({c}24=0,0,{c}25/{c}24)", PCT, ""),
        (27, "Margin uplift", "bps", "=({c}26-{c}13)*10000", BPS, ""),
    ]
    for r, label, unit, f, fmt, how in calc:
        total = r in (23, 25, 26, 27)
        put(lv, f"A{r}", label, bold=total)
        put(lv, f"B{r}", unit)
        for col in "CDE":
            formula = f.format(c=col)
            put(lv, f"{col}{r}", formula, fmt, GREEN if ("!" in formula and "{c}" not in f) else BLACK, bold=total)
            if total:
                lv[f"{col}{r}"].fill = TOTAL_FILL
        put(lv, f"F{r}", how, font=SUB)
    put(lv, "A29", "Base case in one line", bold=True)
    put(lv, "D29", '="Three levers worth about INR "&TEXT(D23,"#,##0")&" crore would lift FY26 EBITDA margin from "&TEXT(D13,"0.00%")&" to "&TEXT(D26,"0.00%")&" (+"&TEXT(D27,"0")&" bps)."')
    widths(lv, {"A": 40, "B": 14, "C": 13, "D": 13, "E": 13, "F": 70})

    # ---------------- Sensitivity ----------------
    se = wb.create_sheet("Sensitivity")
    title(se, "Sensitivity — FY26 EBITDA margin", "Rows: change in gross margin (bps of revenue). Columns: change in employee cost (bps of revenue; positive = higher cost).")
    put(se, "A4", "Gross margin Δ (bps) ↓ / Employee cost Δ (bps) →", bold=True)
    for j, bps in enumerate(SENSITIVITY_BPS):
        col = get_column_letter(2 + j)
        put(se, f"{col}4", bps, BPS, BLUE, bold=True)
    for i, bps in enumerate(SENSITIVITY_BPS):
        r = 5 + i
        put(se, f"A{r}", bps, BPS, BLUE, bold=True)
        for j in range(len(SENSITIVITY_BPS)):
            col = get_column_letter(2 + j)
            put(se, f"{col}{r}", f"='P&L'!$C$23+$A{r}/10000-{col}$4/10000", PCT)
            if bps == 0 and SENSITIVITY_BPS[j] == 0:
                se[f"{col}{r}"].fill = KEY_FILL
    put(se, f"A{6 + len(SENSITIVITY_BPS)}", "Each 10 bps of revenue is worth about INR", font=SUB)
    put(se, f"B{6 + len(SENSITIVITY_BPS)}", "='P&L'!C5*10/10000", CR, GREEN)
    put(se, f"C{6 + len(SENSITIVITY_BPS)}", "crore of EBITDA on the FY26 base.", font=SUB)
    widths(se, {"A": 44, **{get_column_letter(2 + j): 11 for j in range(len(SENSITIVITY_BPS))}})

    # ---------------- Checks ----------------
    ck = wb.create_sheet("Checks")
    title(ck, "Checks — the model reconciles to what DMart reported", "TRUE means the check passes.")
    header(ck, 4, ["Check", "Model", "Pass"])
    checks = [
        ("FY26 EBITDA rounds to the reported INR 5,255 crore", "='P&L'!C10", f"=ROUND(B5,0)={opn('reported_ebitda', 'FY26')}", CR),
        ("FY26 EBITDA growth rounds to the reported 15.7%", "='P&L'!E10", f"=ROUND(B6,3)={opn('reported_ebitda_growth', 'FY26')}", PCT1),
        ("FY25 EBITDA margin rounds to the reported 7.9%", "='P&L'!B23", f"=ROUND(B7,3)={opn('reported_ebitda_margin', 'FY25')}", PCT),
        ("FY26 EBITDA margin rounds to the reported 7.8%", "='P&L'!C23", f"=ROUND(B8,3)={opn('reported_ebitda_margin', 'FY26')}", PCT),
        ("FY25 profit before tax equals the filing", "='P&L'!B14", f"=ABS(B9-{inp('profit_before_tax', 'FY25')})<0.01", CR2),
        ("FY26 profit before tax equals the filing", "='P&L'!C14", f"=ABS(B10-{inp('profit_before_tax', 'FY26')})<0.01", CR2),
        ("FY25 profit after tax equals the filing", "='P&L'!B16", f"=ABS(B11-{inp('profit_after_tax', 'FY25')})<0.01", CR2),
        ("FY26 profit after tax equals the filing", "='P&L'!C16", f"=ABS(B12-{inp('profit_after_tax', 'FY26')})<0.01", CR2),
        ("Bridge adds up to FY26 EBITDA", "=Bridge!B10", "=Bridge!B12", CR),
        ("Base-case levers add up", "=Levers!D23", "=ABS(B14-(Levers!D14+Levers!D15+Levers!D22))<0.01", CR),
    ]
    for i, (label, model, test, fmt) in enumerate(checks, start=5):
        put(ck, f"A{i}", label)
        put(ck, f"B{i}", model, fmt, GREEN)
        put(ck, f"C{i}", test)
    put(ck, "A15", "All checks pass", bold=True)
    put(ck, "C15", "=AND(C5:C14)", bold=True)
    ck["C15"].fill = TOTAL_FILL
    widths(ck, {"A": 52, "B": 14, "C": 10})

    for ws in wb.worksheets:
        ws.freeze_panes = "A5" if ws.title not in ("Cover",) else None
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())

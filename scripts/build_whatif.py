#!/usr/bin/env python3
"""
Builds VPT_WhatIf_Analysis.xlsx from config/parameters.json + data/findings.json.

Sheets
  00_ReadMe              how the what-if works and how to change it
  01_Inputs              every parameter, grouped, with range, basis, owner and the KPIs it governs
  02_Live_Calculator     formula-driven calculator: change an input, see the outputs recalculate in Excel
  03_Scenarios           scenario definitions and the levers each one moves
  04_Scenario_Summary    KPI outputs per scenario against the baseline, with deltas and flags
  05_Scenario_Detail     long-format detail: scenario x output x baseline x scenario x delta
  06_Sensitivity         one-at-a-time tornado analysis
  07_BreakEven           break-even solves and the cover-policy trade-off curve
  08_TwoWay              demand growth x price inflation grid for funding gap and utilisation
  09_Lever_Catalogue     every lever: what it does, working range, owner, effect
  10_Model_Logic         the equations, in one place
  11_Refresh             how to recalculate after a data or parameter change
"""
import os
import json

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "..", "data", "whatif_results.json")) as f:
    W = json.load(f)
with open(os.path.join(BASE, "..", "config", "parameters.json")) as f:
    CFG = json.load(f)

OUT = os.path.join(BASE, "..", "data", "VPT_WhatIf_Analysis.xlsx")

NAVY = "1F3B57"
TEAL = "2A7F8F"
LIGHT = "EEF3F8"
GREY = "595959"
GREEN = "E2EFDA"
AMBER = "FCE4D6"
RED = "F8CBAD"
BLUE = "DDEBF7"
INPUT_FILL = "FFF2CC"

H_FILL = PatternFill("solid", fgColor=NAVY)
H_FONT = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=9.5)
TITLE = Font(name="Calibri", size=15, bold=True, color=NAVY)
SUB = Font(name="Calibri", size=9, italic=True, color=GREY)
THIN = Side(style="thin", color="DDDDDD")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = '#,##0'
PCT = '0.0%'
NUM = '#,##0'
RATIO = '0.00'

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, subtitle=None, widths=None, tab=NAVY):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE
    if subtitle:
        c = ws.cell(row=2, column=1, value=subtitle)
        c.font = SUB
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(6, min(12, len(widths or [1] * 8))))
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_properties.tabColor = tab
    return ws, 4


def header(ws, row, headers, height=32):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.row_dimensions[row].height = height


def put(ws, row, values, wrap=(), bold=False, fill=None, fmt=None, height=None):
    for c, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = Font(name="Calibri", size=9.5, bold=bold)
        cell.border = BOX
        cell.alignment = Alignment(wrap_text=(c in wrap), vertical="top")
        if fmt and c in fmt and isinstance(v, (int, float)):
            cell.number_format = fmt[c]
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)
    if height:
        ws.row_dimensions[row].height = height


def note(ws, row, text):
    c = ws.cell(row=row, column=1, value=text)
    c.font = SUB
    return row + 1


B = W["baseline_result"]
BP = W["baseline_params"]
FACTS = W["baseline_facts"]

# ============================================================ 00 README
ws, r = sheet("00_ReadMe", "What-if & Scenario Analysis – Consumption, Purchasing and Budget",
              f"Parameter set v{W['config_version']}  |  data as at {W['as_at']}  |  every input lives in config/parameters.json",
              widths=[3, 34, 104], tab=NAVY)
r = 4
ws.cell(row=r, column=2, value="VITAL PAKISTAN TRUST").font = Font(name="Calibri", size=16, bold=True, color=NAVY)
r += 1
ws.cell(row=r, column=2, value="What-if analysis for the Consumption Analysis dashboard").font = Font(
    name="Calibri", size=11, color=GREY)
r += 2
guide = [
    ("What this workbook answers",
     "“If demand, prices, supply, policy or funding change, what happens to consumption, purchasing, stock, expiry risk, "
     "stock-outs, budget and cost per beneficiary?” Each scenario is measured against the baseline case."),
    ("Horizon",
     "Forward 12 months at the current operating level. The baseline consumption is the annualised run rate inflated for the "
     "approved demand and price assumptions, so it differs from the historic period totals in the analysis workbook."),
    ("Everything is parameterised",
     f"All {len(CFG['model'])} model assumptions and all {len(CFG['thresholds'])} thresholds live in config/parameters.json. "
     "The workbook is generated from that file plus the current data, so no value is typed twice and none is hard-coded."),
    ("How to run a new what-if",
     "1) Edit the parameter or add a scenario in config/parameters.json.  2) Run scripts/refresh_all.py (or "
     "scripts/whatif_model.py then scripts/build_whatif.py).  3) The scenario summary, sensitivity, break-even and two-way "
     "tables recalculate.  4) For a quick manual check, use 02_Live_Calculator, which recalculates inside Excel."),
    ("Live calculator",
     "02_Live_Calculator is formula-driven: change a yellow input cell and the outputs recalculate without re-running anything. "
     "It reproduces the core equations, so it is also the simplest way to explain the model to non-analysts."),
    ("Reading the scenarios",
     "S0 is the baseline. S1–S4 are single-shock cases, S5 and S6 are policy / commercial levers, S7 is the combined downside "
     "and S8 is the same downside with the available levers applied — the gap between S7 and S8 is the value of acting."),
    ("Sensitivity",
     "06_Sensitivity moves one lever at a time between a low and a high value and ranks the levers by the size of their effect "
     "(tornado). Use it to answer “what should we watch most closely?”."),
    ("Break-even",
     "07_BreakEven answers “how far can this move before we have a funding problem?” and shows the trade-off between stock cover, "
     "purchase cost and expiry risk."),
    ("Two-way table",
     "08_TwoWay shows funding gap and budget utilisation for every combination of demand growth and price inflation, so a reader "
     "can locate the current position and see which direction hurts most."),
    ("Caveat",
     "The model is deliberately transparent and linear in its drivers; it is a decision aid, not a forecast. Assumptions are "
     "stated in 10_Model_Logic and should be challenged at each quarterly review."),
]
for label, text in guide:
    c = ws.cell(row=r, column=2, value=label)
    c.font = Font(name="Calibri", size=10, bold=True)
    c.alignment = Alignment(vertical="top")
    t = ws.cell(row=r, column=3, value=text)
    t.font = Font(name="Calibri", size=10)
    t.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 14 * (len(text) // 100 + 1)
    r += 1

r += 1
ws.cell(row=r, column=2, value="Baseline position (forward 12 months)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
r += 1
for label, val in [
    ("Projected consumption value", f"PKR {B['consumption_value']:,.0f}"),
    ("Projected purchase requirement", f"PKR {B['purchase_value']:,.0f}"),
    ("Consumption-to-purchase ratio", f"{B['ratio']:.2f}"),
    ("Budget available", f"PKR {B['budget_available']:,.0f}"),
    ("Budget utilisation", f"{B['budget_utilisation']:.1%}"),
    ("Funding gap", f"PKR {B['funding_gap']:,.0f}"),
    ("Expiry risk value", f"PKR {B['expiry_risk_value']:,.0f}"),
    ("Stock-out risk items", f"{B['stockout_items']}"),
    ("Working capital in stock", f"PKR {B['working_capital']:,.0f}"),
    ("Cost per beneficiary", f"PKR {B['cost_per_beneficiary']:,.0f}"),
]:
    ws.cell(row=r, column=2, value=label).font = BODY
    ws.cell(row=r, column=3, value=val).font = Font(name="Calibri", size=10, bold=True)
    r += 1

# ============================================================ 01 INPUTS
ws, r = sheet("01_Inputs", "All parameters – the single source of truth",
              "Grouped by driver. Yellow cells are the live inputs of the workbook model; the same values are held in config/parameters.json.",
              widths=[26, 42, 12, 10, 12, 12, 46, 20, 44], tab=TEAL)
header(ws, r + 1, ["Parameter key", "Parameter", "Current value", "Unit", "Min", "Max",
                   "Basis / source", "Owner", "Governs (KPIs affected)"])
rr = r + 2
group_row = {}
last_group = None
for spec in CFG["model"]:
    g = spec["group"]
    if g != last_group:
        c = ws.cell(row=rr, column=1, value=g.upper())
        c.font = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=TEAL)
        for cc in range(1, 10):
            ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=TEAL)
        rr += 1
        last_group = g
    v = BP.get(spec["key"], spec["value"])
    fmt = {}
    if spec["unit"] == "%":
        fmt = {3: PCT}
    elif spec["unit"] == "PKR":
        fmt = {3: MONEY}
    elif spec["unit"] == "x" or spec["unit"] == "Ratio":
        fmt = {3: '0.00'}
    put(ws, rr, [spec["key"], spec["name"], v, spec["unit"], spec["min"], spec["max"],
                 spec["basis"], spec["owner"], spec["governs"]], wrap=(2, 7, 9), fmt=fmt)
    ws.cell(row=rr, column=3).fill = PatternFill("solid", fgColor=INPUT_FILL)
    group_row[spec["key"]] = rr
    rr += 1

rr += 1
ws.cell(row=rr, column=1, value="THRESHOLDS (status bands)").font = Font(
    name="Calibri", size=9.5, bold=True, color="FFFFFF")
for cc in range(1, 10):
    ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=NAVY)
rr += 1
for t in CFG["thresholds"]:
    put(ws, rr, [t["id"], t["name"], t["value"], t["unit"], "", "", t["use"], "Business owner", "All status flags"],
        wrap=(2, 7, 9))
    ws.cell(row=rr, column=3).fill = PatternFill("solid", fgColor=INPUT_FILL)
    rr += 1

rr += 1
ws.cell(row=rr, column=1, value="DATA-DERIVED BASELINE FACTS (recomputed at every refresh, not entered by hand)").font = Font(
    name="Calibri", size=9.5, bold=True, color="FFFFFF")
for cc in range(1, 10):
    ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=GREY)
rr += 1
header(ws, rr, ["Fact", "Value", "Unit", "", "", "", "How it is derived", "", ""])
rr += 1
for label, key, unit, how in [
    ("Annualised consumption run rate", "run_rate_12", "PKR", "Average of the last three months x 12"),
    ("Annual budget (product categories)", "annual_budget", "PKR", "Product budget in the window, annualised"),
    ("Stock on hand value", "on_hand_value", "PKR", "Receipts less issues, valued at batch cost"),
    ("Open commitment", "open_commitment", "PKR", "Ordered less received, at PO price"),
    ("Annual beneficiaries", "annual_beneficiaries", "People", "Service-delivery counts, annualised"),
    ("Near-expiry share of stock", "near_expiry_share", "%", "Value expiring within 90 days / stock value"),
    ("Average supplier lead time", "avg_lead_days", "Days", "PO date to first receipt"),
    ("Items below the watch threshold", "stockout_items", "Items", "Cover below 60 days"),
]:
    v = FACTS.get(key)
    fmt = {}
    if unit == "PKR":
        fmt = {2: MONEY}
    elif unit == "%":
        fmt = {2: PCT}
    put(ws, rr, [label, v, unit, "", "", "", how, "", ""], wrap=(7,), fmt=fmt)
    rr += 1

# ============================================================ 02 LIVE CALCULATOR
ws, r = sheet("02_Live_Calculator", "Live calculator – change a yellow cell and the outputs recalculate in Excel",
              "Reproduces the core model equations. Use it for quick conversations; use 03–08 for the full scenario set.",
              widths=[4, 46, 18, 14, 60], tab="F2A65A")
r += 1
ws.cell(row=r, column=2, value="INPUTS").font = Font(name="Calibri", size=10, bold=True, color=NAVY)
r += 1
in_start = r
INPUTS = [
    ("Demand growth", "demand_growth_pct", "demand_growth_pct"),
    ("Price inflation", "price_inflation_pct", "price_inflation_pct"),
    ("One-off price shock", "price_shock_pct", "price_shock_pct"),
    ("Bulk / framework discount", "bulk_discount_pct", "bulk_discount_pct"),
    ("Lead-time multiplier", "lead_time_multiplier", "lead_time_multiplier"),
    ("Supplier reliability", "supplier_reliability_pct", "supplier_reliability_pct"),
    ("Target stock cover (days)", "target_cover_days", "target_cover_days"),
    ("Safety stock (days)", "safety_stock_days", "safety_stock_days"),
    ("Excess stock used instead of buying", "excess_stock_release_pct", "excess_stock_release_pct"),
    ("Wastage rate", "wastage_rate_pct", "wastage_rate_pct"),
    ("Funding reduction", "funding_cut_pct", "funding_cut_pct"),
    ("Budget growth / reallocation", "budget_growth_pct", "budget_growth_pct"),
]
for label, key, pkey in INPUTS:
    ws.cell(row=r, column=2, value=label).font = BODY
    c = ws.cell(row=r, column=3, value=BP.get(key, 0))
    c.fill = PatternFill("solid", fgColor=INPUT_FILL)
    c.font = Font(name="Calibri", size=9.5, bold=True)
    c.border = BOX
    c.number_format = '0.00' if key in ("lead_time_multiplier",) else (
        '0' if key.endswith("_days") else PCT)
    ws.cell(row=r, column=4, value="input").font = SUB
    ws.cell(row=r, column=5, value="").font = SUB
    r += 1

r += 1
ws.cell(row=r, column=2, value="FIXED BASELINE FACTS (from data)").font = Font(
    name="Calibri", size=10, bold=True, color=NAVY)
r += 1
fact_rows = {}
for label, key, fmt in [
    ("Annualised consumption run rate (PKR)", "run_rate_12", MONEY),
    ("Annual budget (PKR)", "annual_budget", MONEY),
    ("Stock on hand (PKR)", "on_hand_value", MONEY),
    ("Open commitment (PKR)", "open_commitment", MONEY),
    ("Annual beneficiaries", "annual_beneficiaries", NUM),
    ("Near-expiry share of stock", "near_expiry_share", PCT),
    ("Average lead time (days)", "avg_lead_days", '0'),
]:
    ws.cell(row=r, column=2, value=label).font = BODY
    c = ws.cell(row=r, column=3, value=FACTS[key])
    c.number_format = fmt
    c.font = BODY
    c.border = BOX
    fact_rows[key] = r
    r += 1

# input cell addresses
IN = {}
for i, (label, key, _) in enumerate(INPUTS):
    IN[key] = f"$C${in_start + i}"
FR = {k: f"$C${v}" for k, v in fact_rows.items()}

r += 1
ws.cell(row=r, column=2, value="OUTPUTS (formulas)").font = Font(name="Calibri", size=10, bold=True, color=NAVY)
r += 1
out_start = r
D = ('1+' + IN["demand_growth_pct"])
P = (f"(1+{IN['price_inflation_pct']})*(1+{IN['price_shock_pct']})*(1-{IN['bulk_discount_pct']})")
LEAD = f"{FR['avg_lead_days']}*{IN['lead_time_multiplier']}"
SAFE = f"{IN['safety_stock_days']}+{LEAD}*(1-{IN['supplier_reliability_pct']})"
CONS = f"{FR['run_rate_12']}*{D}*{P}"
TARGET = f"{CONS}*({IN['target_cover_days']}+{SAFE})/365"
AVAIL = f"{FR['on_hand_value']}+{FR['open_commitment']}"
TOPP = f"MAX(0,{TARGET}-{AVAIL})"
REL = f"MAX(0,{AVAIL}-{TARGET})*{IN['excess_stock_release_pct']}"
PURCH = f"MAX(0,{CONS}+{TOPP}+{IN['wastage_rate_pct']}*{CONS}-{REL})"
BUD = f"{FR['annual_budget']}*(1-{IN['funding_cut_pct']})*(1+{IN['budget_growth_pct']})"
EXPIRY = (f"{FR['near_expiry_share']}*{FR['on_hand_value']}+{IN['wastage_rate_pct']}*{CONS}")

OUTPUTS = [
    ("Demand factor", f"={D}", '0.000', "1 + demand growth"),
    ("Price factor", f"={P}", '0.000', "(1+inflation) x (1+shock) x (1-discount)"),
    ("Projected consumption value (PKR)", f"={CONS}", MONEY, "Run rate x demand factor x price factor"),
    ("Target stock value (PKR)", f"={TARGET}", MONEY, "Consumption x (cover + safety) / 365"),
    ("Available stock (PKR)", f"={AVAIL}", MONEY, "On hand + open commitment"),
    ("Top-up purchase (PKR)", f"={TOPP}", MONEY, "Shortfall against target stock"),
    ("Stock released instead of buying (PKR)", f"={REL}", MONEY, "Excess above target x release rate"),
    ("Purchase requirement (PKR)", f"={PURCH}", MONEY, "Consumption + top-up + wastage - released"),
    ("Consumption-to-purchase ratio", f"=IF({PURCH}=0,0,{CONS}/{PURCH})", RATIO, "Consumption / purchase"),
    ("Budget available (PKR)", f"={BUD}", MONEY, "Budget x (1-cut) x (1+growth)"),
    ("Budget utilisation", f"=IF({BUD}=0,0,{CONS}/{BUD})", PCT, "Consumption / budget"),
    ("Funding gap (PKR)", f"=MAX(0,{CONS}-{BUD})", MONEY, "Consumption less budget available"),
    ("Expiry risk value (PKR)", f"={EXPIRY}", MONEY, "Near-expiry share x stock + wastage rate x consumption"),
    ("Cost per beneficiary (PKR)", f"=IF({FR['annual_beneficiaries']}=0,0,{CONS}/{FR['annual_beneficiaries']})",
     MONEY, "Consumption / beneficiaries"),
]
for label, formula, fmt, explain in OUTPUTS:
    ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=9.5, bold=True)
    c = ws.cell(row=r, column=3, value=formula)
    c.number_format = fmt
    c.font = Font(name="Calibri", size=10, bold=True, color=NAVY)
    c.fill = PatternFill("solid", fgColor=LIGHT)
    c.border = BOX
    ws.cell(row=r, column=5, value=explain).font = SUB
    r += 1

r += 1
ws.cell(row=r, column=2,
        value="Tip: this sheet recalculates in Excel only. The values in 03–08 are computed by the Python model so that the "
              "workbook, the documents and the question bank always agree.").font = SUB

# ============================================================ 03 SCENARIOS
ws, r = sheet("03_Scenarios", "Scenario definitions and levers",
              "Each scenario is a named set of parameter overrides applied to the baseline. Add a scenario by adding an entry to config/parameters.json.",
              widths=[7, 44, 66, 70], tab="4C78A8")
header(ws, r + 1, ["ID", "Scenario", "Rationale", "Levers moved (baseline → scenario)"])
rr = r + 2
for sc in W["scenarios"]:
    lev = "; ".join(f"{k}: {v['baseline']} → {v['scenario']}" for k, v in sc["changed"].items()) or "none (reference case)"
    put(ws, rr, [sc["id"], sc["name"], sc["rationale"], lev], wrap=(2, 3, 4))
    ws.row_dimensions[rr].height = 44
    rr += 1

# ============================================================ 04 SCENARIO SUMMARY
ws, r = sheet("04_Scenario_Summary", "Scenario summary – forward 12 months",
              "Δ columns compare the scenario with the baseline S0. Colour scales are relative across the scenario set.",
              widths=[7, 40, 16, 16, 16, 12, 14, 14, 12, 14, 14], tab="4C78A8")
hr = r + 1
header(ws, hr, ["ID", "Scenario", "Consumption", "Purchase requirement", "Funding gap",
                "Utilisation", "Δ Consumption", "Δ Purchase", "Stock-outs", "Expiry risk", "Working capital"])
rr = hr + 1
first = rr
for sc in W["scenarios"]:
    res = sc["result"]
    put(ws, rr, [sc["id"], sc["name"], res["consumption_value"], res["purchase_value"],
                 res["funding_gap"], res["budget_utilisation"], sc["delta"]["consumption_value"],
                 sc["delta"]["purchase_value"], res["stockout_items"], res["expiry_risk_value"],
                 res["working_capital"]],
        fmt={3: MONEY, 4: MONEY, 5: MONEY, 6: PCT, 7: MONEY, 8: MONEY, 9: NUM, 10: MONEY, 11: MONEY})
    for c in (3, 4, 5, 7, 8, 10, 11):
        ws.cell(row=rr, column=c).number_format = MONEY
    if sc["id"] == "S0":
        for c in range(1, 12):
            ws.cell(row=rr, column=c).fill = PatternFill("solid", fgColor=LIGHT)
            ws.cell(row=rr, column=c).font = Font(name="Calibri", size=9.5, bold=True)
    rr += 1
last = rr - 1
ws.conditional_formatting.add(
    f"E{first}:E{last}",
    ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile", mid_value=50,
                   mid_color="FFEB9C", end_type="max", end_color="F8CBAD"))
ws.conditional_formatting.add(
    f"F{first}:F{last}",
    CellIsRule(operator="greaterThan", formula=["1"], fill=PatternFill("solid", bgColor=RED)))
ws.conditional_formatting.add(
    f"I{first}:I{last}",
    CellIsRule(operator="greaterThan", formula=["4"], fill=PatternFill("solid", bgColor=AMBER)))

ch = BarChart()
ch.type = "col"
ch.title = "Funding gap by scenario"
ch.y_axis.title = "PKR"
ch.height = 8
ch.width = 20
data = Reference(ws, min_col=5, min_row=hr, max_row=last)
cats = Reference(ws, min_col=1, min_row=first, max_row=last)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.legend = None
ch.y_axis.numFmt = '#,##0'
ch.series[0].graphicalProperties.solidFill = NAVY
ws.add_chart(ch, f"A{last + 3}")

ch2 = BarChart()
ch2.type = "col"
ch2.title = "Purchase requirement and expiry risk by scenario"
ch2.height = 8
ch2.width = 20
d1 = Reference(ws, min_col=4, min_row=hr, max_row=last)
d2 = Reference(ws, min_col=10, min_row=hr, max_row=last)
ch2.add_data(d1, titles_from_data=True)
ch2.add_data(d2, titles_from_data=True)
ch2.set_categories(cats)
ch2.y_axis.numFmt = '#,##0'
ch2.y_axis.title = "PKR"
ws.add_chart(ch2, f"A{last + 20}")

# ============================================================ 05 DETAIL
ws, r = sheet("05_Scenario_Detail", "Scenario detail – every output for every scenario",
              "Long format so it can be filtered and pivoted. Δ = scenario less baseline.",
              widths=[7, 34, 30, 16, 16, 16, 12], tab="72B7B2")
header(ws, r + 1, ["Scenario", "Scenario name", "Output", "Baseline", "Scenario value", "Delta", "Delta %"])
rr = r + 2
LABELS = [
    ("consumption_value", "Consumption value (PKR)", MONEY, True),
    ("purchase_value", "Purchase requirement (PKR)", MONEY, True),
    ("top_up", "Top-up purchase (PKR)", MONEY, True),
    ("excess_stock_released", "Stock released instead of buying (PKR)", MONEY, True),
    ("budget_available", "Budget available (PKR)", MONEY, True),
    ("funding_gap", "Funding gap (PKR)", MONEY, True),
    ("budget_utilisation", "Budget utilisation", PCT, False),
    ("ratio", "Consumption-to-purchase ratio", RATIO, False),
    ("expiry_risk_value", "Expiry risk value (PKR)", MONEY, True),
    ("overstock_value", "Overstock value (PKR)", MONEY, True),
    ("stockout_items", "Stock-out risk items", NUM, False),
    ("working_capital", "Working capital in stock (PKR)", MONEY, True),
    ("cost_per_beneficiary", "Cost per beneficiary (PKR)", MONEY, True),
    ("beneficiaries", "Beneficiaries (annual)", NUM, False),
    ("lead_time_days", "Lead time used (days)", '0.0', False),
    ("safety_days", "Safety stock incl. variability (days)", '0.0', False),
    ("demand_factor", "Demand factor", '0.000', False),
    ("price_factor", "Price factor", '0.000', False),
]
for sc in W["scenarios"]:
    for key, label, fmt, is_money in LABELS:
        base = B.get(key)
        val = sc["result"].get(key)
        delta = (val - base) if isinstance(val, (int, float)) and isinstance(base, (int, float)) else None
        pct = (delta / base) if delta is not None and base else None
        put(ws, rr, [sc["id"], sc["name"], label, base, val, delta, pct],
            fmt={4: fmt, 5: fmt, 6: fmt, 7: PCT})
        rr += 1
ws.auto_filter.ref = f"A{r+1}:G{rr-1}"
ws.freeze_panes = ws.cell(row=r + 2, column=1)

# ============================================================ 06 SENSITIVITY
ws, r = sheet("06_Sensitivity", "Sensitivity analysis – one lever at a time (tornado)",
              "Each lever is moved between its low and high value while everything else stays at baseline. Ranked by swing on each output.",
              widths=[30, 14, 14, 16, 16, 16, 14, 14], tab="E45756")
rr = r + 1
OUT_TITLES = {
    "purchase_value": "Purchase requirement (PKR)",
    "funding_gap": "Funding gap (PKR)",
    "expiry_risk_value": "Expiry risk value (PKR)",
    "stockout_items": "Stock-out risk items",
}
for ridx, (out_key, order_keys) in enumerate(W["sensitivity"]["order"].items()):
    ws.cell(row=rr, column=1, value=f"Output: {OUT_TITLES.get(out_key, out_key)}").font = Font(
        name="Calibri", size=10.5, bold=True, color=NAVY)
    rr += 1
    header(ws, rr, ["Lever", "Low value", "High value", "Result at low", "Result at high",
                    "Swing", "Base", "Rank"])
    rr += 1
    by_key = {row["key"]: row for row in W["sensitivity"]["rows"]}
    start = rr
    for rank, key in enumerate(order_keys, start=1):
        row = by_key[key]
        o = row["outputs"][out_key]
        fmt = MONEY if "value" in out_key else NUM
        put(ws, rr, [key, row["low"], row["high"], o["low"], o["high"], o["swing"], o["base"], rank],
            fmt={2: '0.000', 3: '0.000', 4: fmt, 5: fmt, 6: fmt, 7: fmt})
        rr += 1
    ws.conditional_formatting.add(
        f"F{start}:F{rr-1}",
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F8CBAD"))
    rr += 1

# ============================================================ 07 BREAKEVEN
ws, r = sheet("07_BreakEven", "Break-even and trade-off analysis",
              "How far a driver can move before the budget breaches, and what each stock-cover policy costs.",
              widths=[30, 22, 18, 18, 20, 60], tab="F2A65A")
header(ws, r + 1, ["Lever", "Condition", "Break-even value", "Result at low", "Result at high", "Note"])
rr = r + 2
for b in W["breakeven"]:
    be = b["breakeven"]
    put(ws, rr, [b["key"], b["target"], be if be is not None else "n/a",
                 b["value_at_low"], b["value_at_high"], b["note"]],
        fmt={3: '0.0000', 4: MONEY, 5: MONEY}, wrap=(6,))
    rr += 1

rr += 2
ws.cell(row=rr, column=1, value="Stock-cover policy trade-off").font = Font(
    name="Calibri", size=10.5, bold=True, color=NAVY)
rr += 1
header(ws, rr, ["Target cover (days)", "Purchase requirement", "Expiry risk", "Working capital", "Stock-out items"])
rr += 1
c0 = rr
curve = next((b["curve"] for b in W["breakeven"] if b.get("curve")), [])
for pt in curve:
    put(ws, rr, [pt["x"], pt["purchase"], pt["expiry"], pt["capital"], pt["stockouts"]],
        fmt={1: '0', 2: MONEY, 3: MONEY, 4: MONEY, 5: NUM})
    rr += 1

ch = LineChart()
ch.title = "Purchase cost vs expiry risk by cover policy"
ch.x_axis.title = "Target cover (days)"
ch.y_axis.title = "PKR"
ch.height = 8
ch.width = 20
d1 = Reference(ws, min_col=2, min_row=c0 - 1, max_row=c0 + len(curve) - 1)
d2 = Reference(ws, min_col=3, min_row=c0 - 1, max_row=c0 + len(curve) - 1)
cats = Reference(ws, min_col=1, min_row=c0, max_row=c0 + len(curve) - 1)
ch.add_data(d1, titles_from_data=True)
ch.add_data(d2, titles_from_data=True)
ch.set_categories(cats)
ch.y_axis.numFmt = '#,##0'
ws.add_chart(ch, f"H{c0}")

# ============================================================ 08 TWO-WAY
ws, r = sheet("08_TwoWay", "Two-way table – demand growth x price inflation",
              "Each cell is a full model run. Row = demand growth, column = price inflation. Use it to locate the current position and see which direction hurts more.",
              widths=[22, 16, 16, 16, 16, 16, 16], tab="E45756")
row_key = W["twoway"]["row_key"]
col_key = W["twoway"]["col_key"]
for oi, (out_key, title, fmt) in enumerate([
    ("funding_gap", "Funding gap (PKR)", MONEY),
    ("budget_utilisation", "Budget utilisation", PCT),
]):
    ws.cell(row=r, column=1, value=title).font = Font(name="Calibri", size=10.5, bold=True, color=NAVY)
    r += 1
    header(ws, r, [f"{row_key} \\ {col_key}"] + [f"{c:.0%}" for c in W["twoway"]["col_values"]])
    r += 1
    start = r
    for ri, rv in enumerate(W["twoway"]["row_values"]):
        vals = [f"{rv:.0%}"] + W["twoway"]["tables"][out_key][ri]
        put(ws, r, vals, fmt={i: fmt for i in range(2, 2 + len(W["twoway"]["col_values"]))})
        r += 1
    ws.conditional_formatting.add(
        f"B{start}:{get_column_letter(1 + len(W['twoway']['col_values']))}{r-1}",
        ColorScaleRule(start_type="min", start_color="C6EFCE", mid_type="percentile", mid_value=50,
                       mid_color="FFEB9C", end_type="max", end_color="F8CBAD"))
    r += 2

# ============================================================ 09 LEVER CATALOGUE
ws, r = sheet("09_Lever_Catalogue", "Lever catalogue – what management can actually change",
              "Every lever in the model, with the working range, the owner and the effect it has.",
              widths=[26, 36, 10, 10, 12, 18, 56], tab="1F3B57")
header(ws, r + 1, ["Lever (parameter key)", "Lever", "Current", "Range", "Group", "Owner", "Effect when moved"])
rr = r + 2
EFFECTS = {
    "demand_growth_pct": "Raises consumption, purchasing and the funding gap almost one-for-one; also raises stock-out risk.",
    "service_volume_growth_pct": "Lowers cost per beneficiary when it rises; does not by itself change consumption.",
    "activity_factor": "Scales every projection; use it for project start and end effects.",
    "price_inflation_pct": "The single largest driver of the funding gap after demand; flows straight into purchase value.",
    "price_shock_pct": "One-off price event; use to model a currency move or loss of contract price.",
    "bulk_discount_pct": "Reduces purchase value and improves the ratio; the counterpart of framework contracting.",
    "lead_time_multiplier": "Raises safety stock and target stock, so purchasing rises and stock-out risk grows sharply.",
    "supplier_reliability_pct": "Raises safety stock when it falls; the main driver of stock-out risk.",
    "target_cover_days": "Trades working capital and purchase cost against stock-out protection.",
    "safety_stock_days": "Adds buffer: more stock, fewer stock-outs, more expiry risk.",
    "max_cover_days": "Defines what counts as overstock; driving it down flags more stock as excess.",
    "expiry_alert_days": "Widens the alert window; earlier disposition, less write-off.",
    "wastage_rate_pct": "Directly adds replacement purchase and expiry write-off.",
    "redistribution_share_pct": "Converts potential write-off into usable stock; lowers expiry risk.",
    "excess_stock_release_pct": "Lets stock already held offset new purchases; lowers purchase value.",
    "funding_cut_pct": "Reduces budget available and widens the funding gap without changing consumption.",
    "budget_growth_pct": "Closes the funding gap; the lever for reallocation or additional grant.",
    "order_cycle_days": "Longer cycles mean larger orders and more average stock; shorter cycles mean more transactions.",
}
for spec in CFG["model"]:
    v = BP.get(spec["key"], spec["value"])
    rng = f"{spec['min']} – {spec['max']}"
    put(ws, rr, [spec["key"], spec["name"], v, rng, spec["group"], spec["owner"],
                 EFFECTS.get(spec["key"], "")], wrap=(2, 7))
    ws.cell(row=rr, column=3).fill = PatternFill("solid", fgColor=INPUT_FILL)
    rr += 1

# ============================================================ 10 MODEL LOGIC
ws, r = sheet("10_Model_Logic", "Model logic – the equations, in one place",
              "Deliberately simple and transparent so that any number can be reproduced by hand.",
              widths=[34, 70, 60], tab="595959")
header(ws, r + 1, ["Step", "Equation", "Where it matters"])
rr = r + 2
LOGIС = [
    ("Demand factor", "1 + demand_growth_pct", "Scales consumption volume"),
    ("Price factor", "(1 + price_inflation_pct) x (1 + price_shock_pct) x (1 - bulk_discount_pct)", "Scales value"),
    ("Projected consumption value", "run_rate_12 x demand factor x price factor x activity_factor",
     "Top line of every scenario"),
    ("Lead time used", "avg_lead_days x lead_time_multiplier", "Drives safety stock and stock-out risk"),
    ("Safety stock (days)", "safety_stock_days + lead time x (1 - supplier_reliability_pct)",
     "Buffer against variability"),
    ("Target stock value", "consumption_value x (target_cover_days + safety days) / 365", "Sets the stock position to fund"),
    ("Available stock", "on_hand_value + open_commitment", "Stock already paid for"),
    ("Top-up purchase", "max(0, target stock - available stock)", "Additional buying needed"),
    ("Stock released instead of buying", "max(0, available stock - target stock) x excess_stock_release_pct",
     "Working capital released"),
    ("Purchase requirement", "consumption + top-up + wastage replacement - released", "Cash and procurement load"),
    ("Budget available", "annual_budget x (1 - funding_cut_pct) x (1 + budget_growth_pct)", "The envelope"),
    ("Funding gap", "max(0, consumption - budget available)", "The decision trigger"),
    ("Expiry risk value", "near_expiry_share x stock + wastage_rate x consumption + 0.35 x excess stock x (1 - redistribution)",
     "Write-off exposure"),
    ("Stock-out risk items", "count of items whose projected cover < lead time + safety days", "Service risk"),
    ("Cost per beneficiary", "consumption / (annual_beneficiaries x (1 + service_volume_growth))", "Value for money"),
    ("Cover trade-off curve", "purchase, expiry and working capital recomputed for cover 30–210 days",
     "Policy choice"),
]
for a, b, c in LOGIС:
    put(ws, rr, [a, b, c], wrap=(2, 3))
    rr += 1
rr += 1
for line in [
    "The model is linear in its main drivers by design; it does not attempt to forecast item-level seasonality. Use the "
    "seasonal index in the analysis workbook for peak planning.",
    "Baseline facts are recomputed from the data at each refresh; only the assumptions are entered by a human.",
    "The forward horizon is 12 months at the current operating level. Project end dates are handled separately in the "
    "KPI workbook (FY2026-27 projection).",
    "All figures are PKR, excluding sales tax, on the same measurement basis as the dashboard.",
]:
    c = ws.cell(row=rr, column=1, value="•  " + line)
    c.font = BODY
    rr += 1

# ============================================================ 11 REFRESH
ws, r = sheet("11_Refresh", "How to refresh or change the model",
              "The model is regenerated from configuration and data; nothing is edited inside this workbook except the live calculator.",
              widths=[8, 60, 30, 50], tab="72B7B2")
header(ws, r + 1, ["Step", "Action", "Where", "Effect"])
rr = r + 2
STEPS = [
    ("1", "Change a threshold (status band) or an assumption", "config/parameters.json → thresholds / model", "Every status flag, commentary line and scenario recalculates"),
    ("2", "Add or edit a scenario", "config/parameters.json → scenarios", "It appears in 03 and 04 on the next build"),
    ("3", "Change the sensitivity ranges or outputs", "config/parameters.json → sensitivity", "06_Sensitivity is rebuilt"),
    ("4", "Change the two-way axes", "config/parameters.json → twoway", "08_TwoWay is rebuilt"),
    ("5", "Refresh the data", "scripts/generate_dataset.py (or the Oracle extract)", "findings.json is recomputed"),
    ("6", "Rebuild everything", "python3 scripts/refresh_all.py", "Workbooks, guides and BRD figures regenerate"),
    ("7", "Quick manual check only", "02_Live_Calculator (in Excel)", "Outputs recalculate without rebuilding"),
]
for step in STEPS:
    put(ws, rr, list(step), wrap=(2, 4))
    rr += 1

rr += 1
c = ws.cell(row=rr, column=1,
            value="Governance: thresholds and assumptions are business decisions. Every change is recorded with the old value, "
                  "the new value, the reason, the approver and the effective date, and applies from the next refresh without "
                  "rewriting history.")
c.font = SUB

for wsx in wb.worksheets:
    wsx.sheet_view.zoomScale = 90
wb.active = wb.sheetnames.index("00_ReadMe")
wb.save(OUT)
print("saved", OUT)
print("scenarios", len(W["scenarios"]), "| levers", len(CFG["model"]), "| thresholds", len(CFG["thresholds"]))

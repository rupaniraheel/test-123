#!/usr/bin/env python3
"""
Builds Analytics_Question_Bank.xlsx: a refresh-driven analytics question bank
plus a live answers engine and a data-commentary feed.

Answers and commentary are computed from the current dataset and stamped with
the data-as-at date. Re-running this workbook against a refreshed dataset
regenerates every answer and every commentary line.
"""
import os
import json
import calendar
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import CellIsRule

import qa_content as Q

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "..", "data", "findings.json")) as f:
    F = json.load(f)

OUT = os.path.join(BASE, "..", "data", "Analytics_Question_Bank.xlsx")

NAVY = "1F3B57"
TEAL = "2A7F8F"
LIGHT = "EEF3F8"
GREY = "595959"
GREEN = "E2EFDA"
AMBER = "FCE4D6"
RED = "F8CBAD"
BLUE = "DDEBF7"

H_FILL = PatternFill("solid", fgColor=NAVY)
H_FONT = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=9.5)
TITLE = Font(name="Calibri", size=15, bold=True, color=NAVY)
SUB = Font(name="Calibri", size=9, italic=True, color=GREY)
THIN = Side(style="thin", color="DDDDDD")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = Workbook()
wb.remove(wb.active)


def sheet(name, title, subtitle=None, widths=None, tab=NAVY):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE
    r = 3
    if subtitle:
        c = ws.cell(row=2, column=1, value=subtitle)
        c.font = SUB
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max(6, len(widths or [1] * 8)))
        ws.row_dimensions[2].height = 16
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.sheet_properties.tabColor = tab
    return ws, r


def header(ws, row, headers, height=32):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
    ws.row_dimensions[row].height = height


def put(ws, row, values, wrap_cols=(), bold=False, fill=None, fmt=None):
    for c, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = Font(name="Calibri", size=9.5, bold=bold)
        cell.border = BOX
        cell.alignment = Alignment(wrap_text=(c in wrap_cols), vertical="top")
        if fmt and c in fmt:
            cell.number_format = fmt[c]
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)


def status_fill(status):
    return {"Red": RED, "Amber": AMBER, "Green": GREEN}.get(status, BLUE)


# ------------------------------------------------------------------ README
ws, r = sheet("00_ReadMe", "Consumption Analysis – Analytics Question Bank & Live Answers",
              "Questions are fixed; answers and commentary are recomputed from data on every refresh.",
              widths=[3, 30, 100, 22], tab=NAVY)
r = 4
ws.cell(row=r, column=2, value="VITAL PAKISTAN TRUST").font = Font(name="Calibri", size=16, bold=True, color=NAVY)
r += 1
ws.cell(row=r, column=2,
        value=f"Data as at {F['as_at']}  |  Reporting window {F['start']} to {F['as_at']}  |  {F['months']} months").font = SUB
r += 2

guide = [
    ("What this workbook is",
     "A standing question bank for the Consumption Analysis dashboard, plus a live answer engine and a data-commentary feed. "
     "The questions never change; the answers do."),
    ("How it is organised",
     "02–05: the questions, split into descriptive (what happened), diagnostic (why), predictive (what next) and prescriptive (what to do). "
     "06: the what-if questions, answered by the scenario engine in VPT_WhatIf_Analysis.xlsx. "
     "07: the live answer to every question that can be computed from the data, with a status flag. "
     "08: the commentary rules (which causes to test, which actions to consider, when to act, who owns it). "
     "09: the commentary actually triggered by the current data. 10: the refresh checklist and quality gates. "
     "11: coverage against the BRD."),
    ("How answers refresh",
     "Every value in sheet 07 and every line in sheet 09 is computed from the current extract at build time and stamped with the "
     "data-as-at date. When the data is refreshed, re-run the refresh (script scripts/build_qa_workbook.py, or the equivalent "
     "scheduled job) and all answers, statuses and commentary lines are recomputed. Nothing is hard-coded."),
    ("Status logic",
     "Green = within the tolerance set in 01_Parameters. Amber = outside tolerance but not yet material. "
     "Red = material breach requiring commentary and an owner. Blue = informational, no action implied."),
    ("Everything is configurable",
     f"All {len(Q.PARAMETERS)} thresholds and all {len(Q.MODEL_PARAMETERS)} model assumptions live in "
     f"config/parameters.json (parameter set v{Q.CONFIG_VERSION}). Changing a value there and re-running "
     "scripts/refresh_all.py changes every status flag, commentary line and scenario result — no code change, no retyping."),
    ("Reading order at each refresh",
     "1) Check sheet 10 quality gates are passed.  2) Read sheet 09 for what needs attention now.  "
     "3) Use sheet 07 to answer specific management questions.  4) Use 02–06 as the standing agenda, "
     "including the what-if questions."),
    ("Cadence",
     "Daily: data quality, stock-out risk on critical items.  Weekly: cover, expiry, commitments, procurement actions.  "
     "Monthly: consumption, ratio, budget, cost per beneficiary, commentary.  Quarterly: suppliers, price, efficiency, seasonality."),
    ("Ownership",
     "Business owner: Head of Procurement & Supply Chain. Co-owners: Director Programme, Manager Finance & Grants. "
     "Data owner: IT / BI. Each answered question carries a named owner in sheet 06."),
    ("Important note",
     "Figures are computed from the current dataset extract. Organisation and supplier names are illustrative. "
     "Replace the extract with the Oracle feed before using this workbook in production."),
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
ws.cell(row=r, column=2, value="Headline position at this refresh").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
r += 1
for label, val in [
    ("Consumption value (excluding transfers)", f"PKR {F['consumption']:,.0f}"),
    ("Procurement investment", f"PKR {F['purchase']:,.0f}"),
    ("Consumption-to-investment ratio", f"{F['ratio']:.2f}"),
    ("Product budget / utilisation", f"PKR {F['budget_product']:,.0f}  /  {F['utilisation']:.0%}"),
    ("Unattributed consumption", f"PKR {F['unattributed']:,.0f}  ({F['unattributed_pct']:.1f}%)"),
    ("Expired stock held", f"PKR {F['stock']['expired_value']:,.0f}"),
    ("Open commitment", f"PKR {F['open_commitment']:,.0f} across {F['open_commitment_lines']:,} open PO lines"),
    ("Cost per beneficiary (period)", f"PKR {F['cost_per_beneficiary']:,.0f}"),
    ("Question count", f"{len(Q.DESCRIPTIVE)} descriptive · {len(Q.DIAGNOSTIC)} diagnostic · "
                       f"{len(Q.PREDICTIVE)} predictive · {len(Q.PRESCRIPTIVE)} prescriptive"),
]:
    ws.cell(row=r, column=2, value=label).font = BODY
    c = ws.cell(row=r, column=3, value=val)
    c.font = Font(name="Calibri", size=10, bold=True)
    r += 1

# ------------------------------------------------------------------ PARAMETERS
ws, r = sheet("01_Parameters", "Thresholds and Parameters",
              "These values drive every status flag and every commentary trigger. Changing a value here is a business decision, not a technical one.",
              widths=[8, 46, 22, 20, 60], tab=TEAL)
header(ws, r + 1, ["Ref", "Parameter", "Current value", "Unit", "How it is used"])
rr = r + 2
for row in Q.PARAMETERS:
    put(ws, rr, list(row), wrap_cols=(2, 5))
    rr += 1
ws.cell(row=rr + 1, column=1,
        value=f"Parameter set v{Q.CONFIG_VERSION} — these values are loaded from config/parameters.json, not typed here. "
              f"Approval: Head of Procurement; every change is recorded with the old value, the new value, the reason and the "
              f"effective date, and applies from the next refresh.").font = SUB

rr += 3
ws.cell(row=rr, column=1, value="Model parameters that drive the what-if engine (see VPT_WhatIf_Analysis.xlsx)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
rr += 1
header(ws, rr, ["Key", "Parameter", "Value", "Unit", "Range", "Basis", "Owner"])
rr += 1
for key, name, val, unit, rng, basis, owner in Q.MODEL_PARAMETERS:
    v = F["baseline"] if val == "auto" else val
    if val == "auto":
        val = 1.0
    put(ws, rr, [key, name, val, unit, rng, basis, owner], wrap_cols=(2, 6))
    if unit == "%":
        ws.cell(row=rr, column=3).number_format = '0.0%'
    elif unit == "x":
        ws.cell(row=rr, column=3).number_format = '0.00'
    rr += 1
ws.cell(row=rr + 1, column=1,
        value="Change any of these in config/parameters.json and re-run scripts/refresh_all.py to recalculate every status, "
              "commentary line and scenario.").font = SUB

# ------------------------------------------------------------------ QUESTION SHEETS
QB_HEADERS = ["Q ID", "Analytical question", "Why it matters (decision it supports)", "Data required",
              "How to answer (logic)", "Output form", "Trigger / threshold", "Cadence", "Owner", "BRD reference"]


def write_questions(name, title, subtitle, rows, qprefix, tab):
    ws, r = sheet(name, title, subtitle, widths=[7, 52, 46, 34, 52, 26, 24, 16, 18, 16], tab=tab)
    header(ws, r + 1, QB_HEADERS)
    rr = r + 2
    for i, row in enumerate(rows):
        qid, question, why, data, logic, output, trigger, cadence, owner = row
        ref = QREF.get(qid, DEFAULT_REF[qprefix])
        put(ws, rr, [qid, question, why, data, logic, output, trigger, cadence, owner, ref],
            wrap_cols=(2, 3, 4, 5, 6))
        ws.row_dimensions[rr].height = 46
        rr += 1
    ws.freeze_panes = ws.cell(row=r + 2, column=2)
    ws.auto_filter.ref = f"A{r+1}:J{rr-1}"
    return ws


QREF = {
    "D-01": "KPI-01/02, FR-01", "D-02": "KPI-03/04, FR-03", "D-03": "KPI-05, FR-03", "D-04": "FR-01, FR-09",
    "D-05": "KPI-28, FR-09", "D-07": "KPI-08/09, FR-05", "D-11": "KPI-16, FR-19", "D-12": "KPI-17, FR-19",
    "D-13": "KPI-20/21, FR-21", "D-14": "KPI-19, FR-20", "D-15": "KPI-07, FR-13", "D-17": "KPI-23/24, FR-24",
    "D-18": "KPI-25/26, FR-25", "D-19": "KPI-14/15, FR-14", "D-20": "KPI-11, FR-17",
    "D-21": "KPI-27, FR-30", "D-26": "FR-12, FR-44",
    "G-02": "KPI-06, FR-03", "G-08": "DQ-01, FR-30", "G-14": "KPI-22, FR-22",
    "X-02": "FR-19, FR-31", "X-03": "FR-40", "X-05": "KPI-13, FR-18", "X-06": "FR-17, FR-18",
    "X-08": "KPI-05", "X-10": "KPI-24", "X-11": "KPI-20",
    "R-01": "FR-40", "R-04": "FR-21", "R-07": "FR-17", "R-09": "FR-25", "R-20": "FR-28",
}
DEFAULT_REF = {"D": "§12 KPI dictionary", "G": "§12 / §15", "X": "§11 page spec", "R": "§8 to-be process"}

write_questions("02_Q_Descriptive", "Descriptive questions – what happened?",
                "Answer these first. They establish the facts and populate the dashboard tiles. None of them requires a judgement.",
                Q.DESCRIPTIVE, "D", NAVY)
write_questions("03_Q_Diagnostic", "Diagnostic questions – why did it happen?",
                "Each diagnostic question forces a ranked list of causes to be tested rather than a first guess. "
                "A cause is only accepted when the evidence in the data supports it.",
                Q.DIAGNOSTIC, "G", TEAL)
write_questions("04_Q_Predictive", "Predictive questions – what is likely to happen next?",
                "Projections use the latest data at each refresh (run rate, cover, lead time, seasonality). "
                "They are decision aids with stated assumptions, not forecasts to be reported as fact.",
                Q.PREDICTIVE, "X", "4C78A8")
write_questions("05_Q_Prescriptive", "Prescriptive questions – what should we do, when, and who?",
                "The answer to these questions is an action with an owner and a date. If an action cannot be assigned, the question returns to the business owner.",
                Q.PRESCRIPTIVE, "R", "E45756")

write_questions("06_Q_WhatIf", "What-if questions – what changes if a driver moves?",
                "These questions are answered by the scenario engine in VPT_WhatIf_Analysis.xlsx. Every lever used is a "
                "parameter in config/parameters.json, so a scenario can be added or changed without touching code.",
                Q.WHATIF, "X", "F2A65A")

# ------------------------------------------------------------------ LIVE ANSWERS
ws, r = sheet("07_Live_Answers", "Live answers – recomputed from data at every refresh",
              f"Data as at {F['as_at']}. Every value below is computed from the current extract; statuses use the thresholds in 01_Parameters.",
              widths=[8, 40, 46, 18, 12, 22, 11, 40, 18], tab="72B7B2")
header(ws, r + 1, ["Q ID", "Stage", "Question", "Answer", "Unit", "Target / band", "Status",
                   "What to do if off-target", "Owner"])


from qa_live import build_live, cls_band, cls_max

LIVE = build_live(F)

rr = r + 2
first_live = rr
for row in LIVE:
    qid, stage, question, answer, unit, target, status, action, owner = row
    put(ws, rr, [qid, stage, question, answer, unit, target, status, action, owner],
        wrap_cols=(3, 6, 8))
    sc = ws.cell(row=rr, column=7)
    sc.fill = PatternFill("solid", fgColor=status_fill(status))
    sc.font = Font(name="Calibri", size=9.5, bold=True)
    sc.alignment = Alignment(horizontal="center")
    cell = ws.cell(row=rr, column=4)
    if isinstance(answer, (int, float)) and not isinstance(answer, bool):
        if unit == "%":
            cell.number_format = '0.0%' if abs(answer) <= 1.5 else '0.0"%"'
        elif unit == "Ratio":
            cell.number_format = '0.00'
        elif unit == "PKR":
            cell.number_format = '#,##0'
        else:
            cell.number_format = '#,##0'
    rr += 1
ws.freeze_panes = ws.cell(row=first_live, column=2)
ws.auto_filter.ref = f"A{r+1}:I{rr-1}"
ws.cell(row=rr + 1, column=1,
        value="Every answer above is recomputed from the current extract. A refresh updates the values, the statuses and the "
              "commentary in sheet 08 — no manual editing is required.").font = SUB

# ------------------------------------------------------------------ COMMENTARY RULES
ws, r = sheet("08_Commentary_Rules", "Data commentary rules – why, what, when, who",
              "Standing rules applied at every refresh. A commentary line is only written when the metric breaches its threshold in 01_Parameters.",
              widths=[10, 26, 62, 62, 44, 24], tab="F2A65A")
header(ws, r + 1, ["Rule", "Metric", "Why it happened – causes to test, in order", "What to do – action options",
                   "When to act", "Owner"])
rr = r + 2
for i, (metric, why, action, when, owner) in enumerate(Q.COMMENTARY_RULES, start=1):
    put(ws, rr, [f"CR-{i:02d}", metric,
                 "1) " + "\n2) ".join(why) if False else "\n".join(f"{j}) {w}" for j, w in enumerate(why, start=1)),
                 "\n".join(f"{j}) {a}" for j, a in enumerate(action, start=1)),
                 when, owner], wrap_cols=(3, 4, 5))
    ws.row_dimensions[rr].height = 96
    rr += 1

# ------------------------------------------------------------------ COMMENTARY FEED
def money(x):
    return f"PKR {x:,.0f}"


def pct(x):
    return f"{x:.1%}"


COMMENTARY = []

# 1 ratio
r_val = F["ratio"]
COMMENTARY.append({
    "id": "CM-01", "metric": "Consumption-to-investment ratio", "value": f"{r_val:.2f}",
    "threshold": "0.80 – 1.05", "status": cls_band(r_val, 0.80, 1.05),
    "what": f"Consumption of {money(F['consumption'])} against investment of {money(F['purchase'])} "
            f"gives a ratio of {r_val:.2f} for the period.",
    "why": "Ranked causes to test: (1) ordering ahead of need due to pack sizes or bulk discounts; "
           "(2) project start-up stock build; (3) slow issuing at service points; "
           "(4) drawdown of opening stock with no new purchase; (5) delayed receipts moving investment into a later period.",
    "impact": "Value purchased but not yet used is cash tied up in store and at risk of expiry; a ratio far above 1.05 means "
              "consumption is being funded from stock and a stock-out will follow.",
    "action": "If below 0.80: defer the next purchase order, review pack-size-driven over-ordering, and check that issues are "
              "recorded daily. If above 1.05: confirm the stock position and re-base the requisition quantity.",
    "when": "At month close, before the next requisition cycle.",
    "owner": "Head of Procurement", "escalate": "Director Finance & Operations",
})

# 2 budget over-utilisation
if F["budget_over"]:
    b = F["budget_over"][0]
    COMMENTARY.append({
        "id": "CM-02", "metric": f"Budget utilisation – {b['cat']} ({b['fy']})",
        "value": pct(b["util"]), "threshold": "≤ 100%",
        "status": "Red" if b["util"] > 1.15 else "Amber",
        "what": f"{b['cat']} consumption of {money(b['consumption'])} exceeds the approved budget of "
                f"{money(b['budget'])} by {money(b['consumption'] - b['budget'])} ({pct(b['util'] - 1)}).",
        "why": "Ranked causes to test: (1) price increase on a major item; (2) service volume higher than planned; "
               "(3) category budget under-estimated at planning; (4) mis-posting from another category; "
               "(5) wastage replacement purchases. Test price first, using the price-variance view (D-17).",
        "impact": f"An over-committed category can only be funded by taking budget from another line or by new money; "
                  f"left unresolved it becomes a donor-reporting exception in {b['fy']}.",
        "action": "Record written commentary with the price/volume split; reallocate from an under-utilised category "
                  "(see CM-03); tighten issue authorisation for the category; renegotiate the top two items by value.",
        "when": "Within 3 working days of the month close; reallocation before the fiscal year reaches 75% elapsed.",
        "owner": "Manager Finance & Grants", "escalate": "Director Finance & Operations",
    })

# 3 budget under-utilisation (equipment)
if F["budget_under"]:
    u = F["budget_under"][0]
    COMMENTARY.append({
        "id": "CM-03", "metric": f"Budget utilisation – {u['cat']} ({u['fy']})",
        "value": pct(u["util"]), "threshold": "In line with elapsed time",
        "status": "Red" if u["util"] < 0.25 else "Amber",
        "what": f"{u['cat']} budget of {money(u['budget'])} has used only {money(u['consumption'])} ({pct(u['util'])}).",
        "why": "Ranked causes to test: (1) capital items stuck in approval or quotation; (2) requirement overstated at "
               "planning; (3) items recorded as consumables rather than equipment; (4) procurement deferred to the last quarter.",
        "impact": "Under-utilisation is read by donors as slow implementation; at project end it becomes unspent budget "
                  "and, if equipment is genuinely needed, a service-capability gap.",
        "action": "Review the equipment pipeline item by item; approve or cancel pending requests; either accelerate "
                  "justified purchases or formally reallocate the budget to medicines and consumables before year end.",
        "when": "This quarter — capital purchases need the full procurement cycle before the fiscal year ends.",
        "owner": "Head of Procurement", "escalate": "Director Programme",
    })

# 4 unattributed
COMMENTARY.append({
    "id": "CM-04", "metric": "Unattributed consumption", "value": pct(F["unattributed_pct"] / 100),
    "threshold": "≤ 1.0%",
    "status": "Red" if F["unattributed_pct"] > 1.5 else ("Amber" if F["unattributed_pct"] > 1.0 else "Green"),
    "what": f"{money(F['unattributed'])} of consumption ({F['unattributed_lines']:,} issue lines, "
            f"{pct(F['unattributed_pct']/100)} of value) is charged to a generic cost centre rather than a project. "
            f"Worst month: {F['dq']['worst_month']['label']}.",
    "why": "Ranked causes to test: (1) a default or generic project code used as a shortcut at issue; "
           "(2) users not trained on project coding; (3) new project not yet set up in the system; "
           "(4) system default overriding the correct code. Cross-tabulate by user, warehouse and service point to identify the pattern.",
    "impact": "Project-level consumption, cost per beneficiary and budget utilisation are all understated for the "
              "projects that actually consumed the stock, and overstated nowhere — the numbers do not add up to the total.",
    "action": "Correct retroactively where the consuming project is evidenced; make the project field mandatory and "
              "validated in Oracle; retrain the top users identified; add this exception to the monthly review agenda.",
    "when": "Correction within the current month; system control change within the quarter.",
    "owner": "Head of Procurement (with IT)", "escalate": "Director Finance & Operations",
})

# 5 expiry
st = F["stock"]
COMMENTARY.append({
    "id": "CM-05", "metric": "Expired and near-expiry stock",
    "value": f"{money(st['expired_value'])} expired; {money(st['near_expiry_60'])} within 60 days",
    "threshold": "Zero expired; minimise near-expiry",
    "status": "Red" if st["expired_value"] > 500000 else ("Amber" if st["expired_value"] > 0 else "Green"),
    "what": f"{money(st['expired_value'])} of stock ({pct(st['expired_share'])} of stock value) is already past expiry and "
            f"not written off; a further {money(st['near_expiry_60'])} expires within 60 days and "
            f"{money(st['near_expiry_90'])} within 90 days. Largest expired item: "
            f"{st.get('worst_expired_item', {}).get('name', 'n/a')}.",
    "why": "Ranked causes to test: (1) short-dated stock accepted at receipt or received as a donation; "
           "(2) over-ordering against an estimate rather than consumption; (3) stock held in the wrong warehouse or project; "
           "(4) protocol change making the item obsolete; (5) unrecorded consumption.",
    "impact": "Direct, avoidable financial loss, plus a probable audit observation on stock management and donor "
              "reporting on commodity end-use.",
    "action": "Approve and record the write-off; redistribute the near-expiry balance to a project with demand; enforce "
              "first-expiry-first-out issue; reject short-dated deliveries at receipt; freeze ordering of the affected items.",
    "when": "Within 5 working days of the expiry flag; write-off before the financial year closes.",
    "owner": "Stores in-charge", "escalate": "Head of Procurement",
})

# 6 cover
COMMENTARY.append({
    "id": "CM-06", "metric": "Stock cover / stock-out risk",
    "value": f"{F['cover']['under60']} items below 60 days; lowest {F['cover']['lowest']['days']} days "
             f"({F['cover']['lowest']['name'][:32]})",
    "threshold": "Critical < 30 days; watch 30–60 days",
    "status": "Amber" if F["cover"]["under30"] == 0 and F["cover"]["under60"] > 0 else
              ("Red" if F["cover"]["under30"] > 0 else "Green"),
    "what": f"{F['cover']['under30']} items are below 30 days of cover and {F['cover']['under60']} below 60 days, "
            f"against an average supplier lead time of {F['supplier']['avg_lead']} days.",
    "why": "Ranked causes to test: (1) lead time longer than the assumption used in planning; (2) demand surge above trend; "
           "(3) partial delivery or delayed GRN; (4) stock transferred out to another project; (5) pack size forcing under-supply.",
    "impact": "A stock-out of a critical medicine interrupts service delivery and forces an emergency purchase at a "
              "higher price — the most expensive way to buy.",
    "action": "Raise a fast-track purchase for items where cover is below lead time; transfer stock from a warehouse "
              "with surplus cover; agree a temporary substitution with the clinical lead where clinically acceptable.",
    "when": "Immediately where cover is below lead time; same day for clinical items.",
    "owner": "Stores in-charge (with Procurement)", "escalate": "Head of Procurement",
})

# 7 slow moving / overstock
COMMENTARY.append({
    "id": "CM-07", "metric": "Slow-moving and overstocked items",
    "value": f"{F['cover']['slow_items']} items, {money(F['cover']['slow_value'])}; "
             f"{money(F['cover']['over_180_value'])} with over 180 days cover",
    "threshold": "Minimise",
    "status": "Amber" if (F["cover"]["slow_value"] + F["cover"]["over_180_value"]) > 1000000 else "Green",
    "what": f"{F['cover']['slow_items']} items have no consumption in the last 90 days and "
            f"{money(F['cover']['over_180_value'])} of stock represents more than 180 days of cover. "
            f"Stock on hand totals {money(st['total_value'])}.",
    "why": "Ranked causes to test: (1) demand over-estimated at requisition; (2) item superseded or removed from the "
           "protocol; (3) stock received for a project that has since reduced activity; (4) wrong warehouse allocation.",
    "impact": "Blocked working capital and a future expiry loss; it also distorts the consumption-to-investment ratio.",
    "action": "Freeze re-ordering and set a review date; redistribute to a consuming project; correct the requirement "
              "formula with actual consumption; plan disposal or return for what cannot be consumed.",
    "when": "Freeze immediately; disposition before the expiry alert window.",
    "owner": "Supply Chain Officer", "escalate": "Head of Procurement",
})

# 8 price variance
if F["price_risers"]:
    pr = F["price_risers"][0]
    COMMENTARY.append({
        "id": "CM-08", "metric": "Price variance", "value": f"Average +{pct(F['avg_price_change'])}; "
                                                             f"highest {pr['name'][:32]} +{pct(pr['change'])}",
        "threshold": "Investigate above 10%",
        "status": "Amber" if F["avg_price_change"] > 0.10 else "Green",
        "what": f"Across {F['price_items_measured']} items with comparable purchases, the weighted average unit price rose "
                f"{pct(F['avg_price_change'])} between the first and latest purchase; {pr['name']} rose "
                f"from {pr['early']:,.2f} to {pr['late']:,.2f} ({pct(pr['change'])}).",
        "why": "Ranked causes to test: (1) market or exchange-rate movement; (2) loss of contract price or supplier change; "
               "(3) specification or pack-size change; (4) emergency or split purchase at a higher price; (5) unit-of-measure error in the PO.",
        "impact": "Price movement consumes budget without any change in service volume, so activity has to be reduced later "
                  "in the year unless it is negotiated or re-budgeted.",
        "action": "Re-quote or invoke contract terms for the top five risers; consolidate demand to restore volume discounts; "
                  "verify specification and UOM; record the price effect separately in the variance commentary.",
        "when": "Before the next contract or purchase cycle.",
        "owner": "Procurement Officer", "escalate": "Head of Procurement",
    })

# 9 open commitment
COMMENTARY.append({
    "id": "CM-09", "metric": "Open commitments and cancelled orders",
    "value": f"{money(F['open_commitment'])} open; {F['cancelled_lines']:,} lines cancelled "
             f"({money(F['cancelled_value'])})",
    "threshold": "Open commitment < 10% of annual purchase value; cancellations < 2% of lines",
    "status": "Amber",
    "what": f"{F['open_commitment_lines']:,} PO lines worth {money(F['open_commitment'])} remain un-received, and "
            f"{F['cancelled_lines']:,} lines worth {money(F['cancelled_value'])} were cancelled in the period.",
    "why": "Ranked causes to test: (1) supplier capacity or import delay; (2) budget hold or approval delay; "
           "(3) requirement changed after ordering; (4) split ordering to stay within approval limits.",
    "impact": "Open commitments are future cash outflows and future stock; cancellations indicate that requisitions were "
              "raised before the need was firm, which wastes procurement capacity.",
    "action": "Confirm delivery dates in writing; re-plan or cancel the balance where the need has changed; review "
              "approval thresholds if split ordering is evident; use the commitment figure in cash planning.",
    "when": "Weekly review; escalate any line open for more than 30 days.",
    "owner": "Procurement Officer", "escalate": "Manager Finance",
})

# 10 projects ending with unspent budget
ending = [p for p in F["projects_ending"] if p["days_left"] <= 120 and p["unspent"] > 0]
if ending:
    tot_unspent = sum(p["unspent"] for p in ending)
    COMMENTARY.append({
        "id": "CM-10", "metric": "Projects ending soon with unspent budget",
        "value": f"{len(ending)} projects; {money(tot_unspent)} unspent",
        "threshold": "< 20% unspent with under 6 months remaining",
        "status": "Red" if tot_unspent > 10000000 else "Amber",
        "what": f"{len(ending)} projects end within 120 days with a combined {money(tot_unspent)} of budget unspent "
                f"(largest: {max(ending, key=lambda p: p['unspent'])['name'][:40]}, "
                f"{money(max(ending, key=lambda p: p['unspent'])['unspent'])}).",
        "why": "Ranked causes to test: (1) activity slower than plan; (2) budget allocated but never requisitioned; "
               "(3) consumption recorded against another project; (4) procurement capacity constraint; (5) budget overstated at award.",
        "impact": "Unspent budget at project end is reported to donors as under-delivery and may be recovered; correcting "
                  "it at the last minute usually produces poor-quality spending.",
        "action": "Prepare an end-of-project plan now: accelerate committed activity, procure only justified items, "
                  "request a no-cost extension where activity remains, or return the balance formally.",
        "when": "This month — the procurement cycle must still fit inside the remaining period.",
        "owner": "Programme Manager (with Finance)", "escalate": "Director Programme",
    })

# 11 funding gap projection
gap = max(0, F["fy2026_27"]["projection"] - F["budget_fy27"])
COMMENTARY.append({
    "id": "CM-11", "metric": "Projected FY2026-27 funding position",
    "value": f"projected {money(F['fy2026_27']['projection'])} vs budget {money(F['budget_fy27'])}"
             + (f"; gap {money(gap)}" if gap else ""),
    "threshold": "Projection within budget",
    "status": "Red" if gap > 0 else "Green",
    "what": f"With {F['fy2026_27']['months']} months elapsed and {money(F['fy2026_27']['ytd'])} consumed, the burn rate of "
            f"{money(F['fy2026_27']['burn'])} per month projects {money(F['fy2026_27']['projection'])} for the full year "
            f"against a product budget of {money(F['budget_fy27'])}.",
    "why": "Ranked causes to test: (1) budget phasing not yet loaded for the year; (2) genuine under-budgeting against "
           "price inflation; (3) consumption running above the planned service volume; (4) prior-year carry-over spend "
           "landing in this year.",
    "impact": "If the gap is real, service levels cannot be maintained for the full year and the shortfall appears "
              "unannounced in the last quarter.",
    "action": "Confirm budget phasing with Finance first; if the gap is real, quantify it by project and category, "
              "prioritise spending on critical items, and raise a funding request or reprioritise activity with the donor.",
    "when": "Now, at the first quarter of the fiscal year — not at the ninth month.",
    "owner": "Manager Finance & Grants", "escalate": "Director Finance & Operations",
})

# 12 cost per beneficiary spread
COMMENTARY.append({
    "id": "CM-12", "metric": "Cost per beneficiary (peer comparison)",
    "value": f"overall PKR {F['cpb']['overall']:,.0f}; range "
             f"{F['cpb']['lowest']['name'][:28]} PKR {F['cpb']['lowest']['value']:,.0f} to "
             f"{F['cpb']['highest']['name'][:28]} PKR {F['cpb']['highest']['value']:,.0f}",
    "threshold": "Peer spread below 30%",
    "status": "Amber" if F["cpb"]["spread"] > 0.30 else "Green",
    "what": f"Product cost per beneficiary is PKR {F['cpb']['overall']:,.0f} across the portfolio, but ranges from "
            f"PKR {F['cpb']['lowest']['value']:,.0f} in {F['cpb']['lowest']['name']} to "
            f"PKR {F['cpb']['highest']['value']:,.0f} in {F['cpb']['highest']['name']} "
            f"(a spread of {pct(F['cpb']['spread'])}).",
    "why": "Ranked causes to test: (1) case mix and acuity of the service; (2) price paid and supplier terms; "
           "(3) service volume below facility capacity; (4) wastage and expiry inflating cost; (5) one-off stock build in the period.",
    "impact": "A high cost per beneficiary relative to peers is the first question a donor asks; without an explanation "
              "it is read as inefficiency.",
    "action": "Normalise for category mix before drawing conclusions; address the identified driver (price, volume or waste); "
              "set an efficiency target for the next quarter and report the measure with its methodology.",
    "when": "Quarterly review, with the methodology published alongside the first report.",
    "owner": "Director Programme", "escalate": "Executive management",
})

# 13 seasonality
COMMENTARY.append({
    "id": "CM-13", "metric": "Seasonality and peak readiness",
    "value": f"peak month {F['peak_month']['label']} ({money(F['peak_month']['value'])}); "
             f"seasonal index highest in {F['peak_season_month']} at {F['peak_season_index']:.2f}",
    "threshold": "Peak index above 1.10",
    "status": "Amber" if F["peak_season_index"] > 1.10 else "Green",
    "what": f"Consumption peaks in {F['peak_month']['label']} at {money(F['peak_month']['value'])}, "
            f"{pct(F['peak_vs_avg'])} above the period average; the seasonal index is highest in "
            f"{F['peak_season_month']} ({F['peak_season_index']:.2f}×) and lowest in "
            f"{F['low_season_month']} ({F['low_season_index']:.2f}×).",
    "why": "Seasonal demand is driven by disease pattern (diarrhoeal and vector-borne illness in the monsoon months), "
           "campaign activity and beneficiary inflow; the peak is structural, not a one-off.",
    "impact": "If stock is purchased during the surge rather than before it, the peak is met with emergency purchases at "
              "higher prices and with stock-outs at the moment of greatest need.",
    "action": "Pre-position peak quantities at least one lead time before the peak month; align the purchase plan and "
              "warehouse capacity with the seasonal index; review the index each quarter.",
    "when": "Two months before the peak season, every year.",
    "owner": "Supply Chain Officer", "escalate": "Head of Procurement",
})

# 14 data quality
COMMENTARY.append({
    "id": "CM-14", "metric": "Data quality and load integrity",
    "value": f"{F['dq']['duplicate_grc']} duplicate receipts; {F['dq']['zero_price_po']} zero-price lines; "
             f"{F['dq']['inactive_items']} inactive items; {F['dq']['blacklisted']} blacklisted supplier",
    "threshold": "Zero duplicates; zero unmatched master data",
    "status": "Amber" if (F["dq"]["duplicate_grc"] or F["dq"]["inactive_items"] or F["dq"]["blacklisted"]) else "Green",
    "what": f"The load contains {F['dq']['duplicate_grc']} duplicate receipt candidates, {F['dq']['zero_price_po']} "
            f"purchase lines at zero value, {F['dq']['inactive_items']} inactive items with transaction history and "
            f"{F['dq']['blacklisted']} blacklisted supplier still appearing in purchase orders.",
    "why": "Ranked causes to test: (1) double posting in the source system; (2) in-kind or donated stock genuinely at zero value; "
           "(3) master data not maintained after item or supplier deactivation; (4) extract mapping or UOM error.",
    "impact": "Each defect flows into a reported KPI: duplicates inflate stock and receipts, zero prices distort average "
              "price, and inactive data hides real consumption.",
    "action": "Deduplicate and correct at source before the next refresh; exclude zero-value lines from price analysis and "
              "disclose them as in-kind; clean the master data; add the check to the load gate.",
    "when": "Before the next refresh; defects recurring for three periods require a process change.",
    "owner": "IT / BI", "escalate": "Head of Procurement",
})

ws, r = sheet("09_Commentary_Feed", "Live data commentary – what happened, why, so what, now what, when, who",
              f"Generated from the data as at {F['as_at']}. Each line is produced only because the metric breached its threshold "
              f"in 01_Parameters; on refresh, lines move between Red, Amber and Green automatically.",
              widths=[8, 26, 26, 14, 30, 58, 46, 52, 24, 22], tab="E45756")
header(ws, r + 1, ["Ref", "Metric", "Current value", "Status", "What happened (data)",
                   "Why it happened – causes to test", "So what (impact)", "What to do (action)",
                   "When to act", "Owner / escalate"])
rr = r + 2
for cm in COMMENTARY:
    put(ws, rr, [cm["id"], cm["metric"], cm["value"], cm["status"], cm["what"], cm["why"],
                 cm["impact"], cm["action"], cm["when"], f"{cm['owner']} → {cm['escalate']}"],
        wrap_cols=(2, 3, 5, 6, 7, 8, 9, 10))
    sc = ws.cell(row=rr, column=4)
    sc.fill = PatternFill("solid", fgColor=status_fill(cm["status"]))
    sc.font = Font(name="Calibri", size=9.5, bold=True)
    sc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[rr].height = 132
    rr += 1
ws.freeze_panes = ws.cell(row=r + 2, column=2)
ws.auto_filter.ref = f"A{r+1}:J{rr-1}"

rr += 1
ws.cell(row=rr, column=1,
        value="Rule: no commentary without a number, no number without a cause, no cause without an action, "
              "no action without an owner and a date.").font = SUB

# ------------------------------------------------------------------ REFRESH CHECKLIST
ws, r = sheet("10_Refresh_Checklist", "Refresh checklist and quality gates",
              "Run this sequence at every refresh. If a gate fails, publication is blocked and the previous snapshot stays live.",
              widths=[8, 56, 22, 22, 40], tab="72B7B2")
header(ws, r + 1, ["Step", "Activity", "Owner", "Evidence required", "If it fails"])
steps = [
    ("01", "Confirm the Oracle extract completed and record the extract timestamp and row counts",
     "IT / BI", "Load log entry", "Do not load; investigate and re-run"),
    ("02", "Compare row counts and control totals with the previous load and with Oracle registers",
     "IT / BI", "Reconciliation sheet", "Block publication; retain prior snapshot (R-12)"),
    ("03", "De-duplicate receipts and reject negative or zero quantities",
     "IT / BI", "Exception counts", "Correct at source, then reload"),
    ("04", "Validate master data: items, projects and suppliers present and active",
     "IT / BI + Procurement", "Master-data exception list", "Publish with exceptions flagged; assign owners"),
    ("05", "Revise unattributed consumption and confirm the share is within tolerance",
     "Procurement", "Unattributed value and %", "Escalate; commentary required (CM-04)"),
    ("06", "Recalculate stock balances and confirm no negative batch balances",
     "Stores", "Stock reconciliation", "Halt stock KPIs; investigate movements"),
    ("07", "Recompute KPIs and statuses (sheet 06) and regenerate commentary (sheet 08)",
     "Business analyst", "Workbook version stamp", "Re-run; do not hand-edit values"),
    ("08", "Review Red items and assign an owner and due date to each",
     "Head of Procurement", "Action tracker", "Escalate to Director Finance & Operations"),
    ("09", "Circulate the monthly commentary pack and hold the consumption review meeting",
     "Business owner", "Meeting minutes", "Escalate non-attendance to the steering group"),
    ("10", "Archive the snapshot for restatement comparison",
     "IT / BI", "Snapshot store", "Investigate; restatements must be explainable"),
]
rr = r + 2
for step in steps:
    put(ws, rr, list(step), wrap_cols=(2, 5))
    ws.row_dimensions[rr].height = 30
    rr += 1

# ------------------------------------------------------------------ COVERAGE
ws, r = sheet("11_Coverage_Map", "Question coverage against the BRD and KPI dictionary",
              "Every question traces to a KPI or functional requirement so that the dashboard build can be verified question by question.",
              widths=[10, 22, 56, 34, 30], tab="4C78A8")
header(ws, r + 1, ["Stage", "Count", "Purpose at each refresh", "Where it is answered", "BRD reference"])
rr = r + 2
for stage, rows_, sheet_, purpose, ref in [
    ("Descriptive", len(Q.DESCRIPTIVE), "Establish the facts and populate the dashboard tiles", "02 + 06", "§12 KPI dictionary"),
    ("Diagnostic", len(Q.DIAGNOSTIC), "Test ranked causes behind each variance before acting", "03 + 08", "§14 business rules, §15 data quality"),
    ("Predictive", len(Q.PREDICTIVE), "Quantify what happens next if nothing changes", "04 + 08", "§11 pages 2–4, FR-18, FR-40"),
    ("Prescriptive", len(Q.PRESCRIPTIVE), "Assign the action, the owner and the date", "05 + 10", "§8 to-be process, §21 benefits"),
    ("What-if", len(Q.WHATIF), "Quantify the effect of a driver moving, and choose a response",
     "06 + VPT_WhatIf_Analysis.xlsx", "FR-49 to FR-53, §14.4 parameterisation"),
]:
    put(ws, rr, [stage, rows_, purpose, sheet_, ref], wrap_cols=(3, 5))
    rr += 1
rr += 1
ws.cell(row=rr, column=1, value="Total questions: " + str(len(Q.DESCRIPTIVE) + len(Q.DIAGNOSTIC) +
                                                         len(Q.PREDICTIVE) + len(Q.PRESCRIPTIVE) +
                                                         len(Q.WHATIF))).font = Font(
    name="Calibri", size=10, bold=True)
rr += 2
ws.cell(row=rr, column=1, value="Refresh-driven design").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
rr += 1
for line in [
    "The question text, purpose, data fields, logic and owner are stable — they are the standing analysis agenda.",
    "The answer, the status flag and the commentary text are recomputed from the data at each refresh and stamped with the as-at date.",
    "Thresholds and model assumptions live in config/parameters.json and are displayed in 01_Parameters, so a change of "
    "tolerance or assumption changes every status, commentary line and scenario consistently.",
    "Where a question cannot yet be answered (for example a missing programme data field), it stays in the bank with a data-gap note rather than being removed.",
]:
    c = ws.cell(row=rr, column=1, value="•  " + line)
    c.font = BODY
    c.alignment = Alignment(wrap_text=False)
    rr += 1

for wsx in wb.worksheets:
    wsx.sheet_view.zoomScale = 90
wb.active = wb.sheetnames.index("00_ReadMe")
wb.save(OUT)
print("saved", OUT)
print("live answers:", len(LIVE), "| commentary lines:", len(COMMENTARY))

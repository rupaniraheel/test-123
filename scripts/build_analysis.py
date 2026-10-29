#!/usr/bin/env python3
"""
Builds the analysis / dashboard workbook from the generated sample dataset.
Every figure here is derived from VPT_Consumption_Dataset.xlsx, so the workbook
doubles as a numeric check of the BRD's KPI definitions.
"""
import os
from collections import defaultdict
from datetime import date, timedelta
import calendar

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo

BASE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(BASE, "..", "data", "VPT_Consumption_Dataset.xlsx")
OUT = os.path.join(BASE, "..", "data", "VPT_Consumption_Analysis.xlsx")

print("loading source...")
wb_src = load_workbook(SRC, read_only=True, data_only=True)


def rows(name):
    ws = wb_src[name]
    it = ws.iter_rows(values_only=True)
    hdr = next(it)
    # skip an optional note row
    out = []
    for r in it:
        if r[0] is None:
            continue
        out.append(r)
    return hdr, out


_, projects = rows("03_Dim_Project")
_, items = rows("04_Dim_Item")
_, suppliers = rows("05_Dim_Supplier")
_, warehouses = rows("06_Dim_Warehouse")
_, pos = rows("08_Fact_PurchaseOrder")
_, grns = rows("09_Fact_Receipts")
_, issues = rows("10_Fact_Consumption")
_, budgets = rows("11_Fact_Budget")
_, service = rows("12_Fact_ServiceDelivery")
_, stockouts = rows("15_Stockout_Detail")

from datetime import datetime

def norm(rows_, idxs):
    out = []
    for r in rows_:
        r = list(r)
        for i in idxs:
            if isinstance(r[i], datetime):
                r[i] = r[i].date()
        out.append(tuple(r))
    return out

projects = norm(projects, [7, 8])
pos = norm(pos, [1, 3])
grns = norm(grns, [1, 10, 11])
issues = norm(issues, [1, 14])
service = norm(service, [1])

prj = {p[0]: p for p in projects}
itm = {i[0]: i for i in items}
sup = {s[0]: s for s in suppliers}
whs = {w[0]: w for w in warehouses}

TRANSFER = "Inter-store Transfer"
AS_AT = max(r[1] for r in issues)
START = min(r[1] for r in issues)

# ---------------------------------------------------------------- aggregates
monthly = defaultdict(lambda: {"cons": 0.0, "qty": 0.0, "po": 0.0, "recv": 0.0, "lines": 0})
proj_month = defaultdict(float)          # (pid, periodkey) -> consumption
proj_total = defaultdict(float)
proj_fy = defaultdict(float)             # (pid, fy)
proj_cat = defaultdict(float)            # (pid, cat)
item_total = defaultdict(lambda: {"qty": 0.0, "val": 0.0, "cats": None})
item_month = defaultdict(float)          # (item, key)
item_proj = defaultdict(set)
cat_month = defaultdict(float)
cat_total = defaultdict(float)
wh_total = defaultdict(float)
month_qty = defaultdict(float)

for r in issues:
    (_id, idate, charged, orig, item, desc, cat, qty, uom, ucost, tcost,
     wh, grn, batch, exp, itype, sp, by) = r
    key = f"{idate.year}-{idate.month:02d}"
    m = monthly[key]
    m["qty"] += qty
    m["lines"] += 1
    month_qty[key] += qty
    if itype == TRANSFER:
        continue
    val = tcost or 0.0
    m["cons"] += val
    proj_month[(charged, key)] += val
    proj_total[charged] += val
    fy = f"FY{idate.year if idate.month >= 7 else idate.year - 1}-{str((idate.year + 1) if idate.month >= 7 else idate.year)[2:]}"
    proj_fy[(charged, fy)] += val
    proj_cat[(charged, cat)] += val
    it_ = item_total[item]
    it_["qty"] += qty
    it_["val"] += val
    it_["cats"] = cat
    item_month[(item, key)] += val
    item_proj[item].add(charged)
    cat_month[(cat, key)] += val
    cat_total[cat] += val
    wh_total[wh] += val

for r in pos:
    (po, pdate, req, reqdate, pid, sup_, item, desc, cat, qty, uom, price, total,
     cur, status, buyer, lead) = r
    key = f"{pdate.year}-{pdate.month:02d}"
    if status == "Cancelled":
        continue
    monthly[key]["po"] += total or 0.0

for r in grns:
    (_grn, gdate, po, pid, item, desc, qty, uom, wh, batch, mfg, exp, cost, st) = r
    key = f"{gdate.year}-{gdate.month:02d}"
    if po == "OPENING":
        continue          # opening balances are not purchases
    monthly[key]["recv"] += (qty * cost) if cost else 0.0

# budget by project (all years) and util
PRODUCT_CATS = ["Medicines", "Medical Consumables", "Laboratory", "Nutrition", "Equipment", "General Items"]
bud_proj_cat = defaultdict(float)
bud_proj = defaultdict(float)
for r in budgets:
    pid, fy, cat, amt, cur = r
    if cat not in PRODUCT_CATS:
        continue          # HR and programme-operations budget lines are outside product analysis
    bud_proj_cat[(pid, fy, cat)] += amt
    bud_proj[pid] += amt

# beneficiaries by project/month
benef = defaultdict(int)      # (pid,key)
visits = defaultdict(int)
benef_proj = defaultdict(int)
for r in service:
    pid, mdate, key, fy, b, v, adm, sector = r
    benef[(pid, key)] += b
    visits[(pid, key)] += v
    benef_proj[pid] += b

# stock: receipts - issues per item/batch
stock_batch = {}
for r in grns:
    (_grn, gdate, po, pid, item, desc, qty, uom, wh, batch, mfg, exp, cost, st) = r
    k = (item, batch, wh)
    if k in stock_batch:
        stock_batch[k]["qty"] += qty
    else:
        stock_batch[k] = {"item": item, "batch": batch, "wh": wh, "qty": float(qty),
                          "exp": exp, "cost": cost or 0.0, "desc": desc}
for r in issues:
    item, batch, wh, qty = r[4], r[13], r[11], r[7]
    k = (item, batch, wh)
    if k in stock_batch:
        stock_batch[k]["qty"] -= qty
    else:
        stock_batch[k] = {"item": item, "batch": batch, "wh": wh, "qty": -float(qty),
                          "exp": r[14], "cost": r[9] or 0.0, "desc": r[5]}

onhand_item = defaultdict(float)
onhand_item_wh = defaultdict(float)
expired_val = 0.0
expired_qty = 0.0
buckets = {"0-30": 0.0, "31-60": 0.0, "61-90": 0.0, "91-180": 0.0, ">180": 0.0}
for (item, batch, wh), b in stock_batch.items():
    q = b["qty"]
    if q <= 0:
        continue
    onhand_item[item] += q
    onhand_item_wh[(item, wh)] += q
    val = q * b["cost"]
    exp = b["exp"]
    if isinstance(exp, date):
        days = (exp - AS_AT).days
        if days < 0:
            expired_qty += q
            expired_val += val
        elif days <= 30:
            buckets["0-30"] += val
        elif days <= 60:
            buckets["31-60"] += val
        elif days <= 90:
            buckets["61-90"] += val
        elif days <= 180:
            buckets["91-180"] += val
        else:
            buckets[">180"] += val

# last-3-month consumption for days of cover
lookback_start = date(AS_AT.year - (1 if AS_AT.month <= 3 else 0), ((AS_AT.month - 3 - 1) % 12) + 1, 1)
item_recent = defaultdict(float)
for r in issues:
    if r[1] >= lookback_start and r[15] != TRANSFER:
        item_recent[r[4]] += r[7]

# price analysis per item from PO lines
price_hist = defaultdict(list)
for r in pos:
    po, pdate, req, reqdate, pid, sup_, item, desc, cat, qty, uom, price, total, cur, status, buyer, lead = r
    if status == "Cancelled" or not price:
        continue
    price_hist[item].append((pdate, price, sup_, qty))

# supplier performance
sup_perf = defaultdict(lambda: {"pos": 0, "val": 0.0, "lead": [], "full": 0})
po_recv = defaultdict(float)
for r in grns:
    if r[2] and r[2] != "OPENING":
        po_recv[r[2]] += r[6]
for r in pos:
    po, pdate, req, reqdate, pid, sup_, item, desc, cat, qty, uom, price, total, cur, status, buyer, lead = r
    if status == "Cancelled":
        continue
    sp = sup_perf[sup_]
    sp["pos"] += 1
    sp["val"] += total or 0.0
    if po in po_recv:
        sp["full"] += 1
        sp["lead"].append(lead)

# exceptions
unattr_lines = [r for r in issues if r[2] == "GEN-0000"]
unattr_period = defaultdict(float)
for r in unattr_lines:
    unattr_period[f"{r[1].year}-{r[1].month:02d}"] += r[10]
zero_price_po = [r for r in pos if not r[11]]
cancelled_po = [r for r in pos if r[14] == "Cancelled"]
seen = defaultdict(int)
for r in grns:
    seen[(r[0], r[9])] += 1
dupes = {k: v for k, v in seen.items() if v > 1}

# ------------------------------------------------------------------ styling
NAVY = "1F3B57"
TEAL = "2A7F8F"
LIGHT = "EEF3F8"
GREY_TXT = "595959"
MONEY = '#,##0'
MONEY2 = '#,##0.00'
PCT = '0.0%'
NUM = '#,##0'
DATEF = 'dd-mmm-yyyy'

wb = Workbook()
wb.remove(wb.active)

H_FILL = PatternFill("solid", fgColor=NAVY)
H_FONT = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
BODY = Font(name="Calibri", size=9.5)
TITLE = Font(name="Calibri", size=15, bold=True, color=NAVY)
SUB = Font(name="Calibri", size=9, italic=True, color=GREY_TXT)
THIN = Side(style="thin", color="DDDDDD")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def sheet(name, title, subtitle=None, widths=None, tab=None):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = TITLE
    r = 2
    if subtitle:
        ws.cell(row=2, column=1, value=subtitle).font = SUB
        ws.merge_cells(start_row=2, start_column=1,
                       end_row=2, end_column=max(6, min(12, len(widths or [1] * 8))))
        r = 3
    ws.row_dimensions[r].height = 6
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    if tab:
        ws.sheet_properties.tabColor = tab
    return ws, r + 1


def header(ws, row, headers, wrap=True, height=30):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.fill = H_FILL
        cell.font = H_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        cell.border = BOX
    ws.row_dimensions[row].height = height


def put(ws, row, values, fmts=None, bold=False, fill=None):
    for c, v in enumerate(values, start=1):
        cell = ws.cell(row=row, column=c, value=v)
        cell.font = Font(name="Calibri", size=9.5, bold=bold)
        cell.border = BOX
        if fmts and c in fmts:
            cell.number_format = fmts[c]
        if fill:
            cell.fill = PatternFill("solid", fgColor=fill)


def kpi_tile(ws, row, col, label, value, fmt=None, delta=None, width=3):
    c = ws.cell(row=row, column=col, value=label)
    c.font = Font(name="Calibri", size=8.5, bold=True, color=GREY_TXT)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    v = ws.cell(row=row + 1, column=col, value=value)
    v.font = Font(name="Calibri", size=16, bold=True, color=NAVY)
    if fmt:
        v.number_format = fmt
    v.alignment = Alignment(vertical="center")
    if delta:
        d = ws.cell(row=row + 2, column=col, value=delta)
        d.font = Font(name="Calibri", size=8, color=GREY_TXT)
    for rr in range(row, row + (3 if delta else 2)):
        ws.cell(row=rr, column=col).fill = PatternFill("solid", fgColor=LIGHT)
        ws.cell(row=rr, column=col).border = Border(left=Side(style="thick", color=NAVY))
    return row + (3 if delta else 2)


def month_range(first, last):
    out, y, m = [], first.year, first.month
    while (y, m) <= (last.year, last.month):
        out.append(f"{y}-{m:02d}")
        m += 1
        if m == 13:
            y, m = y + 1, 1
    return out


MONTHS = month_range(START, AS_AT)
TOTAL_CONS = sum(v["cons"] for v in monthly.values())
TOTAL_PO = sum(v["po"] for v in monthly.values())
TOTAL_RECV = sum(v["recv"] for v in monthly.values())
TOTAL_BUD = sum(bud_proj.values())
UNATTR = sum(r[10] for r in unattr_lines)
TOTAL_BENEF = sum(benef_proj.values())

# ================================================================= README
ws, r = sheet("README", "VPT – Consumption Analysis Workbook",
              "Analysis pack built from the sample Oracle-style dataset. All figures derive from VPT_Consumption_Dataset.xlsx.",
              widths=[3, 34, 100, 16], tab=NAVY)
notes = [
    ("Purpose", "Demonstrates the KPIs, views and exception logic specified in the BRD, using the sample dataset. Use it to validate definitions before the Oracle build."),
    ("Reporting period", f"{START.strftime('%d-%b-%Y')} to {AS_AT.strftime('%d-%b-%Y')} (27 months, fiscal years FY2024-25 to FY2026-27 Q1)."),
    ("Data as at", AS_AT.strftime('%d-%b-%Y')),
    ("Basis of measurement", "Consumption = inventory issue value excluding inter-store transfers (BRD TR-08). Investment = purchase order value excluding cancelled orders. Received = GRN value."),
    ("Currency", "PKR, excluding sales tax."),
    ("Sheets", "KPI_Dashboard · Trend_Monthly · Project_Analysis · Project_Category · Product_Analysis · Product_Monthly · Category_Analysis · Budget_vs_Consumption · Price_Variance · Stock_Health · Slow_Moving · Supplier_Performance · Cost_per_Beneficiary · Exceptions"),
    ("How to use", "Start at KPI_Dashboard for the headline position, then move to Project_Analysis or Product_Analysis. Exception and stock sheets explain the gaps."),
    ("Headline result", f"Consumption PKR {TOTAL_CONS/1e6:,.1f}m against investment PKR {TOTAL_PO/1e6:,.1f}m — ratio {TOTAL_CONS/TOTAL_PO:.2f}. Unattributed consumption {UNATTR/TOTAL_CONS*100:.1f}% of value."),
    ("Caveat", "Synthetic data. Organisation and supplier names are used for realism only. Replace with Oracle extracts when the extract is available."),
]
r += 1
for label, text in notes:
    ws.cell(row=r, column=2, value=label).font = Font(name="Calibri", size=10, bold=True)
    c = ws.cell(row=r, column=3, value=text)
    c.font = Font(name="Calibri", size=10)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 15 * (len(text) // 105 + 1)
    r += 1

# ================================================================= DASHBOARD
ws, r = sheet("KPI_Dashboard", "Consumption Analysis – Executive Dashboard",
              f"Reporting period {START.strftime('%b-%Y')} to {AS_AT.strftime('%b-%Y')}  |  Data as at {AS_AT.strftime('%d-%b-%Y')}  |  PKR millions where shown",
              widths=[3, 14, 14, 14, 14, 14, 14, 14, 14], tab=TEAL)

row = 5
kpis = [
    ("Consumption value\n(PKR)", TOTAL_CONS, MONEY),
    ("Procurement investment\n(PKR)", TOTAL_PO, MONEY),
    ("Consumption-to-\ninvestment ratio", TOTAL_CONS / TOTAL_PO, '0.00'),
    ("Received value\n(PKR)", TOTAL_RECV, MONEY),
    ("Product budget\n(PKR)", TOTAL_BUD, MONEY),
    ("Budget utilisation", TOTAL_CONS / TOTAL_BUD, PCT),
    ("Unattributed\nconsumption", UNATTR / TOTAL_CONS, PCT),
    ("Cost per beneficiary\n(PKR)", TOTAL_CONS / TOTAL_BENEF, MONEY2),
]
col = 2
for label, val, fmt in kpis:
    c = ws.cell(row=row, column=col, value=label)
    c.font = Font(name="Calibri", size=8.5, bold=True, color=GREY_TXT)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    v = ws.cell(row=row + 1, column=col, value=val)
    v.font = Font(name="Calibri", size=15, bold=True, color=NAVY)
    v.number_format = fmt
    v.alignment = Alignment(horizontal="center")
    for rr in (row, row + 1):
        cell = ws.cell(row=rr, column=col)
        cell.fill = PatternFill("solid", fgColor=LIGHT)
        cell.border = BOX
    col += 1
ws.row_dimensions[row].height = 30
ws.row_dimensions[row + 1].height = 24

# trend table feeding the charts
tr = row + 4
ws.cell(row=tr, column=2, value="Monthly trend – consumption, investment and received value (PKR)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, tr + 1, ["Month", "Consumption", "Investment", "Received", "3-mo avg consumption", "Consumption / investment"])
first_data = tr + 2
for i, key in enumerate(MONTHS):
    rr = first_data + i
    ws.cell(row=rr, column=2, value=f"{calendar.month_abbr[int(key[5:])]}-{key[:4]}")
    ws.cell(row=rr, column=3, value=round(monthly[key]["cons"], 2)).number_format = MONEY
    ws.cell(row=rr, column=4, value=round(monthly[key]["po"], 2)).number_format = MONEY
    ws.cell(row=rr, column=5, value=round(monthly[key]["recv"], 2)).number_format = MONEY
    if i >= 2:
        vals = [monthly[MONTHS[j]]["cons"] for j in range(i - 2, i + 1)]
        ws.cell(row=rr, column=6, value=round(sum(vals) / 3, 2)).number_format = MONEY
    po_v = monthly[key]["po"]
    ws.cell(row=rr, column=7,
            value=round(monthly[key]["cons"] / po_v, 2) if po_v else None).number_format = '0.00'
    for cc in range(2, 8):
        ws.cell(row=rr, column=cc).font = BODY
last_data = first_data + len(MONTHS) - 1

ch = LineChart()
ch.title = "Consumption vs procurement investment by month"
ch.style = 2
ch.y_axis.title = "PKR"
ch.x_axis.title = "Month"
ch.height = 9
ch.width = 22
data = Reference(ws, min_col=3, max_col=4, min_row=tr + 1, max_row=last_data)
cats = Reference(ws, min_col=2, min_row=first_data, max_row=last_data)
ch.add_data(data, titles_from_data=True)
ch.set_categories(cats)
ch.series[0].graphicalProperties.line.width = 22000
ch.series[1].graphicalProperties.line.width = 22000
ch.series[0].graphicalProperties.line.solidFill = NAVY
ch.series[1].graphicalProperties.line.solidFill = TEAL
ch.y_axis.numFmt = '#,##0'
ws.add_chart(ch, f"I{tr}")

# top projects chart
tp = tr
ws.cell(row=tp, column=9, value="").font = SUB
top_proj = sorted(proj_total.items(), key=lambda x: -x[1])[:10]
bar = BarChart()
bar.type = "bar"
bar.title = "Top projects by consumption value"
bar.height = 9
bar.width = 16
bar.y_axis.title = "PKR"
bar.x_axis.title = None
start_row = last_data + 3
ws.cell(row=start_row, column=2, value="Project").font = H_FONT
ws.cell(row=start_row, column=2).fill = H_FILL
ws.cell(row=start_row, column=3, value="Consumption value").font = H_FONT
ws.cell(row=start_row, column=3).fill = H_FILL
for i, (pid, val) in enumerate(top_proj, start=1):
    ws.cell(row=start_row + i, column=2, value=prj.get(pid, [pid, "Unattributed"])[1])
    ws.cell(row=start_row + i, column=3, value=round(val, 2)).number_format = MONEY
data = Reference(ws, min_col=3, min_row=start_row, max_row=start_row + len(top_proj))
cats = Reference(ws, min_col=2, min_row=start_row + 1, max_row=start_row + len(top_proj))
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats)
bar.legend = None
bar.y_axis.numFmt = '#,##0'
bar.series[0].graphicalProperties.solidFill = NAVY
ws.add_chart(bar, f"I{start_row}")

# category chart data
cs = start_row + len(top_proj) + 3
ws.cell(row=cs, column=2, value="Category").font = H_FONT
ws.cell(row=cs, column=2).fill = H_FILL
ws.cell(row=cs, column=3, value="Consumption value").font = H_FONT
ws.cell(row=cs, column=3).fill = H_FILL
cats_sorted = sorted(cat_total.items(), key=lambda x: -x[1])
for i, (cat, val) in enumerate(cats_sorted, start=1):
    ws.cell(row=cs + i, column=2, value=cat)
    ws.cell(row=cs + i, column=3, value=round(val, 2)).number_format = MONEY
pie = BarChart()
pie.type = "bar"
pie.title = "Consumption by product category"
pie.height = 8
pie.width = 16
data = Reference(ws, min_col=3, min_row=cs, max_row=cs + len(cats_sorted))
cats_ref = Reference(ws, min_col=2, min_row=cs + 1, max_row=cs + len(cats_sorted))
pie.add_data(data, titles_from_data=True)
pie.set_categories(cats_ref)
pie.legend = None
pie.y_axis.numFmt = '#,##0'
pie.series[0].graphicalProperties.solidFill = TEAL
ws.add_chart(pie, f"I{cs}")

# attention list
att = cs + len(cats_sorted) + 3
ws.cell(row=att, column=2, value="Attention list (threshold breaches)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, att + 1, ["Exception", "Count / value", "Note"])
att_rows = [
    ("Items below 30 days of cover", sum(1 for i in onhand_item if onhand_item[i] > 0 and
                                        (onhand_item[i] / max(item_recent[i] / 91.0, 0.001)) < 30),
     "Check reorder and open commitments"),
    ("Batches expiring within 60 days (value, PKR)", round(buckets["0-30"] + buckets["31-60"], 2),
     "Redistribute or write off"),
    ("Expired stock not written off (value, PKR)", round(expired_val, 2), "Approve write-off"),
    ("Unattributed consumption (value, PKR)", round(UNATTR, 2), "Correct project coding at issue"),
    ("Zero-value purchase lines", len(zero_price_po), "Confirm in-kind / donated stock"),
    ("Cancelled purchase orders", len(cancelled_po), "Remove from commitment reporting"),
    ("Duplicate receipt candidates", len(dupes), "Deduplicate before loading"),
    ("Item-project-months with unmet demand", len(stockouts), "Review supply and lead time"),
]
for i, (a, b, c) in enumerate(att_rows, start=1):
    ws.cell(row=att + 1 + i, column=2, value=a).font = BODY
    cell = ws.cell(row=att + 1 + i, column=3, value=b)
    cell.font = BODY
    cell.number_format = MONEY if isinstance(b, float) else NUM
    ws.cell(row=att + 1 + i, column=4, value=c).font = BODY

# ================================================================= TREND
ws, r = sheet("Trend_Monthly", "Monthly Trend Analysis",
              "Consumption excludes inter-store transfers. Growth is month-on-month; moving averages use prior months only.",
              widths=[13, 16, 16, 16, 16, 12, 14, 14, 16], tab="2A7F8F")
hr = r + 1
header(ws, hr, ["Month", "Consumption value", "Quantity issued", "Investment (PO)", "Received value",
                "MoM growth", "3-mo moving avg", "12-mo moving avg", "Cost per beneficiary"])
rr = hr + 1
for i, key in enumerate(MONTHS):
    ws.cell(row=rr, column=1, value=f"{calendar.month_abbr[int(key[5:])]}-{key[:4]}")
    ws.cell(row=rr, column=2, value=round(monthly[key]["cons"], 2)).number_format = MONEY
    ws.cell(row=rr, column=3, value=round(month_qty[key])).number_format = NUM
    ws.cell(row=rr, column=4, value=round(monthly[key]["po"], 2)).number_format = MONEY
    ws.cell(row=rr, column=5, value=round(monthly[key]["recv"], 2)).number_format = MONEY
    cur_val = monthly[key]["cons"]
    if i > 0:
        prev = monthly[MONTHS[i - 1]]["cons"]
        if prev:
            ws.cell(row=rr, column=6, value=round(cur_val / prev - 1, 4)).number_format = PCT
    if i >= 2:
        vals = [monthly[MONTHS[j]]["cons"] for j in range(i - 2, i + 1)]
        ws.cell(row=rr, column=7, value=round(sum(vals) / 3, 2)).number_format = MONEY
    if i >= 11:
        vals = [monthly[MONTHS[j]]["cons"] for j in range(i - 11, i + 1)]
        ws.cell(row=rr, column=8, value=round(sum(vals) / 12, 2)).number_format = MONEY
    b = sum(benef[(p, key)] for p in prj)
    if b:
        ws.cell(row=rr, column=9, value=round(cur_val / b, 2)).number_format = MONEY2
    for cc in range(1, 10):
        ws.cell(row=rr, column=cc).font = BODY
    rr += 1
ws.cell(row=rr, column=1, value="Total").font = Font(name="Calibri", size=9.5, bold=True)
ws.cell(row=rr, column=2, value=round(sum(monthly[k]["cons"] for k in MONTHS), 2)).number_format = MONEY
ws.cell(row=rr, column=3, value=round(sum(month_qty[k] for k in MONTHS))).number_format = NUM
ws.cell(row=rr, column=4, value=round(sum(monthly[k]["po"] for k in MONTHS), 2)).number_format = MONEY
ws.cell(row=rr, column=5, value=round(sum(monthly[k]["recv"] for k in MONTHS), 2)).number_format = MONEY
for cc in range(1, 6):
    ws.cell(row=rr, column=cc).font = Font(name="Calibri", size=9.5, bold=True)
    ws.cell(row=rr, column=cc).fill = PatternFill("solid", fgColor=LIGHT)
ch = LineChart()
ch.title = "Monthly consumption value and trend"
ch.height = 9
ch.width = 24
ch.y_axis.title = "PKR"
data = Reference(ws, min_col=2, max_col=2, min_row=hr, max_row=rr - 1)
data2 = Reference(ws, min_col=7, max_col=7, min_row=hr, max_row=rr - 1)
cats = Reference(ws, min_col=1, min_row=hr + 1, max_row=rr - 1)
ch.add_data(data, titles_from_data=True)
ch.add_data(data2, titles_from_data=True)
ch.set_categories(cats)
ch.series[0].graphicalProperties.solidFill = NAVY
ch.series[0].graphicalProperties.line.width = 22000
ch.series[1].graphicalProperties.line.solidFill = TEAL
ch.series[1].graphicalProperties.line.width = 18000
ch.y_axis.numFmt = '#,##0'
ws.add_chart(ch, f"A{rr+3}")

# seasonality block
sr = rr + 25
ws.cell(row=sr, column=1, value="Seasonality – average consumption value by calendar month").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, sr + 1, ["Calendar month", "Avg consumption value", "Avg quantity issued", "Index (Jan-2025 = 100)"])
season_val = defaultdict(list)
season_qty = defaultdict(list)
for key in MONTHS:
    mth = int(key[5:])
    season_val[mth].append(monthly[key]["cons"])
    season_qty[mth].append(month_qty[key])
base = sum(season_val[1]) / max(len(season_val[1]), 1)
rown = sr + 2
for mth in range(1, 13):
    ws.cell(row=rown, column=1, value=calendar.month_name[mth])
    avg = sum(season_val[mth]) / max(len(season_val[mth]), 1)
    ws.cell(row=rown, column=2, value=round(avg, 2)).number_format = MONEY
    ws.cell(row=rown, column=3, value=round(sum(season_qty[mth]) / max(len(season_qty[mth]), 1))).number_format = NUM
    ws.cell(row=rown, column=4, value=round(avg / base * 100, 1)).number_format = '0.0'
    for cc in range(1, 5):
        ws.cell(row=rown, column=cc).font = BODY
    rown += 1

# ================================================================= PROJECT ANALYSIS
ws, r = sheet("Project_Analysis", "Project-wise Consumption versus Investment",
              "Consumption excludes transfers. Ratio = consumption / investment. Variance = investment - consumption. Cumulative covers the full reporting window.",
              widths=[10, 40, 20, 14, 16, 16, 10, 16, 16, 14, 16, 14], tab="1F3B57")
hr = r + 1
header(ws, hr, ["Project ID", "Project", "Donor", "Sector", "Consumption value", "Investment (PO)",
                "Ratio", "Variance", "Consumption FY2026-27", "Beneficiaries", "Cost per beneficiary", "Budget"])
rr = hr + 1
pa_first = rr
for pid, val in sorted(proj_total.items(), key=lambda x: -x[1]):
    p = prj.get(pid)
    po_val = sum(x[12] or 0 for x in pos if x[4] == pid and x[14] != "Cancelled")
    name = p[1] if p else "Unattributed (generic cost centre)"
    donor = p[3] if p else "—"
    sector = p[4] if p else "—"
    b = benef_proj.get(pid, 0)
    fy_val = proj_fy.get((pid, "FY2026-27"), 0)
    bud_v = bud_proj.get(pid, 0)
    put(ws, rr, [pid, name, donor, sector, round(val, 2), round(po_val, 2),
                 round(val / po_val, 2) if po_val else None, round(po_val - val, 2),
                 round(fy_val, 2), b, round(val / b, 2) if b else None, round(bud_v, 2)],
        fmts={5: MONEY, 6: MONEY, 7: '0.00', 8: MONEY, 9: MONEY, 10: NUM, 11: MONEY2, 12: MONEY})
    rr += 1
_tc = sum(proj_total.values())
_tp = sum(sum(x[12] or 0 for x in pos if x[4] == pid and x[14] != "Cancelled") for pid in proj_total)
_tfy = sum(proj_fy.get((pid, "FY2026-27"), 0) for pid in proj_total)
_tb = sum(benef_proj.get(pid, 0) for pid in proj_total)
_tbud = sum(bud_proj.get(pid, 0) for pid in proj_total)
put(ws, rr, ["", "TOTAL", "", "", round(_tc, 2), round(_tp, 2), round(_tc / _tp, 2), round(_tp - _tc, 2),
             round(_tfy, 2), _tb, round(_tc / _tb, 2), round(_tbud, 2)],
    fmts={5: MONEY, 6: MONEY, 7: '0.00', 8: MONEY, 9: MONEY, 10: NUM, 11: MONEY2, 12: MONEY}, bold=True, fill=LIGHT)

note_r = rr + 2
ws.cell(row=note_r, column=1, value="Project progress against life cycle").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, note_r + 1, ["Project ID", "Project", "Start", "End", "Elapsed %", "Status",
                        "Consumption to date", "Budget", "Budget utilised", "Days remaining"])
rr = note_r + 2
for pid, p in prj.items():
    start, end = p[7], p[8]
    total_days = (end - start).days
    elapsed = min(max((AS_AT - start).days, 0), total_days)
    elapsed_pct = elapsed / total_days if total_days else 0
    cons_v = proj_total.get(pid, 0)
    bud_v = bud_proj.get(pid, 0)
    put(ws, rr, [pid, p[1], start, end, round(elapsed_pct, 4), p[9], round(cons_v, 2), round(bud_v, 2),
                 round(cons_v / bud_v, 4) if bud_v else None, max((end - AS_AT).days, 0)],
        fmts={3: DATEF, 4: DATEF, 5: PCT, 7: MONEY, 8: MONEY, 9: PCT, 10: NUM})
    rr += 1

# ================================================================= PROJECT x CATEGORY
ws, r = sheet("Project_Category", "Project × Category Consumption Matrix",
              "Values in PKR. Totals reconcile to Project_Analysis.",
              widths=[10, 38] + [16] * 6 + [16], tab="4C78A8")
cats_list = ["Medicines", "Medical Consumables", "Laboratory", "Nutrition", "Equipment", "General Items"]
hr = r + 1
header(ws, hr, ["Project ID", "Project"] + cats_list + ["Total"])
rr = hr + 1
f0 = rr
for pid in sorted(proj_total, key=lambda x: -proj_total[x]):
    name = prj[pid][1] if pid in prj else "Unattributed"
    row = [pid, name]
    for cat in cats_list:
        row.append(round(proj_cat.get((pid, cat), 0), 2))
    row.append(round(sum(proj_cat.get((pid, c), 0) for c in cats_list), 2))
    put(ws, rr, row, fmts={i: MONEY for i in range(3, 10)})
    rr += 1
tot = ["", "TOTAL"] + [round(sum(proj_cat.get((pid, c), 0) for pid in proj_total), 2) for c in cats_list]
tot.append(round(sum(proj_total.values()), 2))
put(ws, rr, tot, fmts={i: MONEY for i in range(3, 10)}, bold=True, fill=LIGHT)

# ================================================================= PRODUCT ANALYSIS
ws, r = sheet("Product_Analysis", "Product-wise Consumption and Investment",
              "Ranked by consumption value. Share and cumulative share support Pareto analysis. Ratio compares consumption with purchase for the same item.",
              widths=[11, 42, 20, 14, 14, 16, 16, 16, 10, 10, 12, 12], tab="2A7F8F")
hr = r + 1
header(ws, hr, ["Item ID", "Item description", "Category", "UOM", "Qty consumed", "Consumption value",
                "Investment (PO)", "Ratio", "Share", "Cum share", "Avg unit price", "Projects"])
rr = hr + 1
p_first = rr
po_item = defaultdict(float)
po_item_qty = defaultdict(float)
for x in pos:
    if x[14] == "Cancelled":
        continue
    po_item[x[6]] += x[12] or 0.0
    po_item_qty[x[6]] += x[9] or 0.0
total_item_val = sum(v["val"] for v in item_total.values())
cum = 0.0
for item, d in sorted(item_total.items(), key=lambda x: -x[1]["val"]):
    pir = po_item.get(item, 0)
    share = d["val"] / total_item_val if total_item_val else 0
    cum += share
    put(ws, rr, [item, itm[item][1], itm[item][2], itm[item][4], round(d["qty"]), round(d["val"], 2),
                 round(pir, 2), round(d["val"] / pir, 2) if pir else None, round(share, 4), round(min(cum, 1), 4),
                 round(d["val"] / d["qty"], 2) if d["qty"] else 0, len(item_proj[item])],
        fmts={5: NUM, 6: MONEY, 7: MONEY, 8: '0.00', 9: PCT, 10: PCT, 11: MONEY2, 12: NUM})
    rr += 1
_tiq = sum(v["qty"] for v in item_total.values())
_pir_t = sum(po_item.values())
put(ws, rr, ["", "TOTAL", "", "", round(_tiq), round(total_item_val, 2), round(_pir_t, 2),
             round(total_item_val / _pir_t, 2), 1, "", "", ""],
    fmts={5: NUM, 6: MONEY, 7: MONEY, 8: '0.00', 9: PCT}, bold=True, fill=LIGHT)

# ================================================================= PRODUCT MONTHLY (top 15)
ws, r = sheet("Product_Monthly", "Top 15 Products – Monthly Consumption Value",
              "Month-wise value in PKR for the fifteen highest-value products; a heat map would be applied on the Power BI page.",
              widths=[42] + [13] * 27 + [15], tab="4C78A8")
hr = r + 1
header(ws, hr, ["Item description"] + [f"{calendar.month_abbr[int(k[5:])]}-{k[2:4]}" for k in MONTHS] + ["Total"], height=34)
rr = hr + 1
top_items = [i for i, _ in sorted(item_total.items(), key=lambda x: -x[1]["val"])[:15]]
for item in top_items:
    row = [itm[item][1]]
    for key in MONTHS:
        row.append(round(item_month.get((item, key), 0), 2) or None)
    row.append(round(sum(item_month.get((item, k), 0) for k in MONTHS), 2))
    put(ws, rr, row, fmts={i: MONEY for i in range(2, 30)})
    rr += 1

# ================================================================= CATEGORY
ws, r = sheet("Category_Analysis", "Consumption by Category and Month",
              "Values in PKR. Inter-store transfers are excluded.",
              widths=[26] + [13] * 27 + [16], tab="2A7F8F")
hr = r + 1
header(ws, hr, ["Category"] + [f"{calendar.month_abbr[int(k[5:])]}-{k[2:4]}" for k in MONTHS] + ["Total"], height=34)
rr = hr + 1
c0 = rr
for cat, _ in sorted(cat_total.items(), key=lambda x: -x[1]):
    row = [cat] + [round(cat_month.get((cat, k), 0), 2) or None for k in MONTHS]
    row.append(round(cat_total.get(cat, 0), 2))
    put(ws, rr, row, fmts={i: MONEY for i in range(2, 30)})
    rr += 1
tot = ["TOTAL"] + [round(sum(cat_month.get((c, k), 0) for c in cat_total), 2) for k in MONTHS]
tot.append(round(sum(cat_total.values()), 2))
put(ws, rr, tot, fmts={i: MONEY for i in range(2, 30)}, bold=True, fill=LIGHT)

# ================================================================= BUDGET
ws, r = sheet("Budget_vs_Consumption", "Budget versus Consumption by Project, Fiscal Year and Category",
              "Product-category budget only (medicines, consumables, laboratory, nutrition, equipment, general); HR and programme-operations lines excluded. Utilisation = consumption / budget.",
              widths=[10, 34, 12, 22, 16, 16, 16, 12, 14], tab="E45756")
hr = r + 1
header(ws, hr, ["Project ID", "Project", "Fiscal year", "Expense category", "Budget", "Consumption",
                "Variance", "Utilisation", "Flag"])
rr = hr + 1
b0 = rr
for (pid, fy, cat), amt in sorted(bud_proj_cat.items()):
    cons = 0.0
    for issue in issues:
        if issue[2] != pid or issue[15] == TRANSFER:
            continue
        d = issue[1]
        f = f"FY{d.year if d.month >= 7 else d.year - 1}-{str((d.year + 1) if d.month >= 7 else d.year)[2:]}"
        if f == fy and issue[6] == cat:
            cons += issue[10] or 0.0
    util = cons / amt if amt else None
    flag = "No budget" if amt == 0 else ("Over budget" if util and util > 1 else
                                        ("Watch" if util and util > 0.85 else "Within"))
    put(ws, rr, [pid, prj[pid][1] if pid in prj else pid, fy, cat, round(amt, 2), round(cons, 2),
                 round(amt - cons, 2), round(util, 4) if util is not None else None, flag],
        fmts={5: MONEY, 6: MONEY, 7: MONEY, 8: PCT})
    rr += 1
_ba = sum(a for a in bud_proj_cat.values())
_bc = sum(ws.cell(row=x, column=6).value or 0 for x in range(b0, rr))
put(ws, rr, ["", "TOTAL", "", "", round(_ba, 2), round(_bc, 2), round(_ba - _bc, 2),
             round(_bc / _ba, 4) if _ba else None, ""],
    fmts={5: MONEY, 6: MONEY, 7: MONEY, 8: PCT}, bold=True, fill=LIGHT)

# ================================================================= PRICE VARIANCE
ws, r = sheet("Price_Variance", "Unit Price Movement and Variance",
              "Purchase prices exclude zero-value (donated) lines. Variance compares the latest purchase price with the item's standard cost and with its first purchase in the window.",
              widths=[11, 42, 20, 14, 16, 16, 16, 14, 14, 16], tab="E45756")
hr = r + 1
header(ws, hr, ["Item ID", "Item description", "Category", "Standard cost", "First price", "Latest price",
                "Average price", "Change vs first", "Change vs standard", "Qty purchased"])
rr = hr + 1
v0 = rr
for item, hist in sorted(price_hist.items(), key=lambda x: -len(x[1])):
    hist_sorted = sorted(hist, key=lambda h: h[0])
    first = hist_sorted[0][1]
    latest = hist_sorted[-1][1]
    qty_sum = sum(h[3] for h in hist_sorted)
    avg = sum(h[1] * h[3] for h in hist_sorted) / qty_sum if qty_sum else 0
    std = itm[item][6]
    c_first = (latest / first - 1) if first else None
    c_std = (latest / std - 1) if std else None
    put(ws, rr, [item, itm[item][1], itm[item][2], std, first, latest, round(avg, 2),
                 round(c_first, 4) if c_first is not None else None,
                 round(c_std, 4) if c_std is not None else None, round(qty_sum)],
        fmts={4: MONEY2, 5: MONEY2, 6: MONEY2, 7: MONEY2, 8: PCT, 9: PCT, 10: NUM})
    rr += 1
_cf = [ws.cell(row=x, column=8).value for x in range(v0, rr) if isinstance(ws.cell(row=x, column=8).value, float)]
_cs = [ws.cell(row=x, column=9).value for x in range(v0, rr) if isinstance(ws.cell(row=x, column=9).value, float)]
put(ws, rr, ["", "AVERAGE / TOTAL", "", "", "", "", "",
             round(sum(_cf) / len(_cf), 4) if _cf else None,
             round(sum(_cs) / len(_cs), 4) if _cs else None,
             sum(ws.cell(row=x, column=10).value or 0 for x in range(v0, rr))],
    fmts={8: PCT, 9: PCT, 10: NUM}, bold=True, fill=LIGHT)

# ================================================================= STOCK HEALTH
ws, r = sheet("Stock_Health", "Stock on Hand, Cover and Expiry Risk",
              f"Balances as at {AS_AT.strftime('%d-%b-%Y')}. Days of cover uses consumption over the last 90 days.",
              widths=[11, 42, 20, 12, 14, 14, 12, 14, 14, 14], tab="F2A65A")
hr = r + 1
header(ws, hr, ["Item ID", "Item description", "Category", "On hand", "Stock value",
                "Consumed (90 days)", "Days of cover", "Expiring ≤ 60 days", "Expired qty", "Cover status"])
rr = hr + 1
sh0 = rr
exp_item = defaultdict(float)
exp_item_val = defaultdict(float)
for (item, batch, wh), b in stock_batch.items():
    if b["qty"] > 0 and isinstance(b["exp"], date):
        days = (b["exp"] - AS_AT).days
        if days < 0:
            exp_item[item] += b["qty"]
            exp_item_val[item] += b["qty"] * b["cost"]
for item, qty in sorted(onhand_item.items(), key=lambda x: -x[1] * itm[x[0]][6]):
    if qty <= 0:
        continue
    recent = item_recent.get(item, 0)
    std = itm[item][6]
    cover = (qty / (recent / 91.0)) if recent else None
    status = ("Slow-moving" if cover is None else
              "Critical" if cover < 15 else "Low" if cover < 30 else "Adequate")
    put(ws, rr, [item, itm[item][1], itm[item][2], round(qty), round(qty * std, 2), round(recent),
                 round(cover) if cover is not None else None,
                 round(exp_item_val.get(item, 0), 2), round(exp_item.get(item, 0)), status],
        fmts={4: NUM, 5: MONEY, 6: NUM, 7: NUM, 8: MONEY, 9: NUM})
    rr += 1
put(ws, rr, ["", "TOTAL", "",
             round(sum(ws.cell(row=x, column=4).value or 0 for x in range(sh0, rr))),
             round(sum(ws.cell(row=x, column=5).value or 0 for x in range(sh0, rr)), 2),
             round(sum(ws.cell(row=x, column=6).value or 0 for x in range(sh0, rr))), "",
             round(sum(ws.cell(row=x, column=8).value or 0 for x in range(sh0, rr)), 2),
             round(sum(ws.cell(row=x, column=9).value or 0 for x in range(sh0, rr))), ""],
    fmts={4: NUM, 5: MONEY, 6: NUM, 8: MONEY, 9: NUM}, bold=True, fill=LIGHT)

er = rr + 3
ws.cell(row=er, column=1, value="Expiry exposure by bucket (value of stock held)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, er + 1, ["Bucket (days to expiry)", "Value at risk", "Share of stock value"])
stock_value_total = sum(b["qty"] * b["cost"] for b in stock_batch.values() if b["qty"] > 0)
rown = er + 2
for label, val in [("Expired (already past)", expired_val)] + list(buckets.items()):
    ws.cell(row=rown, column=1, value=label).font = BODY
    ws.cell(row=rown, column=2, value=round(val, 2)).number_format = MONEY
    ws.cell(row=rown, column=3, value=round(val / stock_value_total, 4) if stock_value_total else 0).number_format = PCT
    for cc in (2, 3):
        ws.cell(row=rown, column=cc).font = BODY
    rown += 1

# ================================================================= SLOW MOVING
ws, r = sheet("Slow_Moving", "Slow-moving and Non-moving Stock",
              "Items with stock on hand but no or low consumption in the lookback period. Value at standard cost.",
              widths=[11, 42, 20, 14, 14, 16, 18, 18], tab="F2A65A")
hr = r + 1
header(ws, hr, ["Item ID", "Item description", "Category", "On hand", "Stock value",
                "Consumed (90 days)", "Consumed (last 12 mo)", "Classification"])
rr = hr + 1
last12 = date(AS_AT.year - 1 if AS_AT.month <= 12 else AS_AT.year, AS_AT.month, 1)
item_12 = defaultdict(float)
for x in issues:
    if x[1] >= last12 and x[15] != TRANSFER:
        item_12[x[4]] += x[7]
s0 = rr
for item, qty in sorted(onhand_item.items(), key=lambda x: -x[1]):
    if qty <= 0:
        continue
    r90 = item_recent.get(item, 0)
    r12 = item_12.get(item, 0)
    if r90 > 0:
        continue
    put(ws, rr, [item, itm[item][1], itm[item][2], round(qty), round(qty * itm[item][6], 2),
                 round(r90), round(r12), "Non-moving 3 months" if r12 > 0 else "Non-moving 12 months"],
        fmts={4: NUM, 5: MONEY, 6: NUM, 7: NUM})
    rr += 1
put(ws, rr, ["", "TOTAL", "",
             round(sum(ws.cell(row=x, column=4).value or 0 for x in range(s0, rr))),
             round(sum(ws.cell(row=x, column=5).value or 0 for x in range(s0, rr)), 2), "", "", ""],
    fmts={4: NUM, 5: MONEY}, bold=True, fill=LIGHT)

# ================================================================= SUPPLIER
ws, r = sheet("Supplier_Performance", "Supplier Performance",
              "Lead time from PO date to first receipt. Fulfilment counts orders with at least one receipt.",
              widths=[12, 38, 14, 12, 18, 16, 16, 14], tab="4C78A8")
hr = r + 1
header(ws, hr, ["Supplier ID", "Supplier", "City", "PO lines", "Purchase value",
                "Avg lead time (days)", "Agreed lead time", "% fulfilled"])
rr = hr + 1
su0 = rr
for sid, d in sorted(sup_perf.items(), key=lambda x: -x[1]["val"]):
    avg_lead = sum(d["lead"]) / len(d["lead"]) if d["lead"] else None
    put(ws, rr, [sid, sup[sid][1], sup[sid][2], d["pos"], round(d["val"], 2),
                 round(avg_lead) if avg_lead else "", sup[sid][5],
                 d["full"] / d["pos"] if d["pos"] else 0],
        fmts={4: NUM, 5: MONEY, 6: NUM, 7: NUM, 8: PCT})
    rr += 1
put(ws, rr, ["", "TOTAL", "",
             round(sum(ws.cell(row=x, column=4).value or 0 for x in range(su0, rr))),
             round(sum(ws.cell(row=x, column=5).value or 0 for x in range(su0, rr)), 2), "", "", ""],
    fmts={4: NUM, 5: MONEY}, bold=True, fill=LIGHT)

# ================================================================= COST PER BENEFICIARY
ws, r = sheet("Cost_per_Beneficiary", "Cost per Beneficiary and per Visit",
              "Product consumption value divided by programme service counts for the same project and month.",
              widths=[10, 38, 12, 16, 14, 16, 16, 16], tab="72B7B2")
hr = r + 1
header(ws, hr, ["Project ID", "Project", "Month", "Consumption value", "Beneficiaries",
                "Cost per beneficiary", "Patient visits", "Cost per visit"])
rr = hr + 1
d0 = rr
for (pid, key), b in sorted(benef.items()):
    cons = proj_month.get((pid, key), 0)
    if b == 0:
        continue
    y, mth = int(key[:4]), int(key[5:])
    v = visits.get((pid, key), 0)
    put(ws, rr, [pid, prj[pid][1] if pid in prj else pid, f"{calendar.month_abbr[mth]}-{y}",
                 round(cons, 2), b, round(cons / b, 2) if b else None, v,
                 round(cons / v, 2) if v else None],
        fmts={4: MONEY, 5: NUM, 6: MONEY2, 7: NUM, 8: MONEY2})
    rr += 1
_tcons = sum(ws.cell(row=x, column=4).value or 0 for x in range(d0, rr))
_tb = sum(ws.cell(row=x, column=5).value or 0 for x in range(d0, rr))
_tv = sum(ws.cell(row=x, column=7).value or 0 for x in range(d0, rr))
put(ws, rr, ["", "TOTAL / AVERAGE", "", round(_tcons, 2), _tb,
             round(_tcons / _tb, 2) if _tb else None, _tv, round(_tcons / _tv, 2) if _tv else None],
    fmts={4: MONEY, 5: NUM, 6: MONEY2, 7: NUM, 8: MONEY2}, bold=True, fill=LIGHT)

# ================================================================= EXCEPTIONS
ws, r = sheet("Exceptions", "Exceptions and Data Quality",
              "Each block corresponds to an exception type specified in the BRD (FR-29, FR-30, DQ-01 to DQ-06).",
              widths=[38, 18, 18, 44], tab="E45756")
row = r + 1
ws.cell(row=row, column=1, value="Summary of exceptions").font = Font(name="Calibri", size=11, bold=True, color=NAVY)
header(ws, row + 1, ["Exception type", "Count", "Value (PKR)", "Required handling"])
summary = [
    ("Unattributed consumption (generic project code)", len(unattr_lines), round(UNATTR, 2),
     "Report separately; correct project coding at issue (FR-30)"),
    ("Zero-value / donated purchase lines", len(zero_price_po),
     round(sum(x[12] or 0 for x in zero_price_po), 2), "Exclude from price variance; disclose as in-kind"),
    ("Cancelled purchase orders", len(cancelled_po), round(sum(x[12] or 0 for x in cancelled_po), 2),
     "Exclude from investment and commitment (TR-05)"),
    ("Duplicate receipt candidates", len(dupes), None, "Deduplicate on GRN + batch before loading (TR-06)"),
    ("Item-project-months with unmet demand", len(stockouts),
     round(sum(x[3] for x in stockouts), 2), "Review stock-out cause and lead time (KPI-22)"),
    ("Expired stock not written off", None, round(expired_val, 2), "Approve write-off and adjust stock (KPI-21)"),
]
rr = row + 2
for a, b, c, d in summary:
    ws.cell(row=rr, column=1, value=a).font = BODY
    ws.cell(row=rr, column=2, value=b).font = BODY
    cell = ws.cell(row=rr, column=3, value=c if c is not None else "")
    cell.font = BODY
    cell.number_format = MONEY
    ws.cell(row=rr, column=4, value=d).font = BODY
    rr += 1

rr += 2
ws.cell(row=rr, column=1, value="Unattributed consumption by month").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, rr + 1, ["Month", "Unattributed value", "Total consumption", "Share"])
u0 = rr + 2
for i, key in enumerate(MONTHS):
    ws.cell(row=u0 + i, column=1, value=f"{calendar.month_abbr[int(key[5:])]}-{key[:4]}").font = BODY
    cell = ws.cell(row=u0 + i, column=2, value=round(unattr_period.get(key, 0), 2))
    cell.font = BODY
    cell.number_format = MONEY
    cell = ws.cell(row=u0 + i, column=3, value=round(monthly[key]["cons"], 2))
    cell.font = BODY
    cell.number_format = MONEY
    tot_v = monthly[key]["cons"]
    cell = ws.cell(row=u0 + i, column=4,
                   value=round(unattr_period.get(key, 0) / tot_v, 4) if tot_v else None)
    cell.font = BODY
    cell.number_format = PCT

rr = u0 + len(MONTHS) + 2
ws.cell(row=rr, column=1, value="Unattributed consumption by item (top 15)").font = Font(
    name="Calibri", size=11, bold=True, color=NAVY)
header(ws, rr + 1, ["Item", "Value", "Lines", "Share of unattributed"])
u_item = defaultdict(lambda: [0.0, 0])
for x in unattr_lines:
    u_item[x[5]][0] += x[10] or 0
    u_item[x[5]][1] += 1
rn = rr + 2
for name, (val, cnt) in sorted(u_item.items(), key=lambda y: -y[1][0])[:15]:
    ws.cell(row=rn, column=1, value=name).font = BODY
    cell = ws.cell(row=rn, column=2, value=round(val, 2))
    cell.font = BODY
    cell.number_format = MONEY
    ws.cell(row=rn, column=3, value=cnt).font = BODY
    cell = ws.cell(row=rn, column=4, value=round(val / UNATTR, 4) if UNATTR else None)
    cell.font = BODY
    cell.number_format = PCT
    rn += 1

# ================================================================= write
for wsx in wb.worksheets:
    wsx.sheet_view.zoomScale = 90
wb.active = wb.sheetnames.index("KPI_Dashboard")
wb.save(OUT)
print("saved", OUT)
print("cons", round(TOTAL_CONS), "po", round(TOTAL_PO), "budget", round(TOTAL_BUD),
      "unattr%", round(UNATTR / TOTAL_CONS * 100, 2), "benef", TOTAL_BENEF)

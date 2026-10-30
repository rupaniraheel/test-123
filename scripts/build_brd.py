#!/usr/bin/env python3
"""Builds the extensive BRD for the VPT Consumption Analysis dashboard."""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import json

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "..", "data", "summary.json")) as f:
    S = json.load(f)
with open(os.path.join(BASE, "..", "config", "parameters.json")) as f:
    CFG = json.load(f)
with open(os.path.join(BASE, "..", "data", "whatif_results.json")) as f:
    WF = json.load(f)

NAVY = RGBColor(0x1F, 0x3B, 0x57)
GREY = RGBColor(0x59, 0x59, 0x59)
ACCENT = "1F3B57"
LIGHT = "EEF3F8"
LIGHTER = "F7F9FB"

doc = Document()

st = doc.styles["Normal"]
st.font.name = "Calibri"
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
st.paragraph_format.space_after = Pt(5)
st.paragraph_format.line_spacing = 1.12

for s in ("Heading 1", "Heading 2", "Heading 3"):
    h = doc.styles[s]
    h.font.name = "Calibri"
    h.font.color.rgb = NAVY
    h.font.bold = True
doc.styles["Heading 1"].font.size = Pt(14)
doc.styles["Heading 1"].paragraph_format.space_before = Pt(14)
doc.styles["Heading 1"].paragraph_format.space_after = Pt(6)
doc.styles["Heading 2"].font.size = Pt(11.5)
doc.styles["Heading 2"].paragraph_format.space_before = Pt(10)
doc.styles["Heading 2"].paragraph_format.space_after = Pt(4)
doc.styles["Heading 3"].font.size = Pt(10.5)
doc.styles["Heading 3"].font.color.rgb = GREY
doc.styles["Heading 3"].paragraph_format.space_before = Pt(8)

sec = doc.sections[0]
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(1.9)
sec.right_margin = Cm(1.9)
USABLE = 17.2


# ------------------------------------------------------------------ helpers
def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:fill"), color)
    tcPr.append(sh)


def set_cell_text(cell, text, bold=False, size=8.5, color=None, align=None):
    cell.text = ""
    p = cell.paragraphs[0]
    r = p.add_run(str(text))
    r.bold = bold
    r.font.size = Pt(size)
    r.font.name = "Calibri"
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    if align:
        p.alignment = align


def rule(p, color=ACCENT):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "8")
    bot.set(qn("w:space"), "3")
    bot.set(qn("w:color"), color)
    b.append(bot)
    pPr.append(b)


def para(text="", bold=False, italic=False, size=None, color=None, space_after=5,
         align=None, indent=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(space_after)
    if align:
        p.alignment = align
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    return p


def bullet(text, bold_lead="", level=0, style="List Bullet"):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        p.add_run(bold_lead).bold = True
    p.add_run(text)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.left_indent = Cm(0.6 + 0.5 * level)
    return p


def numbered(text, bold_lead=""):
    return bullet(text, bold_lead=bold_lead, style="List Number")


def table(headers, rows, widths, font=8.5, header_fill=ACCENT, zebra=True,
          align_center_cols=(), bold_first_col=False):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        set_cell_text(c, h, bold=True, size=font, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, header_fill)
    # repeat header row across pages
    trPr = t.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:tblHeader")
    trPr.append(th)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            set_cell_text(cells[i], val, size=font,
                          bold=(bold_first_col and i == 0),
                          align=WD_ALIGN_PARAGRAPH.CENTER if i in align_center_cols else None)
            if zebra and ri % 2 == 1:
                shade(cells[i], LIGHTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


def callout(text, size=9.5):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    shade(c, LIGHT)
    set_cell_text(c, text, size=size)
    c.width = Cm(USABLE)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def caption(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(8.5)
    r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(8)


def add_footer():
    section = doc.sections[0]
    p = section.footer.paragraphs[0]
    p.text = ""
    r = p.add_run("Vital Pakistan Trust  |  BRD – Consumption Analysis Dashboard  |  v1.0  |  Internal – Confidential        Page ")
    r.font.size = Pt(7.5)
    r.font.color.rgb = GREY
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    p._p.append(fld)


add_footer()

# ================================================================== COVER
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(70)
r = p.add_run("VITAL PAKISTAN TRUST")
r.font.size = Pt(11)
r.bold = True
r.font.color.rgb = GREY

p = doc.add_paragraph()
r = p.add_run("Business Requirements Document")
r.font.size = Pt(30)
r.bold = True
r.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
r = p.add_run("Consumption Analysis Dashboard")
r.font.size = Pt(20)
r.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(0)

p = doc.add_paragraph()
r = p.add_run("Consumption Analysis in View of Current Projects and Products")
r.font.size = Pt(12)
r.font.color.rgb = GREY
p.paragraph_format.space_after = Pt(14)
rule(p)

table(
    ["Item", "Detail"],
    [
        ["Document reference", "VPT-BRD-PRC-001"],
        ["Version", "1.0 – For Review and Approval"],
        ["Date", "29 October 2026"],
        ["Business owner", "Head of Procurement & Supply Chain"],
        ["Co-owners", "Director Programme; Manager Finance & Grants"],
        ["Prepared by", "Business Analysis – Digital Transformation"],
        ["Solution type", "Management dashboard (Power BI) over an Oracle-sourced reporting mart"],
        ["Primary data source", "Oracle ERP – Procurement (PO), Inventory (GRN, Issues, Stock), Finance (cost, budget)"],
        ["Companion documents", "Problem Statement v1.0; Sample Dataset (VPT_Consumption_Dataset.xlsx); Analysis & KPI Workbook"],
        ["Distribution", "Internal – Confidential. Not for circulation outside VPT without approval."],
    ],
    widths=[4.2, 13.0], font=9, align_center_cols=(),
)

doc.add_paragraph()
para("Approval of this document authorises the requirements baseline. Any subsequent change to scope, "
     "requirement priority or KPI definitions is subject to the change-control process in Section 22.",
     italic=True, size=8.5, color=GREY)

page_break()

# ============================================================ CONTENTS
doc.add_heading("Contents", level=1)
contents = [
    ("1.", "Executive Summary"),
    ("2.", "Introduction and Purpose"),
    ("3.", "Business Context and Objectives"),
    ("4.", "Problem Statement"),
    ("5.", "Scope"),
    ("6.", "Stakeholders, Users and RACI"),
    ("7.", "Current State (As-Is)"),
    ("8.", "Future State (To-Be)"),
    ("9.", "Business Requirements"),
    ("10.", "Functional Requirements"),
    ("11.", "Dashboard and Report Specification"),
    ("12.", "KPI Dictionary"),
    ("13.", "Data Requirements and Oracle Source Mapping"),
    ("14.", "Business Rules and Calculations"),
    ("15.", "Data Quality Requirements"),
    ("16.", "Non-Functional Requirements"),
    ("17.", "Assumptions, Constraints and Dependencies"),
    ("18.", "Risks and Mitigation"),
    ("19.", "Acceptance Criteria and UAT"),
    ("20.", "Delivery Approach, Phases and Timeline"),
    ("21.", "Benefits and Success Measurement"),
    ("22.", "Governance, Change Control and Support"),
    ("23.", "Glossary"),
    ("A.", "Appendix A – Reporting Data Model"),
    ("B.", "Appendix B – Indicative Oracle Table Mapping"),
    ("C.", "Appendix C – Dashboard Layout Wireframes"),
    ("D.", "Appendix D – Sample Dataset Description"),
    ("E.", "Appendix E – Requirement Traceability to Objectives"),
]
for num, title in contents:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run(f"{num}\t{title}")
    r.font.size = Pt(9.5)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.2))

# ============================================================ 1 EXEC SUMMARY
doc.add_heading("1. Executive Summary", level=1)
para(
    "Vital Pakistan Trust (VPT) runs ten concurrent health and nutrition projects funded by seven different donors and "
    "own sources. Each project consumes a continuous stream of medicines, consumables, laboratory supplies, nutrition "
    "commodities, equipment and general items. The transactions that record this consumption already exist in the Oracle "
    "ERP, but they are not presented in a form that allows management to compare what has been consumed against what has "
    "been purchased, by project and by product, over time."
)
para(
    "This Business Requirements Document (BRD) specifies a Consumption Analysis Dashboard that closes that gap. It "
    "defines what the business needs to see, how each measure is calculated, where the data comes from, and how the "
    "solution will be accepted. It does not prescribe the technical build in detail; the technical design, data model "
    "implementation and Power BI development follow in a separate solution design document."
)
doc.add_heading("1.1 What the business will get", level=2)
bullet("A single trusted view of consumption by project, product, category, warehouse and period, refreshed from Oracle without manual compilation.", bold_lead="One number. ")
bullet("Consumption value against procurement investment, per project and per product, with the ratio and variance made explicit — the core question this initiative exists to answer.", bold_lead="Consumption versus investment. ")
bullet("Month-on-month trend, seasonality, growth and abnormal movements, with drill-down to the individual issue, PO and batch.", bold_lead="Trend, not snapshots. ")
bullet("Days of cover, stock turnover, slow-moving and near-expiry stock, stock-outs and open commitments — the evidence on which procurement plans are built.", bold_lead="Forward-looking procurement. ")
bullet("Cost per beneficiary and cost per patient visit, linking product consumption to programme output.", bold_lead="Value for money. ")
bullet("Automatic surfacing of unattributed consumption, zero-value purchases, duplicate receipts, cancelled orders and price variance.", bold_lead="Controls and audit readiness. ")
doc.add_heading("1.2 Headline requirements", level=2)
table(
    ["#", "Requirement", "Priority"],
    [
        ["1", "Project-wise and product-wise consumption analysis for any selected period, with drill-down to transaction level.", "Must"],
        ["2", "Comparison of consumption value with procurement investment, showing ratio, variance and open commitments.", "Must"],
        ["3", "Trend analysis (month-on-month, moving average, seasonality) at total, project and product level.", "Must"],
        ["4", "Stock health analytics: on-hand, days of cover, turnover, slow-moving, near-expiry, expired and stock-out.", "Must"],
        ["5", "Budget versus consumption variance and burn rate by project and expense category.", "Must"],
        ["6", "Price-variance and supplier analytics to evidence value for money.", "Should"],
        ["7", "Cost per beneficiary and cost per visit using programme service-delivery counts.", "Should"],
        ["8", "Exception and data-quality panel with named owners for resolution.", "Must"],
        ["9", "Role-based security: project and donor users see only their own data.", "Must"],
        ["10", "Automated daily refresh from Oracle with visible data-freshness status.", "Must"],
        ["11", "What-if and scenario modelling on configurable parameters, with sensitivity and break-even analysis.", "Should"],
    ],
    widths=[1.0, 13.8, 2.4], font=8.5, align_center_cols=(0, 2),
)
doc.add_heading("1.3 Investment logic and expected benefit", level=2)
para(
    f"The materiality is not marginal. Across the {S['months']}-month sample period used to validate this specification, "
    f"VPT consumed approximately PKR {S['consumption_billable']/1e6:,.0f} million of products against approximately "
    f"PKR {S['purchase_value']/1e6:,.0f} million of procurement — a ratio of {S['consumption_billable']/S['purchase_value']:.2f} — "
    f"while product-related budget lines stood at approximately PKR {S['budget_product']/1e6:,.0f} million. Even a "
    "two-percentage-point improvement in the alignment of purchasing to consumption represents several million rupees a "
    "year in released working capital and avoided expiry. Beyond the direct financial benefit, the dashboard provides the "
    "evidence base required for donor reporting and audit on end-use of supplied commodities."
)
callout(
    "Success in one sentence: within two quarters of go-live, project managers and procurement plan from a shared, "
    "current view of consumption and investment — not from parallel spreadsheets compiled after the fact."
)

page_break()

# ============================================================ 2 INTRODUCTION
doc.add_heading("2. Introduction and Purpose", level=1)
doc.add_heading("2.1 Purpose of this document", level=2)
para(
    "This document records the agreed business requirements for the Consumption Analysis Dashboard. It is written for "
    "business approvers (Procurement, Programme and Finance leadership), for the technical team that will design and "
    "build the solution, and for the testers who will confirm acceptance. It is the contractual baseline for scope."
)
para("Specifically, the BRD aims to:")
bullet("state the business problem and the measurable outcome expected, so that the initiative is judged on business results rather than on delivery of features;")
bullet("specify the functional and non-functional requirements with unique identifiers, priorities and acceptance criteria;")
bullet("define every KPI precisely, including formula, grain, source and owner, so that a number has one meaning across the organisation;")
bullet("map each requirement to Oracle source data and state the business rules applied in transformation;")
bullet("define data-quality expectations and the treatment of known defects;")
bullet("set out the delivery approach, governance, risks and acceptance process.")
doc.add_heading("2.2 Intended audience and how to use it", level=2)
table(
    ["Audience", "Sections of primary interest"],
    [
        ["Executive management / approvers", "1, 3, 4, 5, 21"],
        ["Head of Procurement and supply chain team", "7, 8, 9, 10, 12, 14"],
        ["Programme and project managers", "5, 8, 10, 11, 12"],
        ["Finance and grants", "9, 12, 13, 14"],
        ["IT / ERP and data team", "13, 15, 16, Appendix A and B"],
        ["Developers and BI builders", "10, 11, 12, 13, 14, 16, Appendices A–C"],
        ["Testers and internal audit", "15, 19"],
    ],
    widths=[6.6, 10.6], font=8.5,
)
doc.add_heading("2.3 Definitions used in this document", level=2)
para(
    "“Consumption” means the value and quantity of items issued from a VPT store against a project, as recorded in "
    "Oracle inventory issues. “Investment” means the committed and received value of products purchased for that "
    "project, as recorded in purchase orders and goods receipt notes. “Product” means any item in the item master. "
    "“Project” means a funded programme with a unique project code. “Period” means a calendar month unless otherwise "
    "stated; fiscal years run July to June."
)

# ============================================================ 3 CONTEXT
doc.add_heading("3. Business Context and Objectives", level=1)
doc.add_heading("3.1 Organisational context", level=2)
para(
    "VPT delivers services through static facilities and mobile units across Karachi and interior Sindh. Funding is "
    "project-based and time-bound; each grant carries its own budget, reporting obligations and audit requirements. "
    "Products are procured centrally or at project level and stored in one central medical store and seven project or "
    "field stores. Consumption is recorded when items are issued to a service point or department and charged to a "
    "project code."
)
para(
    "Three structural characteristics shape the requirement: (a) many projects running simultaneously with overlapping "
    "product ranges, (b) product costs that move with inflation and exchange rates, making period comparisons "
    "sensitive to price as well as quantity, and (c) donor accountability that requires traceability from purchase to "
    "end use. Any solution must handle all three."
)
doc.add_heading("3.2 Business objectives", level=2)
table(
    ["Ref", "Business objective", "Measure of achievement"],
    [
        ["BO-1", "Make project and product consumption visible on a current basis", "Dashboard in routine use by all project managers; consumption available within 3 working days of month close"],
        ["BO-2", "Align procurement with actual consumption and lead times", "Reduction in over-purchase and expiry value; documented consumption-based procurement plans"],
        ["BO-3", "Evidence value for money to donors and management", "Cost per beneficiary and per service available for every active project"],
        ["BO-4", "Reduce manual reporting effort and conflicting figures", "At least 70% reduction in time spent compiling consumption reports; one agreed figure per KPI"],
        ["BO-5", "Strengthen control, traceability and audit readiness", "Every consumption value traceable from PO to batch to issue; exceptions surfaced with owners"],
        ["BO-6", "Improve budget control and forecasting", "Budget-versus-consumption variance visible monthly; burn rate and forecast at project level"],
        ["BO-7", "Build sustainable internal analytics capability", "Solution maintained by VPT staff; documented model, definitions and support process"],
    ],
    widths=[1.3, 6.8, 9.1], font=8.5, align_center_cols=(0,),
)
doc.add_heading("3.3 Alignment with organisational strategy", level=2)
para(
    "The dashboard supports VPT's digital-transformation agenda by converting existing ERP data into decision support "
    "without buying new transactional systems, and by establishing reusable data definitions, security patterns and "
    "governance that later initiatives (procurement planning, asset management, grant profitability) can build on."
)

# ============================================================ 4 PROBLEM
doc.add_heading("4. Problem Statement", level=1)
callout(
    "Vital Pakistan Trust cannot currently determine, on a timely and reliable basis, how much of each product has been "
    "consumed by which project, at what cost, and whether that consumption represents value for the investment made in "
    "purchasing it. Because the underlying transactions sit in different Oracle modules and rely on inconsistently "
    "applied project codes, the analysis is manual, delayed and retrospective. Procurement therefore runs on estimated "
    "demand rather than evidence, and over-consumption, under-utilisation, price escalation and stock-out risk are "
    "identified too late to act upon."
)
doc.add_heading("4.1 Problem decomposition", level=2)
table(
    ["Ref", "Problem component", "Business consequence"],
    [
        ["P-1", "No consolidated project × product × period view; data assembled manually from multiple Oracle extracts", "Analysis is slow, non-repeatable and dependent on individuals"],
        ["P-2", "Consumption and procurement reported separately, never reconciled", "Investment cannot be judged against use; over-buying is invisible"],
        ["P-3", "Inconsistent project coding at transaction level (some issues charged to generic cost centres)", "Attribution is unreliable; project managers distrust the figures"],
        ["P-4", "No trend analysis, only static period totals", "Seasonality, growth and abnormal movements are missed"],
        ["P-5", "No stock-health or expiry analytics", "Blocked capital, wastage and stock-outs discovered late"],
        ["P-6", "No linkage between consumption and service delivery output", "Cost per beneficiary cannot be evidenced"],
        ["P-7", "Manual reconciliation across Programme, Procurement and Finance", "Conflicting numbers, duplicated effort, weak audit trail"],
    ],
    widths=[1.4, 8.4, 7.4], font=8.5, align_center_cols=(0,),
)

# ============================================================ 5 SCOPE
doc.add_heading("5. Scope", level=1)
doc.add_heading("5.1 In scope", level=2)
bullet("Consumption of all items in the item master, issued from any VPT warehouse to any project, department or service point.")
bullet("Procurement transactions: requisitions, purchase orders, goods receipts, batches, prices and suppliers.")
bullet("Budget data by project, fiscal year and expense category, for comparison with consumption.")
bullet("Service-delivery counts by project and month, for cost-per-beneficiary and cost-per-visit measures.")
bullet("Master data: donors, projects, items, suppliers, warehouses, calendar.")
bullet("All projects active at any point in the reporting window, including closed projects for the periods during which they were active.")
bullet("Historical load of at least 24 months before go-live, and a monthly historical refresh thereafter.")
bullet("Analysis views, exception reporting, scheduled distribution and export.")
doc.add_heading("5.2 Out of scope (this release)", level=2)
table(
    ["Excluded item", "Rationale / planned treatment"],
    [
        ["Re-design or change of Oracle transactional processes", "Master-data and coding improvements will be raised separately as business-as-usual actions"],
        ["Full grant accounting, donor invoicing and profitability reporting", "Separate finance initiative; this dashboard uses budget data only for comparison"],
        ["Fixed assets, depreciation and equipment maintenance", "Equipment is reported for purchase and issue value only; asset register is out of scope"],
        ["HR, payroll, fleet and travel expenditure", "Out of scope; only product consumption and related budget lines are analysed"],
        ["Physical stock verification, bin-card digitisation and warehouse automation", "Operational initiative, independent of the dashboard"],
        ["Predictive forecasting using machine learning", "Deferred to Phase 2; Phase 1 provides trend-based indicative projection only"],
        ["Mobile application", "Responsive browser view only"],
        ["Real-time (intra-day) reporting", "Daily refresh is sufficient for the business need"],
    ],
    widths=[7.4, 9.8], font=8.5,
)
doc.add_heading("5.3 Phasing", level=2)
table(
    ["Phase", "Content", "Indicative duration"],
    [
        ["Phase 0 – Foundation", "Requirement sign-off, data profiling, KPI definitions, security model agreement", "3 weeks"],
        ["Phase 1 – Core", "Consumption and investment analysis, trend, project and product views, exceptions, budget variance", "8 weeks"],
        ["Phase 2 – Extended", "Stock health, expiry, supplier and price analytics, cost per beneficiary, subscriptions", "5 weeks"],
        ["Phase 3 – Optimise", "Adoption support, forecast view, refinement of thresholds and alerts", "4 weeks"],
    ],
    widths=[3.4, 11.4, 2.4], font=8.5, align_center_cols=(2,),
)
doc.add_heading("5.4 Reporting window and grain", level=2)
para(
    "The reporting window supports any period from July 2024 onward. The finest transaction grain is one inventory "
    "issue line for one item, one batch and one issue date. Self-service analysis is expected at the following grains: "
    "project × item × month; project × category × month; supplier × item × month; warehouse × item × month; and "
    "project × service point × month."
)

page_break()

# ============================================================ 6 STAKEHOLDERS
doc.add_heading("6. Stakeholders, Users and RACI", level=1)
doc.add_heading("6.1 User groups", level=2)
table(
    ["User group", "Indicative size", "Primary need", "Access level"],
    [
        ["Executive management", "5", "One-page progress and value-for-money view; risks and exceptions", "All projects, all data"],
        ["Head of Procurement", "1", "Consumption-based purchasing decisions, supplier and price performance", "All projects, all data"],
        ["Procurement officers", "4", "Requirement quantities, open commitments, lead times, price history", "All projects, transaction detail"],
        ["Programme managers", "3", "Project progress against budget, consumption trend, cost per beneficiary", "Assigned projects"],
        ["Project managers / coordinators", "10", "Own-project consumption versus purchase, stock adequacy, exceptions", "Own project only"],
        ["Finance and grants", "3", "Consumption reconciled to budget, variance commentary for donors", "All projects, financial measures"],
        ["Store / pharmacy in-charge", "8", "Stock on hand, days of cover, near-expiry, slow-moving", "Own warehouse"],
        ["Internal audit / compliance", "2", "Traceability from PO to issue; exception resolution status", "Read-only, all data"],
        ["IT / BI support", "2", "Model maintenance, refresh monitoring, security administration", "Administrative"],
    ],
    widths=[4.2, 2.4, 7.4, 3.2], font=8.5, align_center_cols=(1,),
)
doc.add_heading("6.2 RACI", level=2)
table(
    ["Activity", "Business owner", "Programme", "Finance", "IT / ERP", "BI developer"],
    [
        ["Approve BRD and KPI definitions", "A", "C", "C", "I", "I"],
        ["Provide Oracle access and data dictionary", "C", "I", "I", "R", "I"],
        ["Resolve master-data defects", "A", "C", "C", "R", "I"],
        ["Design and build dashboard", "C", "I", "I", "C", "R"],
        ["Validate figures against source", "A", "R", "R", "C", "C"],
        ["UAT execution and sign-off", "A", "R", "R", "I", "C"],
        ["Monthly commentary on variances", "C", "R", "R", "I", "I"],
        ["Maintain model after handover", "A", "I", "I", "C", "R"],
    ],
    widths=[6.6, 2.6, 2.0, 2.0, 2.0, 2.0], font=8.5, align_center_cols=(1, 2, 3, 4, 5),
)
caption("A = accountable, R = responsible, C = consulted, I = informed.")

# ============================================================ 7 AS-IS
doc.add_heading("7. Current State (As-Is)", level=1)
doc.add_heading("7.1 Process as it operates today", level=2)
table(
    ["Step", "Current activity", "System / tool", "Typical effort", "Pain point"],
    [
        ["1", "Field or facility raises a requirement based on visual stock check and past experience", "Paper / email", "Ongoing", "No consumption evidence behind the quantity"],
        ["2", "Procurement consolidates requirements and raises purchase orders", "Oracle PO", "2–4 days per cycle", "Quantities based on estimates, not trend"],
        ["3", "Goods received, batch and expiry recorded, stock updated", "Oracle inventory", "Same day", "Batch data captured but rarely analysed"],
        ["4", "Items issued to service points and charged to project codes", "Oracle inventory", "Same day", "Project code sometimes generic or missing"],
        ["5", "Month-end: staff export PO, GRN and issue data to spreadsheets", "Oracle + Excel", "3–5 days", "Manual, repetitive, version conflicts"],
        ["6", "Consumption and spend compiled into management reports", "Excel / PowerPoint", "2–4 days", "Lag of 7–15 days after month end"],
        ["7", "Variances investigated on request", "Email / meetings", "Ad hoc", "Reactive; often too late to change procurement"],
        ["8", "Donor and audit queries answered from ad hoc extracts", "Excel", "1–3 days per query", "No standing audit trail"],
    ],
    widths=[1.1, 6.3, 2.6, 2.6, 4.6], font=8, align_center_cols=(0, 2, 3),
)
doc.add_heading("7.2 Quantified pain points", level=2)
bullet("Total elapsed time from month end to a usable consumption report: 10–15 working days (to be confirmed during Phase 0).", bold_lead="Cycle time. ")
bullet("At least four parallel spreadsheets maintained by Procurement, Programme and Finance, with recurring reconciliation differences.", bold_lead="Duplication. ")
bullet("Between 2% and 4% of consumption value is charged to generic cost centres rather than a project, based on the sample data profiling exercise.", bold_lead="Attribution. ")
bullet("No standing measure of days of cover, slow-moving stock or near-expiry value; these are investigated only after a wastage event.", bold_lead="Stock visibility. ")
bullet("Cost per beneficiary has never been calculated consistently across projects; where quoted, methodologies differ.", bold_lead="Value for money. ")
doc.add_heading("7.3 Root causes", level=2)
table(
    ["Root cause", "Type", "Addressed by"],
    [
        ["No reporting layer between Oracle and management reporting", "Technology", "This solution (reporting mart + dashboard)"],
        ["KPI definitions not documented or owned", "Process / governance", "Section 12 KPI dictionary and approval"],
        ["Project code not mandatory or validated at issue", "Data / control", "Master-data action + exception reporting (FR-30)"],
        ["Master data (items, suppliers) not periodically cleansed", "Data", "Data-quality requirements, Section 15"],
        ["No routine consumption review forum", "Process", "Monthly review cadence, Section 21"],
        ["Analyst capacity absorbed by data assembly", "Capacity", "Automation of refresh and reporting"],
    ],
    widths=[7.0, 3.4, 6.8], font=8.5,
)

# ============================================================ 8 TO-BE
doc.add_heading("8. Future State (To-Be)", level=1)
doc.add_heading("8.1 Target process", level=2)
table(
    ["Step", "Future activity", "Owner", "Timing"],
    [
        ["1", "Oracle transactions posted as today; project code mandatory and validated at issue", "Store / procurement", "Daily"],
        ["2", "Automated extract loads the reporting mart; data-quality checks run and exceptions queued", "Automated / IT", "Nightly, before 07:00"],
        ["3", "Dashboard refreshed; data-freshness status visible on every page", "Automated", "Daily"],
        ["4", "Project managers review consumption versus purchase, trend and stock health for their project", "Project managers", "Weekly and at month end"],
        ["5", "Procurement reviews requirement quantities, days of cover, open commitments and price variance to build the purchase plan", "Procurement", "Monthly, by working day 3"],
        ["6", "Finance reviews budget versus consumption variance and records commentary against material variances", "Finance", "Monthly, by working day 5"],
        ["7", "Monthly consumption review meeting uses the dashboard as the single source", "Head of Procurement", "Monthly, working day 6"],
        ["8", "Exceptions assigned, resolved and evidenced; resolution tracked on the dashboard", "Named owners", "Ongoing, aged > 30 days escalated"],
    ],
    widths=[1.1, 8.4, 3.5, 4.2], font=8, align_center_cols=(0,),
)
doc.add_heading("8.2 User journeys", level=2)
doc.add_heading("8.2.1 Project manager – monthly review", level=3)
para("Open the Project view for their project and the selected month. Compare consumption value with procurement "
     "investment for the period and cumulative to date. Check the trend line against project activity. Review the stock "
     "health panel for items with less than one month of cover. Drill into any item with a consumption jump above the "
     "configured threshold, inspect the underlying issues and batches, and add a commentary note. Export the page to PDF "
     "for the project file.")
doc.add_heading("8.2.2 Procurement officer – building the purchase plan", level=3)
para("Open the Procurement Planning view filtered to items below the cover threshold. For each item, review the last "
     "twelve months of consumption, the current on-hand quantity, open commitments and supplier lead time. The "
     "dashboard proposes an indicative requirement quantity for the next period. The officer adjusts for known programme "
     "changes and records the decision, which becomes the documented basis of the requisition.")
doc.add_heading("8.2.3 Finance and grants – variance review", level=3)
para("Open the Budget view, filter by project and fiscal year, and sort by absolute variance. For each material variance, "
     "review the consumption detail that explains it, then record commentary. The export forms the basis of the donor "
     "financial report narrative.")
doc.add_heading("8.2.4 Executive – monthly oversight", level=3)
para("Open the Executive page: total consumption and investment, ratio, year-on-year movement, top projects and products, "
     "exceptions requiring decision. No drilling is required unless a metric is flagged red; the page answers “are our "
     "projects consuming value in line with plan, and where is the risk?”")
doc.add_heading("8.3 Benefits of the future state", level=2)
table(
    ["Benefit", "Mechanism", "Indicative value"],
    [
        ["Faster reporting", "Automated refresh and pre-built views", "From 10–15 days to same-day access after close"],
        ["Reduced manual effort", "No spreadsheet compilation", "70%+ reduction in analyst hours"],
        ["Better purchasing", "Consumption trend and days of cover drive requisitions", "Lower expiry write-off and emergency purchases"],
        ["Improved value for money", "Cost per beneficiary visible", "Evidence for donor negotiation and programme design"],
        ["Stronger control", "Exception panel and traceability", "Audit queries answered in minutes"],
        ["Single version of truth", "One KPI dictionary and one model", "Ends reconciliation disputes"],
    ],
    widths=[3.8, 6.0, 7.4], font=8.5,
)

page_break()

# ============================================================ 9 BUSINESS REQS
doc.add_heading("9. Business Requirements", level=1)
para("Business requirements express what the organisation needs to be able to do. They are numbered for traceability; "
     "functional requirements in Section 10 deliver them.")
brs = [
    ("BR-01", "Consolidated consumption visibility", "Provide one view of consumption quantity and value by project, product, category, warehouse and period, without manual compilation."),
    ("BR-02", "Consumption versus investment", "Show procurement investment alongside consumption so that the relationship between what is bought and what is used is explicit at project and product level."),
    ("BR-03", "Trend and pattern analysis", "Reveal movement over time — growth, seasonality, abnormal fluctuation — at total, project, category and product level."),
    ("BR-04", "Evidence-based procurement planning", "Enable requirement quantities to be derived from consumption, stock cover and lead time rather than judgement."),
    ("BR-05", "Value-for-money demonstration", "Express consumption against service output (beneficiaries, visits) and against budget, to evidence efficiency."),
    ("BR-06", "Stock health and expiry control", "Identify slow-moving, overstocked, near-expiry and expired stock, and stock-out risk, while action is still possible."),
    ("BR-07", "Price and supplier performance", "Track unit-price movement between purchases and compare suppliers on price, lead time and reliability."),
    ("BR-08", "Control and auditability", "Ensure every reported value can be traced through purchase, receipt, batch and issue, and that exceptions are surfaced and owned."),
    ("BR-09", "Budget control", "Compare consumption with budget by project and category, and project the year-end position."),
    ("BR-10", "Appropriate access and confidentiality", "Restrict data by role, project and donor, in line with VPT information-security policy."),
    ("BR-11", "Low-effort operation", "Refresh automatically, require no manual data handling, and be maintainable by VPT staff after handover."),
    ("BR-12", "Consistent definitions", "Apply one documented definition for every KPI across all views, reports and exports."),
]
table(["Ref", "Business requirement", "Description"], [[a, b, c] for a, b, c in brs],
      widths=[1.5, 5.0, 10.7], font=8.5, align_center_cols=(0,))

# ============================================================ 10 FUNCTIONAL
doc.add_heading("10. Functional Requirements", level=1)
para("Priorities follow MoSCoW. Acceptance criteria are stated as the observable behaviour that will be demonstrated "
     "in UAT. “Must” requirements are conditions of acceptance of the release.")
frs = [
    # Reporting & analysis
    ("FR-01", "Consumption by project and period", "Display consumption quantity and value for the selected period for each project, with subtotals and grand total.", "Rows reconcile to the sum of issue lines for the period; totals match the validated control total within PKR 1.", "Must", "BR-01"),
    ("FR-02", "Consumption by product and category", "Display consumption quantity and value by item and by item category for the selected scope, sortable by either measure.", "Category subtotals equal item totals; sorting and filtering do not alter totals.", "Must", "BR-01"),
    ("FR-03", "Consumption versus investment view", "Present consumption value and procurement investment side by side for each project and item, with ratio and variance.", "Ratio equals consumption ÷ investment rounded to two decimals; variance equals investment − consumption.", "Must", "BR-02"),
    ("FR-04", "Cumulative versus period view", "Allow the user to switch between period values and cumulative-to-date values for the project.", "Cumulative equals the sum of all periods from project start to the selected period end.", "Must", "BR-02"),
    ("FR-05", "Month-on-month trend", "Chart consumption value and quantity by month for any scope and period range.", "Chart values equal the corresponding table values for every month displayed.", "Must", "BR-03"),
    ("FR-06", "Growth and moving average", "Calculate month-on-month growth percentage, and three- and twelve-month moving averages.", "Growth equals (current − prior) ÷ prior; blank where prior is zero; moving average uses the stated number of prior months per the KPI dictionary.", "Should", "BR-03"),
    ("FR-07", "Seasonality view", "Present average consumption by calendar month across the selected years to reveal seasonal pattern.", "Averaging excludes months where the project was inactive.", "Should", "BR-03"),
    ("FR-08", "Comparison of periods or projects", "Allow side-by-side comparison of two periods, or of up to five projects, on selected KPIs.", "Comparison uses identical filters and definitions for both sides.", "Should", "BR-03"),
    ("FR-09", "Top-N and Pareto analysis", "Rank products by consumption value or quantity, showing the top N items and cumulative percentage contribution.", "N is user-selectable (5/10/20/50); cumulative percentage reaches 100% at the last item of the full list.", "Must", "BR-01"),
    ("FR-10", "Drill-down to transaction", "Navigate from any aggregated figure to project, category, item, batch, GRN, PO and individual issue lines.", "Every intermediate level reconciles to its parent figure; the lowest level lists the source document numbers.", "Must", "BR-08"),
    ("FR-11", "Consumption per service point", "Show consumption value and quantity by service point or department within a project.", "The sum across service points equals the project total for the same period and filters.", "Should", "BR-01"),
    ("FR-12", "Project life-cycle view", "Show consumption and investment against the project's elapsed timeline, start and end dates, and remaining period.", "Timeline reflects the project master dates; closed projects show values only for active months.", "Must", "BR-02"),
    ("FR-13", "Open commitments", "Report the value and quantity of purchase orders not yet fully received, by project and item, in-transit or pending.", "Open commitment equals ordered quantity less received quantity, valued at PO unit price; cancelled orders excluded.", "Must", "BR-02"),
    ("FR-14", "Cost per beneficiary", "Calculate product consumption value per beneficiary served for each project and month, plus year-to-date.", "Consumption value divided by beneficiaries served for the same project and month; blank if the denominator is zero, with a note.", "Should", "BR-05"),
    ("FR-15", "Cost per patient visit", "Calculate product consumption value per patient visit for each project and month.", "As above using patient visits as denominator.", "Should", "BR-05"),
    ("FR-16", "Consumption per capita of service volume trend", "Chart cost per beneficiary over time to show efficiency movement independently of price inflation.", "Chart uses a constant-unit-cost option in addition to actual cost, so price effect can be separated from volume effect.", "Could", "BR-05"),
    ("FR-17", "Budget versus consumption", "Compare consumption with budget by project, fiscal year and expense category, with variance and utilisation percentage.", "Variance equals budget − consumption; utilisation equals consumption ÷ budget; categories without budget are flagged.", "Must", "BR-09"),
    ("FR-18", "Burn rate and projection", "Show average monthly consumption against available budget and project a simple straight-line year-end position.", "Projection equals year-to-date consumption plus average monthly consumption × remaining months.", "Should", "BR-09"),
    ("FR-19", "Stock on hand and days of cover", "Report closing quantity and value on hand per item and warehouse, with days of cover based on recent consumption.", "On hand equals receipts less issues (plus opening) to the period end; days of cover equals on hand ÷ average daily consumption over the selected lookback.", "Must", "BR-06"),
    ("FR-20", "Stock turnover and slow-moving items", "Identify items with no or low consumption in the last 3, 6 and 12 months, with quantity and value held.", "An item appears in the 6-month list only if it had zero issues in the last 6 months and positive stock on hand.", "Must", "BR-06"),
    ("FR-21", "Near-expiry and expired stock", "List batches expiring within 30, 60, 90 and 180 days, and already expired, with quantity and value at risk.", "A batch appears in exactly one bucket by days to expiry from the report date; expired stock is valued and listed separately.", "Must", "BR-06"),
    ("FR-22", "Stock-out and unmet demand", "Report items and project-months where demand could not be met, plus out-of-stock days where data allows.", "Identified from requisition-to-issue lag or a documented stock-out register; each instance has project, item, month and quantity.", "Should", "BR-06"),
    ("FR-23", "Wastage and write-off", "Record and report wastage, damage and expiry write-offs by project, item and value.", "Write-off entries reconcile to approved inventory adjustment documents.", "Could", "BR-06"),
    ("FR-24", "Price variance and unit price trend", "Show average purchase unit price per item per period and the variance against the prior purchase, standard cost and budget assumption.", "Zero-value and donated lines are excluded from average price and shown separately; method stated on the page.", "Should", "BR-07"),
    ("FR-25", "Supplier analysis", "Compare suppliers on purchase value, average price per item, lead time and order fulfilment (received in full and on time).", "Lead time equals GRN date less PO date; fulfilment percentage equals fully received orders ÷ total orders.", "Should", "BR-07"),
    ("FR-26", "Warehouse and location analysis", "Show consumption and stock by warehouse and district, including transfer movements between stores.", "Inter-store transfers are shown separately and excluded from project consumption value to avoid double counting.", "Should", "BR-01"),
    ("FR-27", "Equipment versus consumables", "Separate capital equipment purchases and issues from consumable consumption in all analyses.", "Equipment items are identified from the item master flag; separate totals and a distinct view are provided.", "Should", "BR-01"),
    ("FR-28", "Donor-wise reporting pack", "Provide a donor-filtered view comprising consumption, investment, budget utilisation and cost per beneficiary.", "Filtering by donor restricts all visuals and exports consistently.", "Should", "BR-05"),
    ("FR-29", "Exception panel", "Surface unattributed consumption, zero-value purchases, duplicate receipts, cancelled orders, negative or unusual quantities, and missing master data.", "Each exception type shows count, value, age and named owner; items can be marked resolved with notes.", "Must", "BR-08"),
    ("FR-30", "Unattributed consumption handling", "Identify consumption charged to generic or invalid project codes and show it as a separate category with the originating project where identifiable.", "Unattributed value is never silently allocated; it appears in totals and in the exception panel.", "Must", "BR-08"),
    ("FR-31", "Threshold alerts and indicators", "Highlight values that breach configurable thresholds (for example consumption variance over 15%, cover below 15 days, expiry within 60 days).", "Thresholds are maintained by an administrator without code change; breached values are visually flagged.", "Should", "BR-08"),
    ("FR-32", "Commentary capture", "Allow an authorised user to record a text note against a project-period KPI or a variance.", "Notes are stored with user and timestamp, visible to authorised users, and retained in exports.", "Should", "BR-09"),
    ("FR-33", "Export and distribution", "Export any page to Excel and PDF with filters applied; schedule a page to be emailed to a distribution list.", "Export reflects the current filter state; scheduled delivery states the as-at date of the data.", "Must", "BR-11"),
    ("FR-34", "Data freshness indicator", "Display the date and time of the last successful data load and any load failure on every page.", "Indicator matches the load log; failure is shown in red with the last successful timestamp.", "Must", "BR-11"),
    ("FR-35", "Search and filter", "Provide filters for period, project, donor, category, item, warehouse, supplier, issue type and service point, with multi-select and search.", "Filters apply consistently across all visuals on the page; selections persist during the session.", "Must", "BR-01"),
    ("FR-36", "Drill-through to source document list", "List the underlying issue, GRN and PO records for any selected figure, exportable for audit.", "The exported list sums to the selected figure and includes document numbers, dates, quantities, costs and batches.", "Must", "BR-08"),
    ("FR-37", "Role-based security", "Restrict data by role: all-data, project-scoped, warehouse-scoped, donor-scoped and read-only audit.", "A project-scoped user cannot see other projects' data in any visual, export or drill-through.", "Must", "BR-10"),
    ("FR-38", "Audit trail of usage and changes", "Log access, exports and configuration changes (thresholds, commentary, ownership).", "Log retains at least 12 months and is accessible to IT and internal audit.", "Should", "BR-10"),
    ("FR-39", "Master-data exception register", "List items, projects and suppliers present in transactions but missing, inactive or invalid in the master data.", "Register is refreshed each load; each entry has a suggested owner action.", "Should", "BR-08"),
    ("FR-40", "Indicative requirement proposal", "For each item, propose an indicative purchase quantity for the next period based on recent consumption, cover target and lead time.", "Proposal equals (average monthly consumption × (cover target months + lead time in months)) − on hand − open commitment, floored at zero; the calculation basis is displayed.", "Should", "BR-04"),
    ("FR-41", "Consumption by issue type", "Separate project issues, department issues and transfers so that only project consumption is counted in project KPIs.", "Totals by issue type reconcile to the grand total of issues.", "Must", "BR-01"),
    ("FR-42", "Unit of measure consistency", "Normalise all quantities to the base unit of measure, with pack size available for reference.", "Pack-based ordering (for example a pack of 1,000 tablets) is converted to base units before aggregation; conversion is documented.", "Must", "BR-12"),
    ("FR-43", "Currency and inflation handling", "Report values in PKR; provide a constant-price view using a documented index for trend comparison.", "Actual and constant-price values are labelled distinctly and never mixed in one series.", "Could", "BR-03"),
    ("FR-44", "Project status and progress", "Show project start and end dates, elapsed and remaining duration, status and budget consumed to date.", "Progress percentage equals elapsed days ÷ total project days; status matches the project master.", "Must", "BR-02"),
    ("FR-45", "Saved views and personalisation", "Allow users to save filter combinations as personal or shared views.", "Saved views are private by default and shareable on request; they do not alter underlying definitions.", "Could", "BR-11"),
    ("FR-46", "Mobile and responsive access", "Provide a responsive layout for tablet and mobile for the executive and project pages.", "Key visuals remain legible and usable at 768 px width; no horizontal scrolling of KPI tiles.", "Could", "BR-11"),
    ("FR-47", "Documentation and definitions access", "Make KPI definitions, calculation rules and the data dictionary available from within the dashboard.", "A definitions page is reachable from every view and matches the approved KPI dictionary version.", "Should", "BR-12"),
    ("FR-49", "Configurable parameter register", "Hold every threshold and assumption in one governed register, displayed in the solution and changeable without code deployment.", "A change to a parameter is reflected in all affected statuses, commentary and scenarios at the next refresh, with no code change.", "Must", "BR-12"),
    ("FR-50", "What-if scenario modelling", "Allow a named scenario to be defined as a set of parameter overrides and evaluated against the baseline.", "Each scenario shows consumption, purchase requirement, funding gap, utilisation, stock-out items, expiry risk and working capital, with the delta against the baseline.", "Should", "BR-04"),
    ("FR-51", "Sensitivity analysis", "Rank the drivers by the size of their effect on a chosen output, moving one at a time between a low and a high value.", "A tornado view is produced for each output; the ranking is reproducible from the parameter register.", "Should", "BR-03"),
    ("FR-52", "Break-even and trade-off analysis", "Solve for the value of a driver at which a condition changes (for example funding gap becomes zero), and show the cost of a policy choice across a range.", "Break-even values are stated with the search range; the stock-cover trade-off curve shows purchase cost, expiry risk and working capital by cover level.", "Should", "BR-04"),
    ("FR-53", "Two-way scenario view", "Present an output across a grid of two drivers so that combinations can be read at a glance.", "Each cell equals a full model run for that combination; the current position is identifiable.", "Could", "BR-03"),
    ("FR-48", "Historical restatement control", "Preserve monthly snapshots so that previously reported figures can be reproduced after later data corrections.", "A closed period can be viewed as originally reported and as currently restated, with the difference shown.", "Should", "BR-08"),
]
table(["ID", "Requirement", "Description", "Acceptance criteria", "Priority", "BR"],
      [list(r) for r in frs],
      widths=[1.2, 3.1, 4.3, 5.0, 1.5, 1.3], font=7.6, align_center_cols=(0, 4, 5))

page_break()

# ============================================================ 11 DASHBOARD SPEC
doc.add_heading("11. Dashboard and Report Specification", level=1)
para("Five pages are specified. Each page states its purpose, audience and content. Wireframe sketches are in Appendix C. "
     "All pages carry the common filter bar, data-freshness indicator and export control.")

pages = [
    ("Page 1 – Executive Summary", "Executive management", [
        ("KPI tiles", "Total consumption value; total procurement investment; consumption-to-investment ratio; budget utilisation; unattributed consumption value; number of open exceptions"),
        ("Trend", "Monthly consumption value with 3-month moving average for the last 13 months; year-on-year comparison"),
        ("Top 10", "Projects by consumption value and products by consumption value, with share percentages"),
        ("Split", "Consumption by category (medicines, consumables, laboratory, nutrition, equipment, general)"),
        ("Attention list", "Projects with consumption variance beyond threshold, items below cover threshold, batches near expiry, open exceptions older than 30 days"),
        ("Navigation", "Click-through to project and product pages with the same filters applied"),
    ]),
    ("Page 2 – Project Analysis", "Programme and project managers, finance", [
        ("Header", "Project name, donor, district, status, start and end dates, elapsed percentage, remaining budget"),
        ("Consumption versus investment", "Period and cumulative values, ratio, variance, open commitments"),
        ("Trend", "Monthly consumption and purchase value by project; project life-cycle line with activity markers"),
        ("Composition", "Consumption by category, by item (Pareto), by service point, by warehouse"),
        ("Value for money", "Cost per beneficiary and cost per visit, trend and comparison with peer projects"),
        ("Stock health for project", "Items below cover threshold, near-expiry batches held for the project, stock-out occurrences"),
        ("Drill", "Item → batch → issue list → GRN → PO"),
    ]),
    ("Page 3 – Product Trend Analysis", "Procurement, pharmacists, stores", [
        ("Selector", "Item or category, with multi-select; defaults to top 20 items by value"),
        ("Trend", "Monthly quantity and value; 12-month moving average; seasonality profile by calendar month"),
        ("Price", "Average purchase unit price by period, variance against previous purchase and standard cost"),
        ("Movement", "Consumption by project for the item; concentration analysis; issue type split"),
        ("Stock", "On hand, days of cover, turnover, slow-moving flag, open commitments, supplier lead time"),
        ("Action", "Indicative requirement proposal with calculation basis (FR-40)"),
    ]),
    ("Page 4 – Budget and Value for Money", "Finance and grants", [
        ("Budget versus consumption", "Matrix of project × expense category with budget, consumption, variance and utilisation"),
        ("Burn rate", "Year-to-date consumption against straight-line expectation, with projected year-end position"),
        ("Efficiency", "Cost per beneficiary and per visit by project and month; trend; peer comparison"),
        ("Exceptions", "Categories overspent, categories unspent above threshold, projects approaching end date with material unspent budget"),
        ("Commentary", "Entry point for variance notes by project and period"),
    ]),
    ("Page 5 – Exceptions and Data Quality", "Procurement leads, internal audit, IT", [
        ("Summary", "Counts and value by exception type; ageing; ownership"),
        ("Unattributed consumption", "By month, by original project where identifiable, by item"),
        ("Zero-value or missing prices", "Purchase lines and issues affected, with impact on reported value"),
        ("Duplicate and cancelled documents", "Duplicate GRN candidates and cancelled POs with value"),
        ("Master-data gaps", "Items, projects and suppliers requiring correction, with suggested action"),
        ("Resolution log", "Status, owner, due date and notes for each exception; escalation for items older than 30 days"),
    ]),
]
for title, audience, content in pages:
    doc.add_heading(title, level=2)
    para(f"Primary audience: {audience}", italic=True, size=9, color=GREY, space_after=4)
    table(["Area", "Content"], [[a, b] for a, b in content], widths=[4.4, 12.8], font=8.5)

doc.add_heading("11.1 Common elements on every page", level=2)
bullet("Filter bar: period range, project, donor, category, item, warehouse, supplier, issue type, service point; with a reset control.")
bullet("Data-freshness indicator with last successful load timestamp and source extract reference.")
bullet("Currency and value basis label (actual PKR, excluding tax) plus the option to switch to constant-price view where enabled.")
bullet("Export to Excel and PDF, and a link to the KPI definitions page.")
bullet("Consistent colour semantics: consumption in navy, investment in teal, budget in grey, favourable variance in green, adverse in red.")

doc.add_heading("11.2 Report outputs outside the dashboard", level=2)
table(
    ["Output", "Frequency", "Recipient", "Content"],
    [
        ["Monthly consumption summary (PDF)", "Monthly, WD3", "Executive management, programme managers", "Pages 1–2 as at month end"],
        ["Procurement requirement report (Excel)", "Monthly, WD3", "Procurement officers", "Item-wise consumption, cover, open commitments, proposal"],
        ["Budget variance pack (Excel)", "Monthly, WD5", "Finance and grants", "Budget versus consumption with commentary"],
        ["Donor consumption statement (PDF)", "Quarterly", "Donor focal persons via Finance", "Project consumption, investment and cost per beneficiary"],
        ["Exception log (Excel)", "Weekly", "Procurement leads, internal audit", "Open exceptions with owner, age and value"],
    ],
    widths=[4.6, 2.6, 4.6, 5.4], font=8.5, align_center_cols=(1,),
)

page_break()

# ============================================================ 12 KPI
doc.add_heading("12. KPI Dictionary", level=1)
para("Every KPI below is approved once and used identically in all views. Where a KPI has a variant (such as constant "
     "price), the variant is a separate row. Definitions are expressed in business language; the technical expression is "
     "in Section 14.")
kpis = [
    ("KPI-01", "Consumption value", "PKR", "Sum of issue line value for the selected scope and period.",
     "Consumption value = Σ (quantity issued × unit cost at issue).", "Issue line", "Project / item / month",
     "Monitor", "All pages"),
    ("KPI-02", "Consumption quantity", "Base UOM", "Sum of quantities issued in base units.",
     "Σ quantity issued, converted from pack to base UOM.", "Issue line", "Project / item / month", "Monitor", "1–3"),
    ("KPI-03", "Procurement investment", "PKR", "Value of products purchased (ordered) for the selected scope and period.",
     "Σ (quantity ordered × PO unit price), excluding cancelled orders.", "PO line", "Project / item / month",
     "Monitor", "1–4"),
    ("KPI-04", "Received value", "PKR", "Value of goods actually received, used to compare with investment committed.",
     "Σ (quantity received × unit cost at receipt).", "GRN line", "Project / item / month", "Monitor", "2, 3"),
    ("KPI-05", "Consumption-to-investment ratio", "Ratio", "Extent to which purchased value has been consumed.",
     "Consumption value ÷ Procurement investment for the same scope and period.", "Project / item",
     "Project / item / period", "0.80 – 1.05", "1–3"),
    ("KPI-06", "Investment variance", "PKR", "Absolute difference between purchased and consumed value.",
     "Procurement investment − consumption value.", "Project / item", "Project / item / period",
     "Investigate > ±15%", "2, 3"),
    ("KPI-07", "Open commitment", "PKR and quantity", "Purchase orders not yet fully received.",
     "Σ ((quantity ordered − quantity received) × PO unit price), excluding cancelled orders.", "PO line",
     "Project / item", "Report", "1, 2"),
    ("KPI-08", "Month-on-month growth", "%", "Rate of change in consumption value against the prior month.",
     "(Current month − prior month) ÷ prior month × 100; blank if prior month is zero.", "Project / item",
     "Month", "Investigate > ±20%", "1–3"),
    ("KPI-09", "Three-month moving average", "PKR", "Smoothed consumption value to reveal underlying trend.",
     "Average of consumption value for the current and two preceding months where the scope was active.",
     "Project / item", "Month", "Monitor", "1–3"),
    ("KPI-10", "Cumulative consumption", "PKR", "Consumption from project start to the selected period end.",
     "Σ consumption value from project start through the selected month.", "Project", "Project",
     "Monitor", "2"),
    ("KPI-11", "Budget utilisation", "%", "Share of the budget consumed.",
     "Consumption value ÷ budget for the same project, fiscal year and category × 100.", "Project / category",
     "Project / fiscal year", "60–85% in line with elapsed time", "4"),
    ("KPI-12", "Burn rate", "PKR per month", "Average monthly consumption, used to project the year-end position.",
     "Year-to-date consumption value ÷ number of months elapsed in the fiscal year.", "Project", "Project / fiscal year",
     "Monitor", "4"),
    ("KPI-13", "Projected year-end consumption", "PKR", "Straight-line projection of full-year consumption.",
     "Year-to-date consumption + (burn rate × remaining months in the fiscal year).", "Project", "Project / fiscal year",
     "Within budget", "4"),
    ("KPI-14", "Cost per beneficiary", "PKR", "Product consumption value attributable to each beneficiary served.",
     "Consumption value for the project and month ÷ beneficiaries served for the same project and month.",
     "Project", "Project / month", "Monitor trend", "2, 4"),
    ("KPI-15", "Cost per patient visit", "PKR", "Product consumption value per recorded visit.",
     "Consumption value ÷ patient visits for the same project and month.", "Project", "Project / month",
     "Monitor trend", "2, 4"),
    ("KPI-16", "Stock on hand", "Quantity and PKR", "Closing stock at the report date.",
     "Opening stock + receipts − issues ± adjustments, valued at batch cost.", "Item / warehouse",
     "Item / warehouse", "Report", "3"),
    ("KPI-17", "Days of cover", "Days", "How long current stock will last at recent consumption rates.",
     "Stock on hand ÷ (average daily consumption over the selected lookback period, default 90 days).", "Item / project",
     "Item / warehouse", "≥ 30 days for critical items", "3"),
    ("KPI-18", "Stock turnover", "Count", "How many times stock is consumed and replaced in a period.",
     "Consumption quantity in the period ÷ average stock on hand in the period.", "Item", "Item / period",
     "Item-specific", "3"),
    ("KPI-19", "Slow-moving value", "PKR", "Value of stock with no consumption for the stated period.",
     "Value on hand of items with zero issues in the last 6 months (and separately 12 months).", "Item",
     "Item / warehouse", "Minimise", "3, 5"),
    ("KPI-20", "Near-expiry value", "PKR", "Value of stock expiring within the selected window.",
     "Value on hand of batches with expiry within 30/60/90/180 days from the report date.", "Batch",
     "Item / batch / warehouse", "Minimise", "3, 5"),
    ("KPI-21", "Expired value", "PKR", "Value of stock already past expiry and not written off.",
     "Value on hand of batches with expiry date before the report date.", "Batch", "Item / batch",
     "Zero", "3, 5"),
    ("KPI-22", "Stock-out count", "Count", "Number of item-project-months where demand could not be met.",
     "Count of instances where requisitioned quantity could not be issued due to nil stock.", "Item / project",
     "Item / project / month", "Minimise", "3"),
    ("KPI-23", "Average unit purchase price", "PKR", "Weighted average purchase price for an item in a period.",
     "Σ purchase line value ÷ Σ quantity purchased; excludes zero-value and donated lines.", "Item / supplier",
     "Item / month", "Monitor", "3"),
    ("KPI-24", "Price variance", "%", "Movement in unit price against the previous purchase or standard cost.",
     "(Current unit price − comparison price) ÷ comparison price × 100.", "Item", "Item / period",
     "Investigate > ±10%", "3"),
    ("KPI-25", "Supplier lead time", "Days", "Average elapsed time from PO to receipt.",
     "Average of (GRN date − PO date) for fully received orders.", "Supplier", "Supplier / period",
     "Within agreed terms", "3"),
    ("KPI-26", "Order fulfilment rate", "%", "Share of order lines received in full within the agreed lead time.",
     "Lines received in full and on time ÷ total order lines × 100.", "Supplier", "Supplier / period",
     "≥ 90%", "3"),
    ("KPI-27", "Unattributed consumption", "PKR and %", "Consumption not charged to a valid project.",
     "Consumption value on issues charged to generic or invalid codes ÷ total consumption value × 100.", "Issue line",
     "Month", "< 1%", "1, 5"),
    ("KPI-28", "Consumption concentration", "%", "Share of consumption value held by the top items.",
     "Value of top 10 items ÷ total consumption value × 100.", "Item", "Period", "Monitor", "1, 3"),
    ("KPI-29", "Inter-store transfer value", "PKR", "Value of stock moved between VPT stores, excluded from consumption.",
     "Σ issue value where issue type = inter-store transfer.", "Issue line", "Warehouse", "Report", "3, 5"),
    ("KPI-30", "Data freshness", "Hours", "Age of the data shown.",
     "Report date and time − last successful load timestamp.", "Load log", "System",
     "≤ 24 hours", "All pages"),
]
table(["ID", "KPI", "Unit", "Definition", "Formula", "Grain", "Reported by", "Target", "Pages"],
      [[a, b, c, d, e, f, g, h, i] for a, b, c, d, e, f, g, h, i in kpis],
      widths=[1.3, 3.1, 1.4, 3.6, 4.4, 1.4, 1.6, 1.7, 1.2], font=7.4,
      align_center_cols=(0, 2, 5, 7, 8))

page_break()

# ============================================================ 13 DATA
doc.add_heading("13. Data Requirements and Oracle Source Mapping", level=1)
doc.add_heading("13.1 Data domains", level=2)
table(
    ["Domain", "Required content", "Grain", "Source module", "Frequency"],
    [
        ["Project / award master", "Project code, name, donor, grant reference, sector, district, facility, start and end dates, status, approved budget, default warehouse", "Project", "Oracle Projects / custom master", "Daily"],
        ["Item master", "Item code, description, category, sub-category, base UOM, pack size, standard cost, reorder level, consumable flag, status", "Item", "Oracle Inventory", "Daily"],
        ["Supplier master", "Supplier code, name, city, tax number, payment terms, agreed lead time, status", "Supplier", "Oracle Purchasing", "Daily"],
        ["Warehouse master", "Warehouse code, name, district, type, in-charge", "Warehouse", "Oracle Inventory", "Daily"],
        ["Purchase orders", "PO number, date, requisition reference, project, supplier, item, quantity, unit price, currency, status, buyer, lead time", "PO line", "Oracle Purchasing", "Daily"],
        ["Goods receipts", "GRN number, date, PO reference, project, item, quantity received, warehouse, batch, manufacture and expiry dates, unit cost, QC status", "GRN line / batch", "Oracle Inventory", "Daily"],
        ["Inventory issues", "Issue number, date, project charged, project consumed, item, quantity, UOM, unit cost, total cost, warehouse, GRN reference, batch, expiry, issue type, service point, user", "Issue line", "Oracle Inventory", "Daily"],
        ["Budget", "Project, fiscal year, expense category, budget amount, currency", "Project / category / year", "Oracle GL / budget module", "Monthly"],
        ["Service delivery", "Project, month, beneficiaries served, patient visits, admissions", "Project / month", "Programme M&E (non-Oracle, supplied by Programme)", "Monthly"],
        ["Calendar", "Date, month, year, quarter, fiscal year, period key, close date", "Date", "Generated", "Static / annual"],
    ],
    widths=[3.3, 7.3, 2.6, 2.6, 1.4], font=7.9, align_center_cols=(2, 4),
)
doc.add_heading("13.2 Sourcing approach", level=2)
para(
    "Three options were considered for extracting data from Oracle. The recommended approach is a reporting mart with a "
    "scheduled extract, because it separates analytics load from the transactional system, supports historical "
    "snapshots and allows business rules to be applied once rather than in every report."
)
table(
    ["Option", "Description", "Advantages", "Risks / considerations", "Recommendation"],
    [
        ["Direct connection to Oracle", "Dashboard queries Oracle views live", "No duplication; always current", "Load on production ERP; complex Oracle joins in the BI layer; hard to apply business rules and history", "Not recommended"],
        ["Scheduled extract to a reporting mart", "Nightly extract into a staging schema, transformed into a star schema for the BI tool", "Protects ERP performance; business rules centralised; supports snapshots and restatement; testable", "Requires extract scheduling and mart maintenance", "Recommended"],
        ["Manual file extracts", "Exports from Oracle loaded by staff", "No technical dependency", "Reintroduces manual effort and error; not sustainable", "Interim only, for historical load if required"],
    ],
    widths=[3.0, 4.6, 4.0, 4.0, 1.6], font=7.9, align_center_cols=(4,),
)
doc.add_heading("13.3 Transformation rules", level=2)
table(
    ["Ref", "Rule", "Detail"],
    [
        ["TR-01", "Unit-of-measure normalisation", "All quantities converted to the item's base UOM using pack size; original UOM retained for reference."],
        ["TR-02", "Value calculation", "Line value = quantity × unit cost. Issue value uses the batch cost at the time of issue; PO value uses PO unit price."],
        ["TR-03", "Currency", "All values reported in PKR. Foreign currency purchases (if any) converted at the PO date rate; rate source documented."],
        ["TR-04", "Tax treatment", "Values reported exclusive of sales tax and other recoverable taxes; the basis is stated on the page."],
        ["TR-05", "Cancelled and rejected documents", "Cancelled POs excluded from investment; rejected GRN lines excluded from receipts and stock."],
        ["TR-06", "Duplicate removal", "GRN lines deduplicated on GRN number + item + batch + quantity + value; the retained record is the earliest."],
        ["TR-07", "Unattributed issues", "Issues charged to generic or invalid project codes retained in totals and flagged with the original consuming project where identifiable."],
        ["TR-08", "Transfers", "Inter-store transfers excluded from project consumption and reported separately to avoid double counting."],
        ["TR-09", "Zero or missing price", "Issues with zero cost valued at the standard cost and flagged; PO lines with zero value excluded from price averages and shown as donated or in-kind."],
        ["TR-10", "Closed projects", "Closed projects retain transactions for active months; post-close charges flagged as late charges."],
        ["TR-11", "Period assignment", "Transactions assigned to the month of the document date; the fiscal year derives from July–June convention."],
        ["TR-12", "Snapshot preservation", "At each month-end close a snapshot is stored so that prior reports can be reproduced and restatements quantified."],
        ["TR-13", "Late-arriving transactions", "Transactions posted after the snapshot appear in the current view as restatement, with the value of the difference shown."],
        ["TR-14", "Service-delivery alignment", "Beneficiary and visit counts supplied by Programme are matched on project and month; missing months leave the ratio blank rather than zero."],
    ],
    widths=[1.6, 4.2, 11.4], font=8.2, align_center_cols=(0,),
)
doc.add_heading("13.4 Data volumes and performance assumptions", level=2)
table(
    ["Object", "Volume to expect (per year)", "Retention", "Notes"],
    [
        ["Purchase order lines", "6,000 – 8,000", "Live + 5 years", "Grows with programme expansion"],
        ["GRN lines", "6,000 – 8,000", "Live + 5 years", "Batch-level grain"],
        ["Inventory issue lines", "9,000 – 12,000", "Live + 5 years", "Principal fact table"],
        ["Item master", "100 – 300 active items", "Full history", "Cleansing required before go-live"],
        ["Reports served", "Up to 40 concurrent users", "n/a", "Page load target under 5 seconds"],
        ["Refresh window", "Nightly, 02:00 – 06:00 PKT", "n/a", "Must not overlap ERP batch"],
    ],
    widths=[4.2, 4.2, 3.0, 5.8], font=8.3, align_center_cols=(1, 2),
)

# ============================================================ 14 RULES
doc.add_heading("14. Business Rules and Calculation Framework", level=1)
table(
    ["Ref", "Rule", "Statement"],
    [
        ["BL-01", "Consumption recognition", "Consumption is recognised on the inventory issue date, not on requisition or receipt."],
        ["BL-02", "Investment recognition", "Investment is recognised at purchase order value on the PO date; received value is recognised separately at GRN date."],
        ["BL-03", "Project attribution", "Consumption is attributed to the project charged on the issue line. Where the charged project is invalid or generic, the item is reported as unattributed and the originally consuming project is retained for correction."],
        ["BL-04", "Period activity", "A project contributes to a period only where the period overlaps the project's start and end dates."],
        ["BL-05", "Ratio calculation", "Where investment in a period is zero, the consumption-to-investment ratio is blank rather than infinite; a note explains that consumption was funded from opening stock."],
        ["BL-06", "Cost of issue", "Where batch cost is available it is used; otherwise the standard cost for the item is applied and the line is flagged."],
        ["BL-07", "Stock balance", "Stock on hand = opening balance + receipts − issues ± approved adjustments. Negative balances are treated as data errors and reported."],
        ["BL-08", "Days of cover", "Uses a 90-day lookback by default; items with no consumption in the lookback are reported as unlimited cover and listed as slow-moving."],
        ["BL-09", "Expiry buckets", "Batches are bucketed by days between the report date and the expiry date; a batch appears in exactly one bucket."],
        ["BL-10", "Fiscal year", "Fiscal year runs 1 July to 30 June and is labelled FY(start)–(end), for example FY2025-26."],
        ["BL-11", "Rounding", "Values are held to two decimals but displayed rounded to the nearest rupee; percentages to one decimal. Totals are computed before rounding."],
        ["BL-12", "Threshold defaults", "Consumption variance 15%, price variance 10%, cover 30 days for critical items, expiry alert 60 days, exception ageing 30 days. Administrator-maintainable."],
        ["BL-13", "Materiality for commentary", "Commentary is mandatory for variance above PKR 250,000 or 10% of the category budget, whichever is lower."],
        ["BL-14", "Report date", "The report date is the later of the last closed month end and the last successful load date, and is displayed on every page."],
    ],
    widths=[1.6, 4.0, 11.6], font=8.3, align_center_cols=(0,),
)

doc.add_heading("14.4 Configurable parameters and what-if capability", level=2)
para(
    "The solution is specified to be configurable rather than hard-coded. Two groups of values are held in a single "
    "governed register: status thresholds (which decide when a metric needs commentary or action) and model assumptions "
    "(which drive every projection). The same register drives the dashboard statuses, the exception flags, the commentary "
    "triggers and the scenario engine, so a change of tolerance or assumption cannot leave the pack inconsistent."
)
table(["Group", "Parameters held", "Why it is configurable", "Owner"],
      [[g, ", ".join(m["name"] for m in CFG["model"] if m["group"] == g),
        "Changes with policy, market or donor terms without a system change", "Business owner"]
       for g in ["Demand", "Price", "Supply", "Policy", "Budget"]],
      widths=[2.4, 8.4, 4.4, 2.0], font=8.0)
para(
    f"At the time of this version the register holds {len(CFG['model'])} model assumptions and {len(CFG['thresholds'])} "
    f"thresholds. The scenario engine evaluates a baseline plus {len(CFG['scenarios']) - 1} defined scenarios, a "
    "one-at-a-time sensitivity ranking, break-even solves and a two-way demand-by-price grid. A new scenario is added by "
    "listing its parameter overrides in the register; no development work is required."
)
table(["Scenario", "Levers moved", "Business question it answers"],
      [[sc["name"], ", ".join(sc["levers"].keys()) if sc["levers"] else "reference case", sc["rationale"]]
       for sc in CFG["scenarios"]],
      widths=[5.0, 5.0, 7.2], font=7.8)
para(
    "Governance: thresholds and assumptions are business decisions, not technical settings. Every change is recorded with "
    "the previous value, the new value, the reason, the approver and the effective date, and applies from the next refresh "
    "without amending historical snapshots. The register is reviewed twice a year and after any structural change "
    "(new grant, new project, system change)."
)

# ============================================================ 15 DATA QUALITY
doc.add_heading("15. Data Quality Requirements", level=1)
para(
    "The reliability of the dashboard depends entirely on the quality of the source data. The following requirements "
    "are therefore stated as conditions of a trustworthy release, and each known defect in the sample data is "
    "deliberately retained so that the handling logic can be tested.")
table(
    ["Ref", "Quality dimension", "Requirement", "Tolerance", "Handling"],
    [
        ["DQ-01", "Completeness – project attribution", "Every issue line carries a valid project code", "Zero unattributed", "Unattributed value reported separately, never allocated; exception raised"],
        ["DQ-02", "Completeness – item master", "Every transacted item exists and is active in the master", "Zero unmatched", "Unmatched items listed in the master-data exception register"],
        ["DQ-03", "Completeness – pricing", "Every issue and PO line carries a cost", "Under 1% of lines", "Zero-cost issues valued at standard cost and flagged; donated stock excluded from price analysis"],
        ["DQ-04", "Uniqueness", "No duplicate receipts for the same document, item and batch", "Zero duplicates", "Deduplication at load; duplicates reported"],
        ["DQ-05", "Validity – dates", "GRN date not earlier than PO date; issue date not earlier than GRN date", "Under 0.5% breaches", "Flagged and excluded from lead-time averages"],
        ["DQ-06", "Validity – quantities", "No negative or zero quantities on issues and receipts", "Zero", "Rejected at load and reported"],
        ["DQ-07", "Consistency – budget", "Every active project has a budget line for each applicable category", "100% of active projects", "Missing budget shown as not budgeted rather than zero"],
        ["DQ-08", "Consistency – service delivery", "Beneficiary counts available for every active project-month", "95%+", "Ratios left blank where counts are missing; gaps listed"],
        ["DQ-09", "Timeliness", "Data available within 24 hours of the transaction posting", "95% of days", "Load-failure alert to IT; freshness indicator on pages"],
        ["DQ-10", "Referential integrity", "All dimension keys resolve", "100%", "Unresolved keys assigned to an 'unknown' member and counted"],
        ["DQ-11", "Accuracy", "Reported figures reconcile to Oracle control totals", "Difference under 0.1%", "Reconciliation report produced each load; variance investigated"],
        ["DQ-12", "Auditability", "Every reported value traceable to source documents", "100%", "Drill-through list with document numbers retained in exports"],
    ],
    widths=[1.4, 3.6, 5.6, 2.4, 4.2], font=7.9, align_center_cols=(0,),
)
doc.add_heading("15.1 Data-quality control total reconciliation", level=2)
para("Each load must reconcile the following control totals before the dashboard is refreshed. A failure blocks the "
     "refresh and raises an alert rather than publishing suspect figures.")
table(
    ["Control", "Source", "Comparison", "Action on breach"],
    [
        ["Total consumption value", "Oracle issue register", "Sum of issue values for the loaded period", "Block refresh; notify IT and Head of Procurement"],
        ["Total receipts value", "Oracle receipt register", "Sum of GRN values", "Block refresh; investigate missing or duplicate lines"],
        ["Stock balance", "Oracle inventory balance", "Opening + receipts − issues ± adjustments", "Report variance; do not overwrite existing snapshot"],
        ["Item count", "Item master", "Distinct items in transactions versus master", "Report unmatched items; continue with exception entry"],
        ["Project count", "Project master", "Distinct projects in transactions versus master", "Report unmatched projects; continue with exception entry"],
    ],
    widths=[3.6, 3.6, 5.6, 4.4], font=8.3,
)

page_break()

# ============================================================ 16 NFR
doc.add_heading("16. Non-Functional Requirements", level=1)
nfrs = [
    ("NFR-01", "Performance", "Any dashboard page renders in under 5 seconds for 90% of requests at 40 concurrent users; drill-through under 8 seconds.", "Load testing at go-live"),
    ("NFR-02", "Refresh", "Daily automated refresh completed before 07:00 PKT; monthly close snapshot stored and immutable.", "Refresh log evidence for 30 consecutive days"),
    ("NFR-03", "Availability", "Available 99% of working hours (07:00–19:00 PKT, Monday–Saturday), excluding approved maintenance.", "Availability report"),
    ("NFR-04", "Security – authentication", "Single sign-on using VPT directory accounts; multi-factor for remote access where policy requires.", "Security review sign-off"),
    ("NFR-05", "Security – authorisation", "Row-level security by project, donor and warehouse aligned to role; verified for every visual and export.", "Negative testing in UAT"),
    ("NFR-06", "Data protection", "No patient-level or personal data in the model; aggregation only. Data at rest encrypted; transport over TLS.", "Security review sign-off"),
    ("NFR-07", "Auditability", "Access, export and configuration changes logged and retained for 12 months.", "Log extract demonstrated"),
    ("NFR-08", "Usability", "A new project manager can locate their project's consumption-versus-investment figure without training beyond a one-page guide.", "Usability test with five users"),
    ("NFR-09", "Accessibility", "Colour contrast meets WCAG 2.1 AA; meaning is never conveyed by colour alone.", "Accessibility checklist"),
    ("NFR-10", "Compatibility", "Current versions of Microsoft Edge and Chrome; responsive down to 768 px width.", "Browser test matrix"),
    ("NFR-11", "Maintainability", "Documented data model, naming conventions and transformation logic; VPT staff able to add a KPI with guidance.", "Handover acceptance"),
    ("NFR-12", "Scalability", "Model supports doubling of transaction volume and addition of projects without redesign.", "Design review"),
    ("NFR-13", "Licensing", "Operates within existing Microsoft licensing; any additional licensing identified and approved before build.", "Procurement confirmation"),
    ("NFR-14", "Localisation", "English interface in this release; PKR currency with thousand separators; dates as dd-mmm-yyyy.", "UAT check"),
    ("NFR-15", "Backup and recovery", "Model and mart recoverable to the previous load within one working day.", "Recovery test"),
    ("NFR-16", "Portability", "Solution documented well enough to be rebuilt by another supplier.", "Handover pack accepted"),
    ("NFR-17", "Configurability", "Every threshold, assumption and scenario is held in a single governed configuration file and can be changed without modifying code.", "Demonstrated parameter change with no code edit"),
    ("NFR-18", "Reproducibility", "Any reported or projected figure can be reproduced from the configuration version and the data snapshot in force.", "Rebuild reproduces figures exactly"),
]
table(["Ref", "Area", "Requirement", "Verification"], [[a, b, c, d] for a, b, c, d in nfrs],
      widths=[1.5, 3.0, 9.0, 3.7], font=8.3, align_center_cols=(0,))

# ============================================================ 17 ASSUMPTIONS
doc.add_heading("17. Assumptions, Constraints and Dependencies", level=1)
doc.add_heading("17.1 Assumptions", level=2)
for a in [
    "Oracle captures all material procurement receipts and inventory issues for every in-scope project, with document dates and project codes available.",
    "Read-only access to Oracle, or an agreed extract mechanism, is granted before build starts.",
    "Project codes and item codes are populated at transaction level for the large majority of records; the residual will be handled through exception reporting rather than manual allocation.",
    "Programme M&E can supply beneficiary and visit counts monthly in a stable format.",
    "Budget data by project, fiscal year and category is available in Oracle or can be supplied monthly by Finance.",
    "Business owners are available for KPI definition workshops within the Phase 0 window.",
    "The dashboard will be hosted in the organisation's existing Microsoft environment with Power BI capacity available.",
    "Historical data quality improves the further back the extract goes; the agreed reporting window begins July 2024.",
]:
    bullet(" " + a, style="List Bullet")
doc.add_heading("17.2 Constraints", level=2)
for c in [
    "Data confidentiality: extracts must remain within approved infrastructure; no project data may be copied to personal devices or external AI tools.",
    "Licensing: additional Power BI capacity or premium features require separate approval and budget.",
    "ERP stability: the reporting solution must not add material load to the production Oracle database.",
    "Resource availability: the Oracle/IT team's capacity is shared with other initiatives; extract work must be scheduled early.",
    "Timeline: the core release must be available before the FY2026-27 first-quarter review.",
]:
    bullet(" " + c, style="List Bullet")
doc.add_heading("17.3 Dependencies", level=2)
table(
    ["Dependency", "Owner", "Needed by", "Impact if late"],
    [
        ["Oracle extract or read-only access", "IT / ERP team", "Phase 1 start", "Build cannot begin; deliverable slips one week per week of delay"],
        ["Approved KPI definitions", "Business owner", "Phase 0 close", "Rework in every visual and report"],
        ["Historical data load and validation", "IT + Procurement", "Phase 1 mid-point", "Trend and year-on-year analysis unavailable"],
        ["Beneficiary and visit counts", "Programme M&E", "Phase 2 start", "Cost-per-beneficiary measures deferred"],
        ["Budget data", "Finance", "Phase 1 mid-point", "Budget variance page deferred"],
        ["Security model approval", "IT / data protection", "Before UAT", "Go-live blocked"],
    ],
    widths=[5.0, 3.2, 2.8, 6.2], font=8.3,
)

# ============================================================ 18 RISKS
doc.add_heading("18. Risks and Mitigation", level=1)
table(
    ["Ref", "Risk", "Likelihood", "Impact", "Mitigation", "Owner"],
    [
        ["R-01", "Poor project coding means unattributed consumption remains high and figures are disputed", "High", "High", "Exception reporting with owners; master-data clean-up as a parallel action; make the project field mandatory at issue", "Head of Procurement"],
        ["R-02", "Oracle access delayed by IT scheduling", "Medium", "High", "Agree mechanism in week one; interim manual extract for history if necessary", "IT manager"],
        ["R-03", "KPI definitions not agreed, causing rework", "Medium", "Medium", "Sign-off gate at end of Phase 0; definitions embedded in the model", "Business owner"],
        ["R-04", "Low adoption after launch", "Medium", "High", "Embed in monthly review meeting; assign exception owners; champion in each department; short training", "Business owner"],
        ["R-05", "Data volumes degrade performance", "Low", "Medium", "Star schema with aggregation; query tests against three years of data before go-live", "BI developer"],
        ["R-06", "Scope creep into full financial reporting", "High", "Medium", "Phasing and change control; out-of-scope list in Section 5.2", "Project lead"],
        ["R-07", "Beneficiary data inconsistent across projects", "High", "Medium", "Publish methodology; blank rather than zero where missing; M&E validation", "Programme"],
        ["R-08", "Key-person dependency in BI build", "Medium", "Medium", "Documentation, handover sessions, code review and repository standards", "IT manager"],
        ["R-09", "Licensing shortfall discovered late", "Low", "High", "Confirm licensing in Phase 0 before build", "IT manager"],
        ["R-10", "Sensitive supplier pricing exposed too broadly", "Low", "Medium", "Restrict cost and supplier detail to authorised roles", "IT / data protection"],
    ],
    widths=[1.2, 5.0, 1.8, 1.4, 6.0, 1.8], font=7.9, align_center_cols=(0, 2, 3),
)

# ============================================================ 19 ACCEPTANCE
doc.add_heading("19. Acceptance Criteria and UAT", level=1)
doc.add_heading("19.1 Release acceptance criteria", level=2)
for i, a in enumerate([
    "All Must-priority functional requirements demonstrated and signed off.",
    "Every KPI reconciles to a control total derived independently from Oracle for three test months (including one month with a known data defect and one month-end close).",
    "Row-level security verified by negative testing for each role, including exports and drill-through.",
    "Performance targets met with three years of data and 40 simulated users.",
    "Data-quality checks and control-total reconciliation run successfully in three consecutive loads.",
    "KPI dictionary and data dictionary published in the dashboard and matching the approved version.",
    "One-page user guide for each persona published; two training sessions delivered.",
    "Support and ownership model agreed and documented.",
], start=1):
    bullet(" " + a, bold_lead=f"{i}. ")
doc.add_heading("19.2 UAT scenarios", level=2)
table(
    ["ID", "Scenario", "Steps", "Expected result"],
    [
        ["UT-01", "Project consumption for a period", "Select project, select month; view page 2", "Consumption value matches the sum of issue lines for that project and month"],
        ["UT-02", "Consumption versus investment", "Compare period values and ratio for a project", "Ratio and variance calculate per KPI-05 and KPI-06 and reconcile to the underlying lists"],
        ["UT-03", "Trend and growth", "View 13-month trend and growth percentage", "Growth equals the formula in KPI-08; blanks where prior month is zero"],
        ["UT-04", "Product drill-down", "Drill from top product to batch and issue list", "List sums to the selected product figure and shows document numbers"],
        ["UT-05", "Unattributed consumption", "Open the exception panel for a month with unattributed issues", "Value is reported separately, appears in totals, and is not allocated to any project"],
        ["UT-06", "Zero-value purchase", "Filter purchase analysis for donated stock", "Excluded from average price; shown separately as in-kind"],
        ["UT-07", "Duplicate receipts", "Run the duplicate GRN check", "Duplicate pair identified once, with value and recommended action"],
        ["UT-08", "Near-expiry report", "Set expiry window to 60 days", "Only batches expiring within 60 days appear; values tie to stock on hand"],
        ["UT-09", "Budget variance", "Select a project and fiscal year on page 4", "Budget, consumption, variance and utilisation reconcile; missing budget flagged"],
        ["UT-10", "Cost per beneficiary", "View a project-month with known beneficiary count", "Value equals consumption ÷ beneficiaries; blank where count missing"],
        ["UT-11", "Role-based security", "Log in as a project-scoped user", "No other project's data visible in visuals, filters, exports or drill-through"],
        ["UT-12", "Data quality gate", "Simulate a failed control-total reconciliation", "Refresh blocked, alert raised, previous figures retained and freshness indicator shows failure"],
        ["UT-13", "Export and schedule", "Export page 2 to Excel; schedule weekly email", "Export reflects current filters; scheduled email received with as-at date"],
        ["UT-14", "Definitions", "Open KPI definitions from page 3", "Definitions match the approved dictionary version"],
    ],
    widths=[1.4, 4.2, 5.2, 6.4], font=7.9, align_center_cols=(0,),
)

page_break()

# ============================================================ 20 DELIVERY
doc.add_heading("20. Delivery Approach, Phases and Timeline", level=1)
table(
    ["Phase", "Key activities", "Deliverables", "Duration", "Exit criteria"],
    [
        ["Phase 0 – Discover and define", "Requirements workshops; KPI definition sign-off; Oracle data profiling; security model; sample data validation", "Approved BRD; KPI dictionary; data profile report; security model note", "3 weeks", "BRD and KPI dictionary signed off"],
        ["Phase 1 – Core build", "Extract and staging; star schema; core pages 1, 2 and 5; budget variance; exception logic", "Working dashboard (pages 1, 2, 5); data model documentation; reconciliation report", "8 weeks", "Control totals reconcile for three months"],
        ["Phase 2 – Extended build", "Stock health, expiry, supplier and price analytics; page 3 and 4 completion; cost per beneficiary; subscriptions", "Completed pages 3 and 4; scheduled distribution; exception resolution log", "5 weeks", "All Must requirements demonstrated"],
        ["Phase 3 – UAT, training and go-live", "UAT execution; defect resolution; training; handover; go-live and hypercare", "UAT sign-off; user guides; handover pack; live solution", "4 weeks", "Acceptance criteria met and approved"],
        ["Phase 4 – Adoption and optimise", "Monthly review support; threshold tuning; benefit tracking; backlog for Phase 2 predictive features", "Benefit report; tuned thresholds; prioritised enhancement backlog", "4 weeks", "Dashboard in routine use by all personas"],
    ],
    widths=[3.2, 5.6, 4.4, 1.6, 2.4], font=7.9, align_center_cols=(3,),
)
para("Indicative end-to-end duration: 20 weeks from Phase 0 start, with the core release (pages 1, 2 and 5) available "
     "around week 11. Timelines assume dependencies in Section 17.3 are met.",
     italic=True, size=8.5, color=GREY)

doc.add_heading("20.1 Roles in delivery", level=2)
table(
    ["Role", "Responsibility"],
    [
        ["Business owner (Head of Procurement)", "Accountable for requirements, definitions, adoption and benefits"],
        ["Business analyst", "Elicitation, documentation, traceability, UAT coordination"],
        ["BI developer", "Data model, transformation logic, dashboard build"],
        ["Oracle/ERP specialist", "Extract design, source queries, table mapping"],
        ["Data quality reviewer", "Control totals, reconciliation, exception validation"],
        ["Programme M&E focal person", "Service-delivery data and methodology"],
        ["Finance focal person", "Budget data and variance validation"],
        ["Change and training lead", "User guides, training, adoption tracking"],
    ],
    widths=[5.4, 11.8], font=8.5,
)

# ============================================================ 21 BENEFITS
doc.add_heading("21. Benefits and Success Measurement", level=1)
table(
    ["Ref", "Benefit", "Metric", "Baseline (to confirm)", "Target", "Measured by"],
    [
        ["BN-01", "Faster reporting", "Days from month end to available consumption report", "10–15 days", "Same day after close; ≤ 3 days for full pack", "Finance"],
        ["BN-02", "Reduced manual effort", "Analyst hours per month spent compiling consumption data", "To be baselined in Phase 0", "70% reduction", "Head of Procurement"],
        ["BN-03", "Better purchasing", "Value of stock expiring or written off per year", "To be baselined", "30% reduction year on year", "Stores"],
        ["BN-04", "Improved stock availability", "Item-project-months out of stock", "To be baselined", "50% reduction", "Stores / Programme"],
        ["BN-05", "Value for money", "Cost per beneficiary, compared across projects", "Not currently measured", "Published monthly for all projects", "Finance / Programme"],
        ["BN-06", "Cleaner attribution", "Unattributed consumption as a percentage of total", "2–4%", "Under 1%", "Procurement"],
        ["BN-07", "Audit readiness", "Time to respond to an audit query on commodity end-use", "1–3 days", "Under 2 hours", "Internal audit"],
        ["BN-08", "Adoption", "Active users per month; projects using the dashboard in review meetings", "0", "All project managers; 100% of projects", "BI usage log / minutes"],
    ],
    widths=[1.3, 3.3, 4.6, 3.0, 2.8, 2.2], font=7.9, align_center_cols=(0,),
)
para("Benefits will be reviewed quarterly by the business owner, using the dashboard's own usage log and the metrics "
     "above. Any benefit not realised within two quarters triggers a review of process adoption rather than of the tool.")

# ============================================================ 22 GOVERNANCE
doc.add_heading("22. Governance, Change Control and Support", level=1)
doc.add_heading("22.1 Governance", level=2)
bullet("A steering group chaired by the Director Finance & Operations meets fortnightly during delivery and quarterly thereafter.")
bullet("The business owner is accountable for content and adoption; IT is accountable for platform, security and refresh.")
bullet("Requirements are baselined on approval of this BRD; changes follow the process below.")
doc.add_heading("22.2 Change control", level=2)
table(
    ["Change type", "Example", "Approval"],
    [
        ["Minor (no scope change)", "Label change, colour, threshold default", "Business owner"],
        ["Moderate", "New visual on an existing page, additional filter, new exception type", "Business owner + IT"],
        ["Major", "New data source, new page, change to a KPI definition after go-live, additional user group", "Steering group"],
        ["Emergency", "Fix to an incorrect figure in production", "Business owner, notified to steering group at next meeting"],
    ],
    widths=[3.6, 9.4, 4.2], font=8.5,
)
doc.add_heading("22.3 Support model after handover", level=2)
table(
    ["Level", "Scope", "Owner", "Target response"],
    [
        ["Level 1", "Access issues, how-to questions, refresh status enquiry", "IT service desk", "4 working hours"],
        ["Level 2", "Data discrepancies, failed loads, security exceptions", "IT / BI support", "1 working day"],
        ["Level 3", "Model changes, new requirements, performance tuning", "BI developer / supplier", "Next change cycle"],
    ],
    widths=[2.4, 7.6, 4.2, 3.0], font=8.5,
)
doc.add_heading("22.4 Documentation and knowledge transfer", level=2)
bullet("Data model, transformation rules, KPI dictionary, security design and refresh schedule held in a controlled repository.")
bullet("Two handover sessions with IT and one with business super-users, recorded and minuted.")
bullet("A named business super-user per department trained to answer first-line content questions.")

# ============================================================ 23 GLOSSARY
doc.add_heading("23. Glossary", level=1)
gloss = [
    ("Allocation", "Distribution of shared consumption across projects. Not performed in this release; unattributed value is reported instead."),
    ("Batch / lot", "A quantity of an item produced or received together, carrying a common expiry date."),
    ("Burn rate", "Average monthly consumption value used to project the full-year position."),
    ("Consumption", "Value and quantity of items issued from store and charged to a project, department or service point."),
    ("Consumption-to-investment ratio", "Consumption value divided by procurement investment for the same scope and period."),
    ("Days of cover", "Stock on hand divided by average daily consumption; indicates how long stock will last."),
    ("Dimension", "A business attribute used to slice data, such as project, item, donor, warehouse or period."),
    ("FEFO", "First expiry, first out — the issue method that releases the earliest-expiring batch first."),
    ("Fact table", "The table holding transactional measures at a defined grain, such as one issue line."),
    ("Fiscal year", "VPT's financial year, 1 July to 30 June."),
    ("Grain", "The level of detail at which a fact is recorded."),
    ("GRN", "Goods receipt note; the document recording receipt of purchased items into store."),
    ("Investment (procurement)", "Committed value of products purchased, recognised at purchase order value."),
    ("KPI", "Key performance indicator; a defined measure with an owner and a target."),
    ("Open commitment", "Value of purchase orders not yet fully received."),
    ("Pareto analysis", "Ranking of items by value to show that a minority of items accounts for the majority of consumption."),
    ("Price variance", "Movement in the unit price of an item between purchases or against standard cost."),
    ("Restatement", "Change to previously reported figures caused by later data corrections."),
    ("Slow-moving stock", "Stock with no or minimal consumption over the defined lookback period."),
    ("Snapshot", "A stored copy of figures as at a month-end close, used to reproduce earlier reports."),
    ("Star schema", "A data model with a central fact table joined to conformed dimension tables."),
    ("Stock-out", "A situation in which required stock is not available, so demand cannot be met."),
    ("UAT", "User acceptance testing; business validation that the solution meets the requirements."),
    ("Unattributed consumption", "Consumption charged to a generic or invalid project code rather than to a project."),
]
table(["Term", "Definition"], [list(g) for g in gloss], widths=[4.2, 13.0], font=8.3)

page_break()

# ============================================================ APPENDIX A
doc.add_heading("Appendix A – Reporting Data Model", level=1)
para("The reporting model follows a star schema. Dimensions are conformed so that every fact can be sliced by the same "
     "attributes, and the transactional facts share keys that allow drill-through from consumption to purchase.")
doc.add_heading("A.1 Dimensions", level=2)
table(
    ["Dimension", "Key", "Principal attributes"],
    [
        ["Dim_Donor", "DonorKey", "Donor name, type, grant reference, currency, reporting frequency"],
        ["Dim_Project", "ProjectKey", "Project code, name, donor, sector, district, facility, start and end dates, status, approved budget, default warehouse"],
        ["Dim_Item", "ItemKey", "Item code, description, category, sub-category, base UOM, pack size, standard cost, consumable flag, reorder level"],
        ["Dim_Supplier", "SupplierKey", "Supplier name, city, tax number, payment terms, agreed lead time, status"],
        ["Dim_Warehouse", "WarehouseKey", "Warehouse code, name, district, type, in-charge"],
        ["Dim_Date", "DateKey", "Date, month, quarter, fiscal year, period key, close date"],
        ["Dim_ServicePoint", "ServicePointKey", "Service point or department within a project"],
        ["Dim_ExceptionType", "ExceptionTypeKey", "Exception category and default owner"],
    ],
    widths=[3.6, 2.8, 10.8], font=8.3,
)
doc.add_heading("A.2 Facts", level=2)
table(
    ["Fact", "Grain", "Measures"],
    [
        ["Fact_PurchaseOrder", "One PO line", "Quantity ordered, unit price, line value, lead time days"],
        ["Fact_Receipts", "One GRN line for one batch", "Quantity received, unit cost, received value"],
        ["Fact_Consumption", "One issue line for one item and batch", "Quantity issued, unit cost, issue value"],
        ["Fact_StockBalance", "Item × warehouse × period end", "Opening quantity, receipts, issues, adjustments, closing quantity and value"],
        ["Fact_Budget", "Project × fiscal year × expense category", "Budget amount"],
        ["Fact_ServiceDelivery", "Project × month", "Beneficiaries served, patient visits, admissions"],
        ["Fact_Exception", "One exception instance", "Value at risk, age in days, status"],
    ],
    widths=[3.6, 5.2, 8.4], font=8.3,
)
doc.add_heading("A.3 Relationships", level=2)
bullet("Fact_Consumption.GRNKey → Fact_Receipts.GRNKey → Fact_PurchaseOrder.POKey, giving full traceability from issue to purchase.")
bullet("Fact_Consumption.ProjectKey may point to an 'Unattributed' member; Fact_Consumption.OriginalProjectKey retains the consuming project.")
bullet("All facts join to Dim_Date on DateKey, Dim_Item on ItemKey and Dim_Project on ProjectKey.")
bullet("Fact_StockBalance is a semi-additive measure: balances are additive across items and warehouses, but not across time.")
bullet("Fact_ServiceDelivery joins to Dim_Project and Dim_Date at month grain, used only as a denominator for value-for-money measures.")

# ============================================================ APPENDIX B
doc.add_heading("Appendix B – Indicative Oracle Table Mapping", level=1)
para("Table and column names are indicative and must be confirmed against the VPT Oracle instance during Phase 0. "
     "The purpose of this appendix is to demonstrate that every required attribute has an identified source.")
table(
    ["Reporting attribute", "Indicative Oracle source", "Notes"],
    [
        ["Purchase order number and date", "PO_HEADERS_ALL (SEGMENT1, APPROVED_DATE)", "Use approved date for recognition"],
        ["PO line quantity and price", "PO_LINES_ALL (QUANTITY, UNIT_PRICE, ITEM_ID)", "Exclude cancelled lines via PO_LINE_TYPES / CANCEL_FLAG"],
        ["PO status", "PO_HEADERS_ALL.AUTHORIZATION_STATUS, PO_LINES_ALL.CANCEL_FLAG", "Derive Received / Partially Received from receipt quantities"],
        ["Supplier", "PO_HEADERS_ALL.VENDOR_ID → AP_SUPPLIERS", "Add agreed lead time from a local reference table if absent"],
        ["Receipt number and date", "RCV_SHIPMENT_HEADERS (RECEIPT_NUM, SHIPMENT_DATE)", "One receipt may contain several lines"],
        ["Received quantity and batch", "RCV_SHIPMENT_LINES (QUANTITY_SHIPPED, LOT_NUM)", "Batch attributes from inventory lot table"],
        ["Batch expiry", "INV_LOT_ATTRIBUTES / MTL_LOT_NUMBERS", "Expiry stored as an attribute; confirm attribute name"],
        ["Issue transaction", "MTL_TRANSACTION_ACCOUNTS or MTL_MATERIAL_TRANSACTIONS", "Use issue transaction type; join to cost for value"],
        ["Issue value", "MTL_TRANSACTION_ACCOUNTS (TRANSACTION_COST)", "Batch cost at issue; standard cost fallback"],
        ["Project charged", "MTL_MATERIAL_TRANSACTIONS.PROJECT_ID / transaction account segments", "Confirm whether project is derived from distribution account or project field"],
        ["Item master", "MTL_SYSTEM_ITEMS_VL (SEGMENT1, DESCRIPTION, PRIMARY_UOM_CODE)", "Category from item category assignment"],
        ["Unit of measure and pack", "MTL_UOM_CONVERSIONS", "Pack size from item-level conversion"],
        ["Warehouse", "MTL_PARAMETERS / ORG_ORGANIZATION_DEFINITIONS", "Inventory organisation code"],
        ["Service point", "MTL_MATERIAL_TRANSACTIONS.SUBINVENTORY or a local reference table", "May require a mapping table if not held in Oracle"],
        ["Budget", "GL_BUDGET_BALANCES / GL_BUDGET_ASSIGNMENTS or a Finance-maintained table", "Map GL account to expense category"],
        ["Project master", "PA_PROJECTS_ALL (PROJECT_NUMBER, PROJECT_NAME, START_DATE, COMPLETION_DATE)", "Plus local donor and sector attributes"],
        ["Calendar", "GL_PERIODS / a generated date dimension", "Fiscal year from July–June convention"],
    ],
    widths=[4.2, 7.0, 6.0], font=8.0,
)

# ============================================================ APPENDIX C
doc.add_heading("Appendix C – Dashboard Layout Wireframes", level=1)
para("Wireframes are described as blocks in reading order. They are illustrative of hierarchy and priority rather than "
     "of final visual design.")

def wire(title, blocks):
    doc.add_heading(title, level=2)
    for label, content in blocks:
        t = doc.add_table(rows=1, cols=1)
        t.style = "Table Grid"
        c = t.rows[0].cells[0]
        shade(c, LIGHT if "Filter" in label or "KPI" in label else LIGHTER)
        c.text = ""
        p = c.paragraphs[0]
        r = p.add_run(f"{label}  ")
        r.bold = True
        r.font.size = Pt(8.2)
        r2 = p.add_run(content)
        r2.font.size = Pt(8.2)
        p.paragraph_format.space_after = Pt(1)
        c.width = Cm(USABLE)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)

doc.add_heading("C.1 Page 1 – Executive Summary", level=2)
table(["Block", "Content"], [
    ["Filter / header", "Period range · Donor · Project · Category | Data as at: 30-Sep-2026 06:12 | Export"],
    ["KPI row", "Consumption value · Investment · Ratio · Budget utilisation · Unattributed · Open exceptions"],
    ["Trend panel", "Monthly consumption value, 13 months, with 3-month moving average line and prior-year comparison"],
    ["Ranking panel", "Top 10 projects by consumption value (bar, with share %)"],
    ["Ranking panel", "Top 10 products by consumption value (bar, with share %)"],
    ["Split", "Consumption by category (stacked bar by month)"],
    ["Attention list", "Threshold breaches: variance > 15%, cover < 30 days, expiry < 60 days, exceptions > 30 days"],
], widths=[3.4, 13.8], font=8.2)
doc.add_heading("C.2 Page 2 – Project Analysis", level=2)
table(["Block", "Content"], [
    ["Project header", "Name · donor · district · status · start–end · elapsed % · remaining budget"],
    ["KPI row", "Period consumption · Cumulative consumption · Investment · Ratio · Cost per beneficiary · Days of cover (lowest item)"],
    ["Trend panel", "Consumption and purchase value by month across project life, with activity markers"],
    ["Composition", "Consumption by category (donut) · by item Pareto (bar + cumulative line) · by service point (bar)"],
    ["Efficiency", "Cost per beneficiary and per visit, project trend versus peer median"],
    ["Stock health", "Items below cover threshold · near-expiry batches held · stock-out occurrences"],
    ["Detail grid", "Item × month table with drill-through to issue, batch and PO"],
], widths=[3.4, 13.8], font=8.2)
doc.add_heading("C.3 Page 5 – Exceptions and Data Quality", level=2)
table(["Block", "Content"], [
    ["KPI row", "Total exceptions · value at risk · aged over 30 days · resolved this month"],
    ["Grid", "Exception type × project × month, with value, owner, age and status"],
    ["Panels", "Unattributed consumption by month · zero-price lines · duplicate GRN candidates · cancelled POs · master-data gaps"],
    ["Log", "Resolution log with notes, owner and due date"],
], widths=[3.4, 13.8], font=8.2)

# ============================================================ APPENDIX D
doc.add_heading("Appendix D – Sample Dataset Description", level=1)
para("A synthetic dataset accompanies this BRD (VPT_Consumption_Dataset.xlsx) so that requirements, KPI definitions and "
     "exception logic can be tested before Oracle access is granted. It is realistic in structure, seasonality, price "
     "behaviour and volume, and it deliberately contains the defects described in Section 15.")
table(
    ["Sheet", "Content", "Rows (approx.)"],
    [
        ["README", "Coverage, assumptions, link logic and known defects", "—"],
        ["02_Dim_Donor", "Seven donors including own-source funding", "7"],
        ["03_Dim_Project", "Ten projects across Karachi and interior Sindh", "10"],
        ["04_Dim_Item", "Ninety items across six categories", "90"],
        ["05_Dim_Supplier", "Twelve suppliers with lead times and terms", "12"],
        ["06_Dim_Warehouse", "One central store and seven project or field stores", "8"],
        ["07_Dim_Calendar", "Monthly calendar with fiscal year and period key", "27"],
        ["08_Fact_PurchaseOrder", "PO lines with price, quantity, buyer and lead time", "14,100"],
        ["09_Fact_Receipts", "GRN lines with batch, manufacture and expiry dates", "14,350"],
        ["10_Fact_Consumption", "Issue lines with project charged, batch and cost", "21,900"],
        ["11_Fact_Budget", "Budget by project, fiscal year and expense category", "210"],
        ["12_Fact_ServiceDelivery", "Beneficiaries and visits by project and month", "228"],
        ["13_Data_Dictionary", "Field-level definitions and keys", "25"],
        ["14_Data_Quality_Notes", "Known defects and required dashboard handling", "8"],
        ["15_Stockout_Detail", "Unmet demand by project, item and month", "870"],
    ],
    widths=[4.4, 9.6, 3.2], font=8.3, align_center_cols=(2,),
)
para(f"Headline control totals in the sample dataset: consumption (excluding inter-store transfers) "
     f"PKR {S['consumption_billable']/1e6:,.1f} million; procurement {S['purchase_value']/1e6:,.1f} million; "
     f"consumption-to-investment ratio {S['consumption_billable']/S['purchase_value']:.2f}; unattributed consumption "
     f"{S['unattributed_pct']:.1f}% of value; product budget PKR {S['budget_product']/1e6:,.1f} million.",
     italic=True, size=8.5, color=GREY)

# ============================================================ APPENDIX E
doc.add_heading("Appendix E – Requirement Traceability", level=1)
para("Each business objective is met only if the linked requirements are delivered and accepted.")
table(
    ["Objective", "Business requirements", "Key functional requirements", "KPI evidence"],
    [
        ["BO-1 Visible consumption", "BR-01, BR-12", "FR-01, FR-02, FR-05, FR-35, FR-41, FR-42", "KPI-01, KPI-02, KPI-30"],
        ["BO-2 Aligned procurement", "BR-02, BR-04", "FR-03, FR-04, FR-13, FR-19, FR-40", "KPI-05, KPI-06, KPI-07, KPI-17"],
        ["BO-3 Value for money", "BR-05, BR-07", "FR-14, FR-15, FR-24, FR-25, FR-28", "KPI-14, KPI-15, KPI-23, KPI-24, KPI-25"],
        ["BO-4 Reduced effort", "BR-11, BR-12", "FR-33, FR-34, FR-45, FR-47", "KPI-30"],
        ["BO-5 Control and audit", "BR-08", "FR-10, FR-29, FR-30, FR-36, FR-39, FR-48", "KPI-27"],
        ["BO-6 Budget control", "BR-09", "FR-17, FR-18, FR-32, FR-44", "KPI-11, KPI-12, KPI-13"],
        ["BO-7 Internal capability", "BR-10, BR-11", "FR-37, FR-38, FR-45, FR-47", "Adoption metric BN-08"],
    ],
    widths=[3.6, 3.0, 6.6, 4.0], font=8.0,
)
doc.add_paragraph()
para("— End of document —", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=8.5, color=GREY)

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                   "BRD - Consumption Analysis Dashboard (Vital Pakistan Trust).docx")
doc.save(out)
print("saved", out)

#!/usr/bin/env python3
"""Builds the Analytics Question Bank & Data Commentary Guide (DOCX)."""
import os
import json
import calendar

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import qa_content as Q

BASE = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE, "..", "data", "findings.json")) as f:
    F = json.load(f)
with open(os.path.join(BASE, "..", "data", "whatif_results.json")) as f:
    W = json.load(f)
CFG_MODEL = W["lever_spec"]

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
doc.styles["Heading 1"].paragraph_format.space_before = Pt(13)
doc.styles["Heading 2"].font.size = Pt(11.5)
doc.styles["Heading 2"].paragraph_format.space_before = Pt(9)
doc.styles["Heading 3"].font.size = Pt(10.5)
doc.styles["Heading 3"].font.color.rgb = GREY

sec = doc.sections[0]
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.left_margin = Cm(1.9)
sec.right_margin = Cm(1.9)
USABLE = 17.2


def shade(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:fill"), color)
    tcPr.append(sh)


def set_cell(cell, text, bold=False, size=8.4, color=None, align=None):
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


def rule(p):
    pPr = p._p.get_or_add_pPr()
    b = OxmlElement("w:pBdr")
    bot = OxmlElement("w:bottom")
    bot.set(qn("w:val"), "single")
    bot.set(qn("w:sz"), "8")
    bot.set(qn("w:space"), "3")
    bot.set(qn("w:color"), ACCENT)
    b.append(bot)
    pPr.append(b)


def para(text="", bold=False, italic=False, size=None, color=None, space_after=5, align=None):
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
    return p


def bullet(text, bold_lead="", level=0):
    p = doc.add_paragraph(style="List Bullet")
    if bold_lead:
        p.add_run(bold_lead).bold = True
    p.add_run(text)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.left_indent = Cm(0.6 + 0.5 * level)
    return p


def numbered(text, bold_lead=""):
    p = doc.add_paragraph(style="List Number")
    if bold_lead:
        p.add_run(bold_lead).bold = True
    p.add_run(text)
    p.paragraph_format.space_after = Pt(2.5)
    return p


def table(headers, rows, widths, font=8.3, align_center=(), header_fill=ACCENT, zebra=True):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        set_cell(c, h, bold=True, size=font, color=RGBColor(0xFF, 0xFF, 0xFF))
        shade(c, header_fill)
    trPr = t.rows[0]._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            set_cell(cells[i], v, size=font,
                     align=WD_ALIGN_PARAGRAPH.CENTER if i in align_center else None)
            if zebra and ri % 2 == 1:
                shade(cells[i], LIGHTER)
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Cm(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return t


def callout(text, size=9.5, fill=LIGHT):
    t = doc.add_table(rows=1, cols=1)
    t.style = "Table Grid"
    c = t.rows[0].cells[0]
    shade(c, fill)
    set_cell(c, text, size=size)
    c.width = Cm(USABLE)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def caption(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = True
    r.font.size = Pt(8.3)
    r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(8)


def page_break():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


p = sec.footer.paragraphs[0]
r = p.add_run("Vital Pakistan Trust  |  Analytics Question Bank & Data Commentary Guide  |  v1.0  |  Internal – Confidential        Page ")
r.font.size = Pt(7.5)
r.font.color.rgb = GREY
fld = OxmlElement("w:fldSimple")
fld.set(qn("w:instr"), "PAGE")
p._p.append(fld)

# ============================================================ COVER
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(64)
r = p.add_run("VITAL PAKISTAN TRUST")
r.font.size = Pt(11)
r.bold = True
r.font.color.rgb = GREY

p = doc.add_paragraph()
r = p.add_run("Analytics Question Bank")
r.font.size = Pt(30)
r.bold = True
r.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
r = p.add_run("& Data Commentary Guide")
r.font.size = Pt(18)
r.font.color.rgb = NAVY
p.paragraph_format.space_after = Pt(2)

p = doc.add_paragraph()
r = p.add_run("Consumption Analysis in View of Current Projects and Products")
r.font.size = Pt(12)
r.font.color.rgb = GREY
p.paragraph_format.space_after = Pt(12)
rule(p)

table(["Item", "Detail"], [
    ["Document reference", "VPT-QB-PRC-002"],
    ["Version / date", "1.0 – 29 October 2026"],
    ["Companion documents", "Problem Statement v1.0; BRD VPT-BRD-PRC-001; Analytics_Question_Bank.xlsx; VPT_Consumption_Analysis.xlsx"],
    ["Business owner", "Head of Procurement & Supply Chain"],
    ["Data as at", f"{F['as_at']} (reporting window {F['start']} to {F['as_at']})"],
    ["Purpose", "A standing set of analytical questions with a live answer engine, and a rule-based method for writing data commentary at every refresh."],
    ["Distribution", "Internal – Confidential"],
], widths=[4.2, 13.0], font=9)

doc.add_paragraph()
para("The questions in this document do not change when the data refreshes. The answers, the status flags and the "
     "commentary do — they are recomputed from the current extract and stamped with the data-as-at date. "
     "Every figure quoted in this guide is a snapshot illustrating the method, not a fixed result.",
     italic=True, size=8.6, color=GREY)

page_break()

# ============================================================ CONTENTS
doc.add_heading("Contents", level=1)
for num, title in [
    ("1.", "Why a question bank, and how it stays current"),
    ("2.", "The analytic ladder: descriptive → diagnostic → predictive → prescriptive"),
    ("3.", "Descriptive questions – what happened"),
    ("4.", "Diagnostic questions – why it happened"),
    ("5.", "Predictive questions – what happens next"),
    ("6.", "Prescriptive questions – what to do, when and who"),
    ("7.", "Live answers at this refresh"),
    ("8.", "The data commentary method"),
    ("9.", "Commentary rules by metric"),
    ("10.", "Worked commentary examples"),
    ("11.", "Cadence, ownership and escalation"),
    ("12.", "Refresh governance and thresholds"),
    ("13.", "Data gaps and questions to add next"),
    ("14.", "What-if and scenario analysis"),
    ("15.", "Parameters, configuration and change control"),
]:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(0.4)
    r = p.add_run(f"{num}\t{title}")
    r.font.size = Pt(9.5)
    p.paragraph_format.tab_stops.add_tab_stop(Cm(1.2))

# ============================================================ 1
doc.add_heading("1. Why a question bank, and how it stays current", level=1)
para("A dashboard answers questions. If the questions are not written down, every review meeting starts from a different "
     "place, analysts rebuild the same analysis each month, and figures are compared without agreeing what is being "
     "measured. This question bank fixes the analysis agenda: the same questions, asked in the same order, at the same "
     "grain, every period — so that what changes between periods is the answer, not the method.")
doc.add_heading("1.1 The separation that makes it refresh-safe", level=2)
table(["Layer", "What it contains", "Does it change on refresh?"], [
    ["Question layer", "The question, why it matters, data fields, logic, output, cadence, owner (sheets 02–05 of the workbook)",
     "No. It is the standing agenda and changes only through change control."],
    ["Threshold layer", "Tolerances and bands: ratio band, variance %, cover days, expiry windows, materiality (sheet 01)",
     "Only by business decision, not by data."],
    ["Answer layer", "The computed value, the status flag and the commentary text (sheets 06 and 08)",
     "Yes. Recomputed from the current extract at every refresh and stamped with the data-as-at date."],
], widths=[3.0, 9.6, 4.6], font=8.5)
para("This is why the workbook needs no editing when data changes: the questions, the logic and the thresholds are stable, "
     "and the answers are derived. Re-running the refresh regenerates every value, every Red/Amber/Green flag and every "
     "commentary line.")
doc.add_heading("1.2 What each question record contains", level=2)
for label, text in [
    ("Question ID", " A stable identifier (D-01, G-07, X-03, R-14) so that a meeting, a report or an audit can reference "
                    "the question unambiguously."),
    ("Why it matters", " The decision the answer supports. A question without a decision behind it is removed."),
    ("Data required", " The exact fields, so the answer can be reproduced from Oracle or from the extract."),
    ("How to answer", " The calculation or comparison logic — written so that two analysts would produce the same number."),
    ("Trigger", " The threshold at which the answer requires commentary or action."),
    ("Cadence and owner", " When the question is asked, and who is accountable for the answer and the resulting action."),
]:
    bullet(text, bold_lead=label)

# ============================================================ 2
doc.add_heading("2. The analytic ladder: descriptive → diagnostic → predictive → prescriptive", level=1)
para("Each stage answers a different question and produces a different kind of output. Skipping a stage is the most common "
     "cause of bad decisions: a prediction made without a diagnosis, or an action taken without a prediction.")
table(["Stage", "Question it answers", "Output", "Typical VPT use", "Failure mode if skipped"], [
    ["Descriptive", "What happened, how much, where, when?",
     "Facts, totals, trends, rankings, distributions",
     "Monthly consumption and investment position; project and product totals",
     "Opinions replace facts; every meeting re-litigates the numbers"],
    ["Diagnostic", "Why did it happen?",
     "Ranked causes with supporting evidence",
     "Explaining a variance, a price rise, an expiry, an unattributed balance",
     "The first plausible cause is accepted and the wrong action is taken"],
    ["Predictive", "What is likely to happen next if nothing changes?",
     "Projections with stated assumptions and a range",
     "Stock-out dates, year-end budget position, expiry risk, requirement quantities",
     "Action is taken late; problems arrive as surprises"],
    ["Prescriptive", "What should we do, when, and who does it?",
     "Actions with owner, deadline and escalation path",
     "Purchase plan, freeze list, redistribution, budget reallocation, write-off",
     "Analysis is produced but behaviour does not change"],
], widths=[2.2, 4.4, 3.6, 4.0, 3.0], font=8.0)
doc.add_heading("2.1 The four questions asked of any movement", level=2)
table(["Question", "What a good answer looks like", "Where it lives"], [
    ["What happened?", "A number, a period, a scope and a comparison — never an adjective on its own.",
     "Descriptive questions and the live answer sheet"],
    ["Why did it happen?", "A ranked list of candidate causes with the evidence that supports or eliminates each one.",
     "Diagnostic questions and the commentary 'why' field"],
    ["So what?", "The consequence in money, service or compliance terms, and who is affected.",
     "Commentary 'impact' field"],
    ["Now what — and when, and who?", "A specific action, a deadline, an owner and an escalation route.",
     "Prescriptive questions, action tracker and the 'when/who' fields"],
], widths=[3.4, 9.0, 4.8], font=8.4)
callout("Rule of the bank: no commentary without a number, no number without a cause, no cause without an action, and no "
        "action without an owner and a date.")

page_break()

# ============================================================ 3-6 question tables
def question_section(title, intro, rows, ref_map):
    doc.add_heading(title, level=1)
    para(intro)
    table(["ID", "Question", "Why it matters / decision supported", "Data required", "How to answer", "Trigger", "Cadence", "Owner"],
          [[r[0], r[1], r[2], r[3], r[4], r[6], r[7], r[8]] for r in rows],
          widths=[1.2, 4.4, 3.8, 2.8, 3.4, 1.6, 1.5, 1.5], font=7.3,
          align_center=(0, 6, 7))


question_section("3. Descriptive questions – what happened",
                 "Ask these first, every refresh. They establish the facts that every later stage depends on. "
                 "Each has one correct answer at a given data-as-at date, so there is nothing to debate — only to verify.",
                 Q.DESCRIPTIVE, None)
doc.add_heading("3.1 Notes on the descriptive layer", level=2)
bullet("Always state the as-at date and the filter scope with the number: a value without its scope is not an answer.")
bullet("Consumption excludes inter-store transfers; if transfers are included, totals will double count.")
bullet("Ratios are meaningless where the denominator is zero — the answer must be blank with a note, not zero.")
bullet("Where a value is derived from two sources (consumption and beneficiaries, for example), report the data completeness "
       "of both before reporting the ratio.")

page_break()
question_section("4. Diagnostic questions – why it happened",
                 "Diagnostic questions force causes to be tested in a ranked order rather than guessed. The output is not a "
                 "number but an evidence-backed explanation — and, where the evidence is insufficient, an explicit statement "
                 "of what is missing.",
                 Q.DIAGNOSTIC, None)
doc.add_heading("4.1 How to run a diagnosis", level=2)
numbered("State the observation precisely: metric, scope, period, and the size of the movement.")
numbered("List the plausible causes in order of likelihood for this organisation — the commentary rules in Section 9 provide the standard order.")
numbered("For each cause, state the evidence that would confirm it and the evidence that would eliminate it.")
numbered("Test against the data; record what was eliminated as well as what was confirmed.")
numbered("Convert the confirmed cause into a prescriptive question — and only then into an action.")
para("A diagnosis that names a cause without citing the test that confirmed it is a hypothesis, and should be labelled as one.")

page_break()
question_section("5. Predictive questions – what happens next",
                 "Predictions are generated from the latest data at each refresh. Each carries a stated assumption and is "
                 "recomputed when the data changes, so a projection can never quietly become stale. They are decision aids, "
                 "not commitments.",
                 Q.PREDICTIVE, None)
doc.add_heading("5.1 Standing assumptions used in this release", level=2)
table(["Projection", "Method", "Main assumption", "Sensitivity"], [
    ["Next-period consumption", "Run rate of the last three to twelve months", "Activity continues at the recent level", "Seasonality and project start/end dates"],
    ["Seasonally adjusted forecast", "Run rate × seasonal index for the target month", "The historical seasonal pattern repeats", "Outbreak or campaign activity"],
    ["Stock-out date", "Days of cover from the report date", "Consumption continues at the 90-day average", "Demand surge; delayed delivery"],
    ["Requirement quantity", "Consumption × (cover target + lead time) − on hand − open commitment", "Lead time holds and stock is issued FEFO", "Supplier delay; pack-size rounding"],
    ["Year-end budget position", "Year-to-date + burn rate × remaining months", "Spending pattern is broadly even", "Lumpy capital purchases; funding changes"],
    ["Expiry risk", "Batch expiry compared with days of cover", "Stock continues to be issued at the current rate", "Redistribution; protocol change"],
], widths=[3.4, 5.4, 4.6, 3.8], font=8.2)

page_break()
question_section("6. Prescriptive questions – what to do, when and who",
                 "These questions end in an action, an owner and a date. An action without an owner is not an action, and an "
                 "action without a date is a preference. The workbook's action tracker is the record of these answers.",
                 Q.PRESCRIPTIVE, None)
doc.add_heading("6.1 Standard action library", level=2)
para("The same small set of levers solves most consumption problems. Naming the lever keeps actions consistent.")
table(["Lever", "What it does", "Applies to", "Typical owner"], [
    ["Buy", "Release, fast-track or add a purchase order", "Cover below threshold; stock-out risk", "Procurement officer"],
    ["Defer / hold", "Stop or delay ordering until cover normalises", "Overstock; slow-moving; year-end", "Head of Procurement"],
    ["Transfer", "Move stock between warehouses or projects", "Surplus in one store, deficit in another", "Supply chain officer"],
    ["Consume first", "Change issue sequencing to FEFO", "Near-expiry batches", "Stores in-charge"],
    ["Write off", "Recognise the loss and remove from stock", "Expired or damaged stock", "Stores + Finance"],
    ["Reallocate", "Move budget between categories or projects", "Over/under-utilisation", "Finance & Grants"],
    ["Correct", "Fix coding, master data or duplicate records", "Unattributed consumption; data defects", "Procurement + IT"],
    ["Renegotiate", "Re-quote, invoke contract or change supplier", "Price variance; supplier performance", "Procurement"],
    ["Redesign", "Change the process or system control so the cause cannot recur", "Repeat exceptions over three periods", "Business owner"],
], widths=[2.4, 5.2, 5.4, 4.2], font=8.2)

page_break()

# ============================================================ 7 LIVE ANSWERS
doc.add_heading("7. Live answers at this refresh", level=1)
para(f"The table below reproduces a selection of the answer sheet as at {F['as_at']}. In the workbook, all questions are "
     f"answered and every value is recomputed at the next refresh — the numbers here will therefore be different after the "
     f"next load, while the questions and the thresholds stay the same.")
from qa_live import build_live

LIVE_ROWS = build_live(F)


def pick(qid, needle):
    for q, stage, question, answer, unit, target, status, action, owner in LIVE_ROWS:
        if q == qid and needle.lower() in question.lower():
            return answer, unit, target, status, question
    return None


def fmt_answer(answer, unit):
    if isinstance(answer, (int, float)) and not isinstance(answer, bool):
        if unit == "PKR":
            return f"PKR {answer:,.0f}"
        if unit == "%":
            return f"{answer:.1%}" if abs(answer) <= 1.5 else f"{answer:,.1f}%"
        if unit == "Ratio":
            return f"{answer:.2f}"
        return f"{answer:,.0f}"
    return str(answer)


SELECT = [
    ("D-01", "total consumption value"), ("D-02", "procurement investment"), ("D-03", "ratio"),
    ("D-05", "top 10"), ("D-07", "peak consumption month"),
    ("D-12", "less than 60 days"), ("D-12", "lowest cover"),
    ("D-13", "expired stock"), ("D-13", "expiring within 60"),
    ("D-14", "slow / non-moving items"), ("D-15", "open balance"),
    ("D-16", "cancelled"), ("D-17", "average unit price"),
    ("D-19", "cost per beneficiary"), ("D-20", "budget utilisation"),
    ("D-21", "unattributed"), ("D-26", "ending within 120"),
    ("X-01", "full FY"), ("X-04", "expected to expire"), ("X-06", "unspent budget exposure"),
    ("R-05", "critical items"),
]
live = []
for qid, needle in SELECT:
    got = pick(qid, needle)
    if not got:
        continue
    answer, unit, target, status, question = got
    live.append((qid, question, fmt_answer(answer, unit), target, status))

table(["Q ID", "Question", "Answer at this refresh", "Target / band", "Status"],
      [[a, b, c, d, e] for a, b, c, d, e in live],
      widths=[1.3, 6.6, 4.4, 3.0, 1.9], font=8.0, align_center=(0, 4))
caption("Status colours: Green = within tolerance; Amber = outside tolerance, not yet material; Red = material breach "
        "requiring commentary and an owner; Info = informational.")

# ============================================================ 8 METHOD
page_break()
doc.add_heading("8. The data commentary method", level=1)
para("Commentary is the bridge between a dashboard and a decision. VPT commentary follows a fixed six-part structure. "
     "Every part is required; if a part cannot be completed, that is itself the finding — usually a data gap or an "
     "unassigned owner.")
table(["Part", "Completing question", "Standard", "Bad practice it prevents"], [
    ["1. What happened", "Which metric, in which scope and period, moved by how much?",
     "Cite the number, the period and the comparison. One sentence.",
     "“Consumption was high” — with no figure, scope or comparison."],
    ["2. Why it happened", "Which causes were tested, in what order, and which one the evidence supports?",
     "Name the confirmed cause and the test; list what was eliminated.",
     "Jumping to the first plausible explanation."],
    ["3. So what", "What is the consequence in money, service or compliance terms, and who is affected?",
     "Quantify the impact; name the affected projects, services or donors.",
     "Stopping at the explanation, leaving the reader to work out why it matters."],
    ["4. What to do", "Which lever from the action library applies, and in what quantity?",
     "One primary action, with any supporting action; no more than three.",
     "A long list of options with no decision."],
    ["5. When", "By when must the action be taken for it still to matter?",
     "A date or a deadline convention (for example, before the next requisition cycle).",
     "“As soon as possible”, which is never."],
    ["6. Who", "Who owns the action, and to whom do they escalate?",
     "A named role, an escalation route and the ageing limit.",
     "“The team” — which means nobody."],
], widths=[2.6, 4.6, 5.4, 4.6], font=8.1)

doc.add_heading("8.1 Writing standards", level=2)
for b in [
    "Lead with the number, then the cause, then the action. Managers read the first line and act on the last.",
    "Quantify everything: value, quantity, days, percentage, number of projects or items affected.",
    "Separate fact from inference. If a cause is probable but unproven, say “probable” and state the test that would confirm it.",
    "Never write commentary on a value that failed a quality gate — fix the data first.",
    "Use consistent units and the same measurement basis as the dashboard; do not recompute totals inside the commentary.",
    "Where a variance is inside tolerance, say so in one line and move on. Commentary effort belongs where the value is.",
    "End every commentary line with the owner and the date, so it can be tracked without re-reading the paragraph.",
]:
    bullet(" " + b)

doc.add_heading("8.2 When commentary is mandatory", level=2)
table(["Trigger", "Threshold", "Minimum commentary required"], [
    ["Budget variance", "Above 15% or above PKR 250,000, whichever is lower", "Full six-part commentary"],
    ["Consumption-to-investment ratio", "Outside 0.80 – 1.05", "Full six-part commentary"],
    ["Unattributed consumption", "Above 1% of value", "Full six-part commentary plus a corrective action for the coding process"],
    ["Expiry or wastage", "Any value above PKR 100,000", "Full six-part commentary with disposition"],
    ["Price variance", "Above 10% on any material item", "Cause, negotiation target and budget effect"],
    ["Stock-out or cover below 30 days", "Any critical item", "Action, date and service impact"],
    ["Data-quality gate failure", "Any failed gate", "Two lines: cause and reload plan; publication blocked"],
], widths=[4.6, 4.6, 8.0], font=8.2)

# ============================================================ 9 RULES
page_break()
doc.add_heading("9. Commentary rules by metric", level=1)
para("These are the standing rules applied at each refresh. The causes are listed in the order they should be tested — "
     "from most to least likely in this organisation's data — and the actions are drawn from the action library in "
     "Section 6.1. The workbook holds the same rules with identifiers (CR-01 to CR-11) so that a commentary line can cite "
     "the rule it was written under.")
for i, (metric, why, action, when, owner) in enumerate(Q.COMMENTARY_RULES, start=1):
    doc.add_heading(f"9.{i} {metric}", level=2)
    table(["Element", "Content"], [
        ["Causes to test, in order", "\n".join(f"{j}. {w}" for j, w in enumerate(why, start=1))],
        ["Action options", "\n".join(f"{j}. {a}" for j, a in enumerate(action, start=1))],
        ["When to act", when],
        ["Owner", owner],
    ], widths=[4.2, 13.0], font=8.2)

# ============================================================ 10 WORKED EXAMPLES
page_break()
doc.add_heading("10. Worked commentary examples", level=1)
para(f"The four examples below are written from the current extract ({F['as_at']}) using the six-part structure. "
     f"Numbers will change at the next refresh; the structure does not.")

bo = F["budget_over"][0] if F["budget_over"] else None
ending = [p for p in F["projects_ending"] if p["days_left"] <= 120 and p["unspent"] > 0]

doc.add_heading("10.1 Example 1 – a category budget over-utilisation", level=2)
if bo:
    table(["Part", "Commentary"], [
        ["What happened", f"{bo['cat']} consumption in {bo['fy']} reached PKR {bo['consumption']:,.0f} against an approved "
                          f"budget of PKR {bo['budget']:,.0f} — an overrun of PKR {bo['consumption'] - bo['budget']:,.0f} "
                          f"({bo['util'] - 1:.1%})."],
        ["Why (tested)", "Price effect was tested first using the per-item price trend: the weighted average unit price of the "
                         "category's top items rose across the same period, and volume growth alone does not account for the gap. "
                         "Mis-posting was eliminated by checking project-level postings against service volumes. Cause accepted: "
                         "price movement on high-value items combined with service volume above the planning assumption."],
        ["So what", "The overrun can only be funded from another category or from new money. Left unresolved it becomes a "
                    "donor-reporting exception for the year and a reason for reduced activity in the final quarter."],
        ["What to do", "Reallocate from the under-utilised capital line (see Example 3); renegotiate the two highest-value items "
                       "before the next order; tighten issue authorisation for the category; record the price/volume split in the "
                       "variance note."],
        ["When", "Commentary within 3 working days of the close; reallocation before the fiscal year reaches 75% elapsed."],
        ["Who", "Manager Finance & Grants, escalating to Director Finance & Operations."],
    ], widths=[2.8, 14.4], font=8.4)

doc.add_heading("10.2 Example 2 – unattributed consumption", level=2)
table(["Part", "Commentary"], [
    ["What happened", f"PKR {F['unattributed']:,.0f} of consumption ({F['unattributed_lines']:,} issue lines, "
                      f"{F['unattributed_pct']:.1f}% of value) is charged to a generic cost centre rather than to a project; the "
                      f"worst month was {F['dq']['worst_month']['label']} at {F['dq']['worst_month']['share']:.1%} of that month's value."],
    ["Why (tested)", "The exception was cross-tabulated by user, warehouse, service point and month. Concentration in particular "
                     "warehouses and months rather than across all users eliminated a system-wide defect and pointed to a local "
                     "working practice: a default code used when the project is not immediately known at the point of issue."],
    ["So what", "Project-level consumption, budget utilisation and cost per beneficiary are all understated for the projects that "
                "actually consumed the stock. One donor's commodities appear unaccounted for in the project record."],
    ["What to do", "Correct retroactively where the consuming project is evidenced; make the project field mandatory and validated "
                   "at issue; retrain the users identified; add this exception to the monthly review agenda until it stays below 1% "
                   "for three consecutive months."],
    ["When", "Correction within the current month; the system control change within the quarter."],
    ["Who", "Head of Procurement with IT; escalation to Director Finance & Operations if the share remains above 1.5% for two periods."],
], widths=[2.8, 14.4], font=8.4)

doc.add_heading("10.3 Example 3 – expiry and unspent capital budget", level=2)
table(["Part", "Commentary"], [
    ["What happened", f"PKR {F['stock']['expired_value']:,.0f} of stock is past expiry and not yet written off "
                      f"({F['stock']['expired_share']:.1%} of stock value), and a further PKR {F['stock']['near_expiry_60']:,.0f} expires "
                      f"within 60 days. In the same period, the capital (equipment) budget used only "
                      f"{F['budget_under'][0]['util']:.1%} of its allocation."],
    ["Why (tested)", "Batch tracing showed the expiry is concentrated in short-dated receipts and in items ordered above consumption. "
                     "The capital line is unused because requests are stuck in approval and quotation and because the requirement was "
                     "overstated at planning. The two are linked: budget sits unused in one line while the consumable line overruns."],
    ["So what", f"Avoidable loss of PKR {F['stock']['expired_value'] + F['stock']['near_expiry_60']:,.0f} at risk, plus an audit "
                f"observation on stock management, plus a service-capability gap where equipment was planned but not purchased."],
    ["What to do", "Write off the expired stock with approval; redistribute the near-expiry balance to a project with demand; "
                   "reject short-dated deliveries at receipt; review and either approve or cancel the capital pipeline and reallocate "
                   "the released budget to the over-committed category."],
    ["When", "Write-off and redistribution within 5 working days of the flag; capital decision this quarter, before the procurement "
             "window closes."],
    ["Who", "Stores in-charge for the stock action; Head of Procurement for the capital decision; Finance for the write-off and reallocation."],
], widths=[2.8, 14.4], font=8.4)

doc.add_heading("10.4 Example 4 – a projection that requires a decision now", level=2)
table(["Part", "Commentary"], [
    ["What happened", f"With {F['fy2026_27']['months']} months of FY2026-27 elapsed and PKR {F['fy2026_27']['ytd']:,.0f} consumed, the "
                      f"burn rate of PKR {F['fy2026_27']['burn']:,.0f} per month projects PKR {F['fy2026_27']['projection']:,.0f} for the "
                      f"full year against a product budget of PKR {F['budget_fy27']:,.0f}."],
    ["Why (tested)", "Budget phasing was checked first: the full-year budget may simply not yet be loaded. If it is loaded, the cause is "
                     "either under-budgeting against price inflation or consumption above the planned service volume — tested by "
                     "comparing unit price trend with service volume trend."],
    ["So what", "If the gap is real, service levels cannot be maintained for the full year, and the shortfall will appear "
                "unannounced in the last quarter when there is no time to respond."],
    ["What to do", "Confirm budget phasing with Finance this week. If the gap is real: quantify it by project and category, protect "
                   "critical medicines first, and prepare a funding request or a documented reprioritisation for the donor."],
    ["When", "Now — at the first quarter, not the ninth month."],
    ["Who", "Manager Finance & Grants, escalating to Director Finance & Operations and the steering group."],
], widths=[2.8, 14.4], font=8.4)

# ============================================================ 11 CADENCE
page_break()
doc.add_heading("11. Cadence, ownership and escalation", level=1)
doc.add_heading("11.1 Which questions are asked when", level=2)
table(["Cadence", "Questions", "Output", "Forum"], [
    ["Daily", "D-22 (data freshness and completeness), G-18 (duplicate posting), R-12 (quality gates), R-05 (critical stock-out)",
     "Load status and any blocking defect", "IT / BI and stores in-charge"],
    ["Weekly", "D-11 to D-15 (stock, cover, expiry, slow-moving, commitments), X-02 (projected stock-out), R-01 to R-05 (actions)",
     "Procurement action list with owners", "Procurement and stores stand-up"],
    ["Monthly", "D-01 to D-10, D-19 to D-21, D-26 (consumption, ratio, trend, value for money, attribution), G-01 to G-09, G-14, "
                "X-01, X-05, X-06, R-06 to R-08, R-22",
     "Monthly consumption commentary pack and action tracker", "Consumption review meeting (working day 6)"],
    ["Quarterly", "D-09, D-18, D-27 (seasonality, suppliers, distribution), G-10, G-11, G-20, X-07, X-10, X-13, R-07, R-09, R-10, R-14, R-16, R-21",
     "Supplier, price and efficiency review; seasonality preparedness", "Steering group"],
    ["Annually", "R-13 (reorder policy), P-01 to P-15 (threshold review), threshold calibration", "Revised parameters and policy",
     "Steering group with Finance"],
], widths=[2.0, 7.6, 4.2, 3.4], font=8.1)

doc.add_heading("11.2 Ownership and escalation", level=2)
table(["Role", "Owns", "Escalates to", "Escalation trigger"], [
    ["Stores in-charge", "Stock accuracy, expiry, issue recording, physical counts", "Head of Procurement", "Any expired batch; cover below 30 days on a critical item"],
    ["Procurement officer", "Purchase plan, PO release, price negotiation, supplier follow-up", "Head of Procurement", "Lead-time breach beyond 20%; price rise above 10%"],
    ["Supply chain officer", "Transfers, redistribution, warehouse balance, order policy", "Head of Procurement", "Surplus and deficit coexisting for the same item"],
    ["Project manager", "Consumption within project, service-point issuing, coding accuracy", "Programme Manager", "Variance above threshold; consumption per visit off norm by 25%"],
    ["Finance & Grants", "Budget versus consumption, reallocation, donor reporting, write-off", "Director Finance & Operations", "Category over-utilisation; projected funding gap"],
    ["IT / BI", "Extract, load, data quality gates, security, refresh status", "Head of Procurement (business impact)", "Failed gate; duplicate posting; master-data gap"],
    ["Business owner (Head of Procurement)", "Overall interpretation, action assignment, adoption", "Steering group", "Any Red item unresolved beyond 30 days"],
], widths=[4.0, 5.8, 4.0, 3.4], font=8.1)

doc.add_heading("11.3 Action tracking discipline", level=2)
bullet("Every Red commentary line creates an action with an owner, a due date and a value at stake.")
bullet("An action open beyond 30 days is escalated to the business owner; beyond 60 days to the steering group.")
bullet("Actions are closed only with evidence — a GRN number, a corrected issue, an approved write-off, a signed reallocation.")
bullet("The same exception recurring for three consecutive periods is converted into a prescriptive process-change question (R-21).")

# ============================================================ 12 GOVERNANCE
doc.add_heading("12. Refresh governance and thresholds", level=1)
doc.add_heading("12.1 What changes on a refresh, and what does not", level=2)
table(["Element", "Refresh behaviour", "Control"], [
    ["Data values and all answers", "Recomputed from the new extract", "Quality gates must pass before publication"],
    ["Status flags (Red/Amber/Green)", "Recalculated against the parameters in force", "Thresholds are version-controlled"],
    ["Commentary lines", "Regenerated for the metrics that breach tolerance", "A human reviewer signs off the monthly pack"],
    ["Questions, logic and owners", "Unchanged", "Change control, as in the BRD"],
    ["Previously reported figures", "Preserved in the month-end snapshot", "Restatements shown separately with the reason"],
], widths=[4.6, 6.4, 6.2], font=8.2)

doc.add_heading("12.2 Threshold change control", level=2)
bullet("Thresholds are business parameters, not technical settings. Changes require the business owner's approval.")
bullet("Every change is recorded with the old value, the new value, the reason and the effective date.")
bullet("A threshold change applies from the next refresh; historical statuses are not rewritten retroactively.")
bullet("Thresholds are reviewed annually, or sooner if a metric breaches tolerance repeatedly without a credible action.")

doc.add_heading("12.3 Publication gate", level=2)
para("A refresh may be published only when: the extract completed; control totals reconcile to the Oracle registers within "
     "tolerance; duplicates and negative quantities are cleared; master-data exceptions are listed; and the workbook has "
     "been regenerated (no hand-edited values). If any condition fails, the previous snapshot remains live and the failure "
     "is logged with an owner and a reload plan.")

# ============================================================ 13 GAPS
page_break()
doc.add_heading("13. Data gaps and questions to add next", level=1)
para("Some questions cannot yet be answered fully because the enabling data is not in the extract. They remain in the bank "
     "with a gap note, so that the gap is visible and can be closed deliberately rather than forgotten.")
table(["Question", "What is missing", "Workaround for now", "To close the gap"], [
    ["Stock-out days and duration (D-12, X-02)", "A stock-out register or nil-stock transaction log",
     "Inferred from unmet demand and issue gaps", "Record nil-stock incidents in Oracle or a simple register at the store"],
    ["Wastage and write-off value (G-07, X-11)", "Approved inventory adjustment documents",
     "Batch expiry compared with remaining quantity", "Route write-offs through a documented adjustment with a reason code"],
    ["Consumption against treatment protocol norms (G-16)", "Standard consumption norms per service type",
     "Peer comparison between projects", "Agree norms with the clinical lead by service point"],
    ["Split of consumption by age, sex or beneficiary group", "Beneficiary-level dispensing data (not held in ERP)",
     "Not attempted; kept out of scope", "Only if programme M&E introduces dispensing-level data"],
    ["User-level issue behaviour (G-08, R-08)", "A user dimension in the extract",
     "Counts by warehouse and service point", "Add the issuing user to the extract"],
    ["True budget phasing by month (X-05, X-15)", "Monthly budget profile rather than an annual figure",
     "Straight-line assumption", "Finance to load the approved phasing profile"],
    ["Supplier contract price and validity (G-11, R-10)", "Contract reference and price validity dates",
     "Comparison with the previous purchase price", "Add contract number and validity to the PO extract"],
], widths=[4.4, 4.4, 4.2, 4.2], font=8.0)

doc.add_heading("13.1 Questions to add in the next release", level=2)
for q in [
    "Which items show a structural change in consumption after a protocol or guideline revision?",
    "What is the consumption elasticity to service volume — how much additional product does a 10% increase in beneficiaries require?",
    "What is the cost of stock-outs in emergency purchase premium and lost service volume?",
    "Which projects have the highest consumption per service point per month, and is that a case-mix effect or a control weakness?",
    "What would be the financial effect of consolidating the top ten items into a single framework contract?",
    "How does consumption per beneficiary move after a price change, once the volume effect is removed?",
    "Which batches were received short-dated, and which suppliers delivered them?",
]:
    bullet(" " + q)

# ============================================================ 14 WHAT-IF
page_break()
doc.add_heading("14. What-if and scenario analysis", level=1)
WB = W["baseline_result"]
para("Descriptive analysis tells us where we are; the what-if engine tells us what happens if a driver moves. Every lever "
     "in the model is a parameter, so a scenario is a named set of parameter overrides rather than a separate spreadsheet. "
     "The engine is documented in the workbook's Model_Logic sheet and reproduces in a few lines of arithmetic, which is "
     "deliberate: a scenario that management cannot follow is a scenario management will not use.")
doc.add_heading("14.1 How the model works", level=2)
table(["Step", "Equation", "What it produces"], [
    ["Consumption", "annualised run rate × (1 + demand growth) × price factor × activity factor",
     "The forward 12-month value of products to be consumed"],
    ["Price factor", "(1 + inflation) × (1 + one-off shock) × (1 − framework discount)", "The price effect, isolated from volume"],
    ["Target stock", "consumption × (cover days + safety days) / 365", "The stock position that must be funded"],
    ["Purchase requirement", "consumption + top-up − released excess stock + wastage replacement",
     "The buying and cash requirement"],
    ["Budget available", "annual budget × (1 − funding cut) × (1 + budget growth)", "The envelope the plan must fit inside"],
    ["Funding gap", "max(0, consumption − budget available)", "The decision trigger"],
    ["Expiry risk", "near-expiry share × stock + wastage × consumption + excess stock × (1 − redistribution)",
     "Write-off exposure the policy creates"],
    ["Stock-out risk", "items whose projected cover is below lead time + safety stock", "Service risk by item"],
    ["Cost per beneficiary", "consumption ÷ beneficiaries", "Value for money"],
], widths=[3.0, 8.0, 6.2], font=8.2)
para(f"All parameters live in one configuration file (parameter set v{W['config_version']}). The baseline facts — run rate, "
     f"stock, commitment, budget, beneficiaries, lead time — are recomputed from the data at every refresh; only the "
     f"assumptions and the scenarios are entered by a person.", italic=True, size=8.5, color=GREY)

doc.add_heading("14.2 The what-if question set", level=2)
para("Fourteen standing what-if questions sit alongside the descriptive, diagnostic, predictive and prescriptive banks "
     "(sheet 06 of the workbook). They are the questions management actually asks when a driver moves.")
table(["ID", "What-if question", "Levers used", "Decision it supports", "Owner"],
      [[r[0], r[1], r[3], r[2], r[8]] for r in Q.WHATIF],
      widths=[1.3, 6.0, 3.6, 4.2, 2.1], font=7.6, align_center=(0,))

doc.add_heading("14.3 Scenario results at this refresh", level=2)
para(f"Forward 12 months, measured against the baseline. Data as at {W['as_at']}. These figures recalculate on refresh; "
     f"the scenario definitions do not.")
rows = []
for sc in W["scenarios"]:
    r = sc["result"]
    rows.append([sc["id"], sc["name"], f"{r['consumption_value']/1e6:,.1f}", f"{r['purchase_value']/1e6:,.1f}",
                 f"{r['funding_gap']/1e6:,.1f}", f"{r['budget_utilisation']:.1%}", str(r["stockout_items"]),
                 f"{r['expiry_risk_value']/1e6:,.2f}", f"{r['working_capital']/1e6:,.1f}"])
table(["ID", "Scenario", "Consumption (PKR m)", "Purchase (PKR m)", "Funding gap (PKR m)", "Utilisation",
       "Stock-out items", "Expiry risk (PKR m)", "Working capital (PKR m)"],
      rows, widths=[1.2, 5.4, 2.0, 1.9, 1.9, 1.6, 1.5, 1.8, 1.9], font=7.6,
      align_center=(0, 2, 3, 4, 5, 6, 7, 8))
caption("Baseline S0 is the current parameter set; it already carries the approved demand and price assumptions, which is "
        "why it is not identical to the historic period totals in the analysis workbook.")
doc.add_heading("14.4 Reading the scenarios", level=2)
s7 = next(x for x in W["scenarios"] if x["id"] == "S7")["result"]
s8 = next(x for x in W["scenarios"] if x["id"] == "S8")["result"]
s5 = next(x for x in W["scenarios"] if x["id"] == "S5")["result"]
s3 = next(x for x in W["scenarios"] if x["id"] == "S3")["result"]
s6 = next(x for x in W["scenarios"] if x["id"] == "S6")["result"]
bullet(f"Combined downside (S7) opens a funding gap of PKR {s7['funding_gap']/1e6:,.1f} million and pushes "
       f"{s7['stockout_items']} items below the cover threshold. Applying the policy and commercial levers (S8) reduces the "
       f"gap to PKR {s8['funding_gap']/1e6:,.1f} million — PKR {(s7['funding_gap']-s8['funding_gap'])/1e6:,.1f} million, or "
       f"{(1-s8['funding_gap']/s7['funding_gap']):.0%} of the downside, is within management control.",
       bold_lead="The value of acting. ")
bullet(f"Moving from 90 to 60 days of cover (S5) releases PKR "
       f"{(WB['working_capital']-s5['working_capital'])/1e6:,.1f} million of working capital and cuts the purchase "
       f"requirement by PKR {(WB['purchase_value']-s5['purchase_value'])/1e6:,.1f} million, at a small change in expiry risk.",
       bold_lead="The cheapest lever. ")
bullet(f"A 60% lead-time increase with reliability falling to 70% (S3) barely moves the funding gap but takes stock-out items "
       f"from {WB['stockout_items']} to {s3['stockout_items']}. Supply risk does not announce itself in the budget — it shows "
       f"up at the service point.", bold_lead="The quiet risk. ")
bullet(f"Framework contracting (S6) brings the baseline gap to zero, but it cannot absorb a demand or price shock on its own: "
       f"it is a margin lever, not a capacity lever.", bold_lead="The limit of commercial action. ")

doc.add_heading("14.5 Sensitivity – what to watch", level=2)
order = W["sensitivity"]["order"]["funding_gap"]
by_key = {r["key"]: r for r in W["sensitivity"]["rows"]}
para("One lever at a time, moved between a low and a high value, ranked by the effect on the funding gap. Monitoring effort "
     "should follow this ranking.")
table(["Rank", "Lever", "Low → high", "Funding gap at low (PKR m)", "at high (PKR m)", "Swing (PKR m)"],
      [[str(i), key, f"{by_key[key]['low']} → {by_key[key]['high']}",
        f"{by_key[key]['outputs']['funding_gap']['low']/1e6:,.1f}",
        f"{by_key[key]['outputs']['funding_gap']['high']/1e6:,.1f}",
        f"{by_key[key]['outputs']['funding_gap']['swing']/1e6:,.1f}"]
       for i, key in enumerate(order, start=1)],
      widths=[1.3, 5.0, 3.0, 2.8, 2.6, 2.5], font=8.0, align_center=(0, 2, 3, 4, 5))
sw = by_key[order[0]]["outputs"]["stockout_items"]["swing"]
sw_key = order[0]
lead_row = by_key["lead_time_multiplier"]
para(f"On the funding gap, {sw_key} dominates. On stock-out risk the ranking is different: lead time and supplier "
     f"reliability move {lead_row['outputs']['stockout_items']['swing']} items between the low and high case — the largest "
     f"swing on that output. Two different risks, two different monitoring measures.")

doc.add_heading("14.6 Break-even and the cover trade-off", level=2)
para("Break-even converts the budget position into a tolerance that can be watched, and the cover curve shows the cost of "
     "protection.")
be_rows = []
for b in W["breakeven"]:
    val = b["breakeven"]
    if isinstance(val, float) and b["key"].endswith("_pct"):
        shown = f"{val:.1%}"
    elif val is None:
        shown = "n/a"
    else:
        shown = f"{val:,.2f}"
    be_rows.append([b["key"], b["target"], shown, b["note"]])
table(["Lever", "Condition solved", "Break-even value", "Note"], be_rows,
      widths=[3.8, 3.4, 2.6, 7.4], font=8.2)

curve = next((b.get("curve") for b in W["breakeven"] if b.get("curve")), [])
if curve:
    picks = [curve[0], curve[5], curve[10], curve[15], curve[20], curve[-1]]
    table(["Cover policy (days)", "Purchase requirement (PKR m)", "Expiry risk (PKR m)",
           "Working capital (PKR m)", "Stock-out items"],
          [[f"{p['x']:.0f}", f"{p['purchase']/1e6:,.1f}", f"{p['expiry']/1e6:,.2f}",
            f"{p['capital']/1e6:,.1f}", str(p["stockouts"])] for p in picks],
          widths=[3.4, 4.0, 3.4, 3.4, 3.0], font=8.2, align_center=(0, 1, 2, 3, 4))
    para("The curve is flat on the left and rises on the right: below roughly 90 days each additional day of cover buys "
         "protection cheaply; above it, the purchase requirement and working capital rise faster than the risk falls. That "
         "knee is the defensible policy point, and it should be reviewed twice a year.", italic=True, size=8.5, color=GREY)

doc.add_heading("14.7 Two-way view – demand growth × price inflation", level=2)
para("A single break-even point is easy to read; a map is easier to plan with. Rows are demand growth, columns are price "
     "inflation, and each cell is a full model run showing the funding gap in PKR millions.")
tw = W["twoway"]
header_cells = [f"{c:.0%}"] + [f"{v/1e6:,.1f}" for v in tw["tables"]["funding_gap"][0]] if False else None
grid = []
for i, rv in enumerate(tw["row_values"]):
    grid.append([f"{rv:.0%}"] + [f"{v/1e6:,.1f}" for v in tw["tables"]["funding_gap"][i]])
table(["Demand ↓ / Price →"] + [f"{c:.0%}" for c in tw["col_values"]], grid,
      widths=[3.2] + [2.3] * len(tw["col_values"]), font=8.2, align_center=tuple(range(0, 7)))
para("Read across a row to see the price effect and down a column to see the demand effect; the current position sits near "
     "the middle of the grid. Anything in the lower-right quadrant is beyond what commercial action alone can absorb.",
     italic=True, size=8.5, color=GREY)

# ============================================================ 15 PARAMETERS
page_break()
doc.add_heading("15. Parameters, configuration and change control", level=1)
para(f"Nothing in this pack is hard-coded. All {len(Q.PARAMETERS)} thresholds and all {len(Q.MODEL_PARAMETERS)} model "
     f"assumptions are held in a single configuration file, parameter set v{Q.CONFIG_VERSION}. The workbooks, this guide and "
     f"the BRD figures are generated from that file plus the data, so a change is made once and takes effect everywhere.")
doc.add_heading("15.1 Parameters by group", level=2)
groups = {}
for key, name, val, unit, rng, basis, owner in Q.MODEL_PARAMETERS:
    g = next((m["group"] for m in CFG_MODEL if m["key"] == key), "Other")
    if isinstance(val, (int, float)) and unit == "%":
        shown = f"{val:.2%}"
    elif isinstance(val, (int, float)):
        shown = f"{val:g}"
    else:
        shown = str(val)
    groups.setdefault(g, []).append([key, name, shown, unit, rng, owner])
for g, rows in groups.items():
    doc.add_heading(f"{g}", level=3)
    table(["Key", "Parameter", "Value", "Unit", "Working range", "Owner"], rows,
          widths=[4.0, 5.6, 1.8, 1.4, 2.4, 2.0], font=8.0, align_center=(2, 3))
doc.add_heading("15.2 Thresholds", level=2)
table(["Ref", "Threshold", "Value", "Unit", "How it is used"],
      [[p[0], p[1], p[2], p[3], p[4]] for p in Q.PARAMETERS],
      widths=[1.2, 5.2, 2.6, 1.6, 6.6], font=7.8, align_center=(0, 2, 3))
doc.add_heading("15.3 Change control", level=2)
table(["Change", "Approval", "Effect", "Recorded where"], [
    ["Threshold (status band)", "Business owner", "All status flags and commentary lines from the next refresh",
     "Configuration history with old value, new value, reason, date"],
    ["Model assumption", "Business owner with Finance", "Every what-if output and scenario", "Same, plus a note in the quarterly review minutes"],
    ["New or edited scenario", "Business analyst, noted to the owner", "Scenario summary, sensitivity and documents",
     "Configuration file, with an identifier and a rationale"],
    ["New question", "Business owner", "Question bank and coverage map", "Question bank with a BRD reference"],
    ["KPI definition change", "Steering group", "Every reported figure, and comparability with history",
     "BRD change control log"],
], widths=[3.6, 3.4, 6.4, 3.8], font=8.1)
doc.add_heading("15.4 Rules for changing a parameter", level=2)
for t in [
    "Change the value in the configuration file, never in a report or a workbook cell.",
    "Record the old value, the new value, the reason, the approver and the effective date before the next refresh.",
    "Re-run the full refresh so that every artefact is rebuilt from the same numbers; do not hand-edit outputs.",
    "A parameter change applies from the next refresh forward. Historical snapshots keep the parameters they were built with.",
    "Re-read the affected commentary: a threshold change can close open actions or open new ones, and both need a decision.",
    "Review the whole set twice a year, and after any structural change such as a new grant, a new project or a system change.",
]:
    bullet(" " + t)

doc.add_paragraph()
para("— End of document —", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, size=8.5, color=GREY)

OUT = os.path.join(BASE, "..", "Analytics Question Bank & Data Commentary Guide.docx")
doc.save(OUT)
print("saved", OUT)

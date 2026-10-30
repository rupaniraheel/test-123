# Vital Pakistan Trust — Consumption Analysis Dashboard

Business analysis pack for the dashboard **“Consumption Analysis in View of Current Projects and Products”**
(data source: Oracle ERP — procurement, inventory, finance).

## Deliverables

| File | What it is |
|---|---|
| `Problem Statement - Consumption Analysis Dashboard.docx` | One-to-two page problem definition (context, problem, impact, scope, KPIs, next steps). |
| `BRD - Consumption Analysis Dashboard (Vital Pakistan Trust).docx` | Extensive Business Requirements Document — ~12,300 words, 49 tables, 23 sections plus 5 appendices. |
| `data/VPT_Consumption_Dataset.xlsx` | Realistic synthetic Oracle-style dataset: 27 months (Jul-2024 → Sep-2026), 14,100 PO lines, 16,800 GRN lines, 23,200 issue lines, plus dimensions, budget and service delivery. |
| `data/VPT_Consumption_Analysis.xlsx` | Analysis / dashboard workbook built **from** that dataset — 15 sheets of KPIs, trends, project, product, budget, price, stock, supplier and exception views with charts. |
| `Analytics Question Bank & Data Commentary Guide.docx` | 84-question analytical bank (descriptive / diagnostic / predictive / prescriptive), the six-part data-commentary method, worked examples, cadence, ownership and refresh governance. |
| `data/Analytics_Question_Bank.xlsx` | Refresh-driven workbook: parameters, the four question banks, **live answers with status flags**, commentary rules and the generated commentary feed, refresh checklist. |

## BRD contents (at a glance)

1. Executive summary · 2. Introduction · 3. Business context and objectives · 4. Problem statement
5. Scope and phasing · 6. Stakeholders and RACI · 7. As-is process · 8. To-be process and user journeys
9. 12 business requirements · 10. **48 functional requirements** (MoSCoW, acceptance criteria, traceability)
11. 5-page dashboard specification · 12. **30-KPI dictionary** (formula, grain, target, owner)
13. Oracle data requirements, sourcing options and **14 transformation rules**
14. **14 business rules / calculations** · 15. Data-quality requirements and control-total reconciliation
16. 16 non-functional requirements · 17. Assumptions, constraints, dependencies · 18. Risks
19. Acceptance criteria and 14 UAT scenarios · 20. Delivery phases and timeline · 21. Benefits and success measures
22. Governance, change control, support · 23. Glossary · Appendices A–E (data model, Oracle table mapping,
wireframes, dataset description, requirement traceability)

## Dataset design notes

* Synthetic but internally consistent: consumption is issued **FEFO** from received batches, so stock never goes negative.
* Realistic behaviour: ~9% p.a. price inflation, programme growth, monsoon-driven demand surge (Aug–Sep 2025),
  a funding-shock dip (Mar 2026), pack-size-rounded ordering, supplier lead times, batch and expiry tracking.
* Deliberate, documented defects so the dashboard’s exception logic is testable:
  unattributed project codes (~2.1% of value), zero-value/donated purchases, duplicate GRNs, cancelled POs,
  blacklisted supplier with history, near-expiry and expired stock, unmet demand.
* Headline control totals: consumption **PKR 259.5m**, procurement **PKR 301.3m**, ratio **0.86**,
  product budget **PKR 337.1m** (77% utilised).

Organisation, supplier and product names are used for realism only. Replace with Oracle extracts before external use.

## Refresh model (answers change, questions do not)

The question bank is the stable layer: question, purpose, data fields, logic, cadence and owner.
The answers, status flags and commentary are derived, so a data refresh changes them without anyone editing the workbook.

```bash
python3 scripts/refresh_all.py                # full chain: data → findings → workbooks → docs
python3 scripts/refresh_all.py --no-generate  # refresh answers from the existing dataset only
```

Chain: `generate_dataset.py` → `analyze_findings.py` (all metrics → `data/findings.json`) →
`build_analysis.py` → `build_qa_workbook.py` → `build_qa_doc.py` → `build_brd.py`.
Thresholds live in `scripts/qa_content.py` (`PARAMETERS`) and are reproduced in the workbook's `01_Parameters` sheet.

Requires `openpyxl` and `python-docx`.

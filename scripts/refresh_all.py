#!/usr/bin/env python3
"""
Single refresh command for the VPT consumption-analytics pack.

    python3 scripts/refresh_all.py               # rebuild sample data + all answers
    python3 scripts/refresh_all.py --no-generate # refresh answers from the current dataset

Order of execution
  1. generate_dataset.py   – (optional) regenerate the Oracle-style sample extract
  2. analyze_findings.py   – recompute every metric into findings.json
  3. build_analysis.py     – rebuild the analysis / dashboard workbook
  4. build_qa_workbook.py  – rebuild the question bank, live answers and commentary feed
  5. build_qa_doc.py       – rebuild the question bank and commentary guide
  6. build_brd.py          – refresh the BRD's headline figures

Nothing in the answers or in the commentary is hard-coded: steps 2–6 derive
everything from the dataset, so a refreshed extract changes every value,
status flag and commentary line while the questions and thresholds stay fixed.
"""
import subprocess
import sys
import os
import json

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)


def run(script, args=()):
    print(f"\n=== {script} {' '.join(args)} ===")
    r = subprocess.run([sys.executable, os.path.join(BASE, script), *args], cwd=BASE)
    if r.returncode != 0:
        print(f"FAILED: {script}")
        sys.exit(r.returncode)


def main():
    generate = "--no-generate" not in sys.argv

    if generate:
        run("generate_dataset.py")
    elif not os.path.exists(os.path.join(ROOT, "data", "VPT_Consumption_Dataset.xlsx")):
        print("No dataset found; run without --no-generate first.")
        sys.exit(2)

    run("analyze_findings.py")
    run("build_analysis.py")
    run("build_qa_workbook.py")
    run("build_qa_doc.py")
    run("build_brd.py")

    with open(os.path.join(ROOT, "data", "findings.json")) as f:
        F = json.load(f)
    with open(os.path.join(ROOT, "data", "summary.json")) as f:
        S = json.load(f)

    print("\n" + "=" * 78)
    print(f"REFRESH COMPLETE — data as at {F['as_at']}")
    print("=" * 78)
    print(f"Consumption            PKR {F['consumption']:>15,.0f}")
    print(f"Procurement investment PKR {F['purchase']:>15,.0f}")
    print(f"Ratio                  {F['ratio']:>19.2f}")
    print(f"Budget utilisation     {F['utilisation']:>19.1%}")
    print(f"Unattributed           {F['unattributed_pct']:>18.1f}%")
    print(f"Expired stock          PKR {F['stock']['expired_value']:>15,.0f}")
    print(f"Open commitment        PKR {F['open_commitment']:>15,.0f}")
    print(f"Issue lines            {S['issue_lines']:>19,}")
    print("\nOutputs regenerated: analysis workbook, question bank workbook, guide, BRD.")
    print("Next: review Red items in the question bank workbook (sheet 08), assign owners, "
          "and record the snapshot for restatement.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
What-if / scenario engine for the VPT consumption analysis.

Everything is driven by config/parameters.json:
  * thresholds      – status bands used by the question bank
  * model           – the assumptions (demand, price, supply, policy, budget)
  * scenarios       – named lever combinations
  * sensitivity     – one-at-a-time ranges (tornado)
  * breakeven       – which lever is solved for, and the search range
  * twoway          – row/column levers for the two-way table

The engine is deterministic and side-effect free: given parameters and the
baseline facts it returns the full set of projections. Both the workbook and
the documents call it, so they cannot disagree.
"""
import json
import os
import copy

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "..", "config", "parameters.json")

DAYS = 365.0


def load_config(path=CONFIG):
    with open(path) as f:
        return json.load(f)


def default_params(cfg, findings):
    """Resolve every model parameter, substituting data-derived values for 'auto'."""
    p = {}
    for spec in cfg["model"]:
        v = spec["value"]
        if v == "auto":
            if spec["key"] == "activity_factor":
                v = 1.0   # next 12 months at the current operating level
            else:
                v = findings.get("baseline", {}).get(spec["key"], 0)
        p[spec["key"]] = v
    return p


def baseline_facts(findings):
    """Current, annualised operating facts used as the starting point of every scenario."""
    b = dict(findings["baseline"])
    b["activity_factor"] = 1.0
    return b


# --------------------------------------------------------------------------
# The model
# --------------------------------------------------------------------------
def run_model(params, base, items, cfg=None):
    """
    Project the next 12 months of consumption, purchasing, stock, budget and
    risk outcomes for one parameter set.

    Documented equations (also reproduced in the workbook's Model_Logic sheet):

      demand_factor      = 1 + demand_growth_pct
      price_factor       = (1 + price_inflation_pct) x (1 + price_shock_pct) x (1 - bulk_discount_pct)
      consumption_value  = run_rate_12 x demand_factor x price_factor x activity_factor
      lead_time_days     = avg_lead_days x lead_time_multiplier
      safety_days        = safety_stock_days + lead_time_days x (1 - supplier_reliability_pct)
      target_stock_value = consumption_value x (target_cover_days + safety_days) / 365
      available_stock    = on_hand_value + open_commitment
      top_up             = max(0, target_stock_value - available_stock)
      released           = max(0, available_stock - target_stock_value) x excess_stock_release_pct
      wastage_replacement= wastage_rate_pct x consumption_value
      purchase_value     = consumption_value + top_up + wastage_replacement - released
      budget_available   = annual_budget x (1 - funding_cut_pct) x (1 + budget_growth_pct)
      funding_gap        = max(0, consumption_value - budget_available)
      expiry_risk        = near_expiry_share x on_hand_value
                           + wastage_rate_pct x consumption_value
                           + 0.35 x excess_stock_value x (1 - redistribution_share_pct)
      stockout_items     = count(items whose projected cover < lead_time_days + safety_stock_days)
    """
    p = params
    b = base

    demand_factor = 1 + p["demand_growth_pct"]
    price_factor = (1 + p["price_inflation_pct"]) * (1 + p["price_shock_pct"]) * (1 - p["bulk_discount_pct"])
    activity = p.get("activity_factor", 1.0)

    consumption = b["run_rate_12"] * demand_factor * price_factor * activity

    lead_time_days = b["avg_lead_days"] * p["lead_time_multiplier"]
    safety_days = p["safety_stock_days"] + lead_time_days * (1 - p["supplier_reliability_pct"])
    target_stock = consumption * (p["target_cover_days"] + safety_days) / DAYS

    available = b["on_hand_value"] + b["open_commitment"]
    top_up = max(0.0, target_stock - available)
    excess = max(0.0, available - target_stock)
    released = excess * p.get("excess_stock_release_pct", 0.0)
    wastage_replacement = p["wastage_rate_pct"] * consumption
    purchase = max(0.0, consumption + top_up + wastage_replacement - released)

    budget_available = b["annual_budget"] * (1 - p["funding_cut_pct"]) * (1 + p["budget_growth_pct"])
    funding_gap = max(0.0, consumption - budget_available)

    # item-level stock-out and overstock analysis
    stockout_items = 0
    excess_value = 0.0
    overstock_items = 0
    risk_items = []
    for it in items:
        cover = it.get("cover_days")
        if cover is None:
            continue
        projected_cover = cover / demand_factor
        if projected_cover < lead_time_days + safety_days:
            stockout_items += 1
            if len(risk_items) < 8:
                risk_items.append({"id": it["id"], "name": it["name"], "cover": projected_cover})
        if projected_cover > p["max_cover_days"]:
            overstock_items += 1
            excess_value += it["onhand_value"] * (1 - p["max_cover_days"] / projected_cover)

    expiry_risk = (b["near_expiry_share"] * b["on_hand_value"]
                   + p["wastage_rate_pct"] * consumption
                   + 0.35 * excess_value * (1 - p["redistribution_share_pct"]))

    beneficiaries = b["annual_beneficiaries"] * (1 + p["service_volume_growth_pct"])

    return {
        "consumption_value": round(consumption, 2),
        "purchase_value": round(purchase, 2),
        "top_up": round(top_up, 2),
        "excess_stock_released": round(released, 2),
        "wastage_replacement": round(wastage_replacement, 2),
        "target_stock_value": round(target_stock, 2),
        "available_stock": round(available, 2),
        "ratio": round(consumption / purchase, 3) if purchase else None,
        "budget_available": round(budget_available, 2),
        "budget_utilisation": round(consumption / budget_available, 3) if budget_available else None,
        "funding_gap": round(funding_gap, 2),
        "expiry_risk_value": round(expiry_risk, 2),
        "overstock_value": round(excess_value, 2),
        "overstock_items": overstock_items,
        "stockout_items": stockout_items,
        "working_capital": round(target_stock, 2),
        "beneficiaries": round(beneficiaries),
        "cost_per_beneficiary": round(consumption / beneficiaries, 2) if beneficiaries else None,
        "lead_time_days": round(lead_time_days, 1),
        "safety_days": round(safety_days, 1),
        "demand_factor": round(demand_factor, 3),
        "price_factor": round(price_factor, 3),
        "risk_items": risk_items,
    }


# --------------------------------------------------------------------------
# Scenario helpers
# --------------------------------------------------------------------------
def apply_levers(params, levers):
    p = dict(params)
    p.update(levers)
    return p


def run_scenarios(cfg, findings, items):
    base_params = default_params(cfg, findings)
    facts = baseline_facts(findings)
    out = []
    baseline_result = None
    for sc in cfg["scenarios"]:
        p = apply_levers(base_params, sc["levers"])
        res = run_model(p, facts, items, cfg)
        changed = {k: {"baseline": base_params.get(k), "scenario": v}
                   for k, v in sc["levers"].items() if base_params.get(k) != v}
        if sc["id"] == "S0" or baseline_result is None:
            if sc["id"] == "S0":
                baseline_result = res
        out.append({
            "id": sc["id"], "name": sc["name"], "rationale": sc["rationale"],
            "levers": sc["levers"], "changed": changed, "result": res,
        })
    # deltas against the baseline scenario
    baseline_result = baseline_result or out[0]["result"]
    for row in out:
        row["delta"] = {
            k: (round(row["result"][k] - baseline_result[k], 2)
                if isinstance(row["result"].get(k), (int, float)) and isinstance(baseline_result.get(k), (int, float))
                else None)
            for k in ("consumption_value", "purchase_value", "funding_gap",
                      "expiry_risk_value", "stockout_items", "working_capital",
                      "cost_per_beneficiary", "overstock_value")
        }
    return base_params, facts, baseline_result, out


def run_sensitivity(cfg, findings, items):
    base_params = default_params(cfg, findings)
    facts = baseline_facts(findings)
    base_res = run_model(base_params, facts, items, cfg)
    spec = cfg["sensitivity"]
    rows = []
    for lever in spec["levers"]:
        key = lever["key"]
        low_p = apply_levers(base_params, {key: lever["low"]})
        high_p = apply_levers(base_params, {key: lever["high"]})
        low_r = run_model(low_p, facts, items, cfg)
        high_r = run_model(high_p, facts, items, cfg)
        row = {"key": key, "low": lever["low"], "high": lever["high"], "outputs": {}}
        for out in spec["outputs"]:
            row["outputs"][out] = {
                "base": base_res[out], "low": low_r[out], "high": high_r[out],
                "swing": round(abs(high_r[out] - low_r[out]), 2),
            }
        rows.append(row)
    order = {out: sorted(rows, key=lambda r: -r["outputs"][out]["swing"]) for out in spec["outputs"]}
    return base_res, rows, order


def solve_breakeven(cfg, findings, items, key, target):
    """Bisection on one lever: the value at which the target condition flips."""
    base_params = default_params(cfg, findings)
    facts = baseline_facts(findings)

    def f(x):
        p = apply_levers(base_params, {key: x})
        r = run_model(p, facts, items, cfg)
        if target == "funding_gap == 0":
            return r["funding_gap"]
        if target == "purchase_value minimised":
            return r["purchase_value"]
        return r.get(target, 0)

    spec = next((s for s in cfg["breakeven"]["solves"] if s["key"] == key), None)
    lo = spec["min"] if spec else -1
    hi = spec["max"] if spec else 1
    f_lo, f_hi = f(lo), f(hi)
    out = {"key": key, "target": target, "range": [lo, hi],
           "value_at_low": round(f_lo, 2), "value_at_high": round(f_hi, 2), "breakeven": None, "note": ""}
    if target == "curve":
        # trade-off curve: what each cover policy costs and what risk it carries
        grid = [lo + (hi - lo) * i / 30 for i in range(31)]
        pts = []
        for x in grid:
            r = run_model(apply_levers(base_params, {key: x}), facts, items, cfg)
            pts.append({"x": round(x, 1), "purchase": r["purchase_value"],
                        "expiry": r["expiry_risk_value"], "capital": r["working_capital"],
                        "stockouts": r["stockout_items"]})
        out["breakeven"] = None
        out["note"] = "Trade-off curve: purchase cost, expiry risk and working capital by cover policy"
        out["curve"] = pts
        return out
    if f_lo <= 0 and f_hi <= 0:
        out["note"] = "No breach anywhere in the search range"
        return out
    if f_lo > 0 and f_hi > 0:
        out["note"] = "Already breached at the lower bound: the break-even value is below the search range"
        return out
    if (f_lo > 0) == (f_hi > 0):
        out["note"] = "Condition does not flip inside the search range"
        return out
    for _ in range(60):
        mid = (lo + hi) / 2
        if (f(mid) > 0) == (f_lo > 0):
            lo = mid
            f_lo = f(mid)
        else:
            hi = mid
    out["breakeven"] = round((lo + hi) / 2, 4)
    out["note"] = "Solved by bisection to 0.0001"
    return out


def two_way(cfg, findings, items):
    base_params = default_params(cfg, findings)
    facts = baseline_facts(findings)
    spec = cfg["twoway"]
    tables = {out: [] for out in spec["outputs"]}
    for rv in spec["row_values"]:
        row_vals = {out: [] for out in spec["outputs"]}
        for cv in spec["col_values"]:
            p = apply_levers(base_params, {spec["row_key"]: rv, spec["col_key"]: cv})
            res = run_model(p, facts, items, cfg)
            for out in spec["outputs"]:
                row_vals[out].append(res[out])
        for out in spec["outputs"]:
            tables[out].append(row_vals[out])
    return {
        "row_key": spec["row_key"], "row_values": spec["row_values"],
        "col_key": spec["col_key"], "col_values": spec["col_values"],
        "tables": tables,
    }


def build_all(cfg=None):
    cfg = cfg or load_config()
    root = os.path.join(BASE, "..", "data")
    with open(os.path.join(root, "findings.json")) as f:
        findings = json.load(f)
    items = findings["item_level"]

    base_params, facts, baseline, scenarios = run_scenarios(cfg, findings, items)
    base_res, sens_rows, sens_order = run_sensitivity(cfg, findings, items)

    be = []
    for spec in cfg["breakeven"]["solves"]:
        be.append(solve_breakeven(cfg, findings, items, spec["key"], spec["target"]))

    tw = two_way(cfg, findings, items)

    return {
        "config_version": cfg["version"],
        "as_at": findings["as_at"],
        "baseline_params": base_params,
        "baseline_facts": facts,
        "baseline_result": base_res,
        "scenarios": scenarios,
        "sensitivity": {"rows": sens_rows, "order": {k: [r["key"] for r in v] for k, v in sens_order.items()},
                        "outputs": cfg["sensitivity"]["outputs"]},
        "breakeven": be,
        "twoway": tw,
        "lever_spec": cfg["model"],
        "scenario_defs": cfg["scenarios"],
    }


if __name__ == "__main__":
    res = build_all()
    out = os.path.join(BASE, "..", "data", "whatif_results.json")
    with open(out, "w") as f:
        json.dump(res, f, indent=2, default=str)
    print("saved", out)
    b = res["baseline_result"]
    print(f"Baseline: consumption {b['consumption_value']:,.0f} purchase {b['purchase_value']:,.0f} "
          f"gap {b['funding_gap']:,.0f} util {b['budget_utilisation']:.1%} "
          f"expiry {b['expiry_risk_value']:,.0f} stockouts {b['stockout_items']}")
    for s in res["scenarios"]:
        r = s["result"]
        print(f"  {s['id']} {s['name'][:46]:46s} gap {r['funding_gap']:>12,.0f} "
              f"purchase {r['purchase_value']:>12,.0f} stockouts {r['stockout_items']:>3} "
              f"expiry {r['expiry_risk_value']:>10,.0f}")

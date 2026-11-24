"""Single source of truth for the DMart EBITDA case.

Reads the sourced filings in data/, converts them to INR crore and computes every number
the model, deck, site and tests use. Run: python scripts/case.py  (writes build/case.json)
"""
from __future__ import annotations

import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
BUILD = ROOT / "build"

TO_CRORE = {"INR": 1e-7, "INR lakh": 1e-2}

# Improvement levers, applied to the FY26 base. These are the case's ASSUMPTIONS, not DMart data.
SCENARIOS = {
    "Cautious": {"emp_bps": 0, "gm_bps": 0, "density_recovery": 0.25, "variable_opex_share": 0.5},
    "Base": {"emp_bps": 10, "gm_bps": 10, "density_recovery": 0.50, "variable_opex_share": 0.5},
    "Stretch": {"emp_bps": 20, "gm_bps": 20, "density_recovery": 0.75, "variable_opex_share": 0.5},
}
SENSITIVITY_BPS = [-30, -20, -10, 0, 10, 20, 30]


def load_financials() -> dict[str, dict[str, float]]:
    out: dict[str, dict[str, float]] = {"FY25": {}, "FY26": {}}
    with open(DATA / "financials.csv", newline="") as f:
        for row in csv.DictReader(f):
            out["FY25"][row["line"]] = float(row["fy25_reported"]) * TO_CRORE[row["fy25_unit"]]
            out["FY26"][row["line"]] = float(row["fy26_reported"]) * TO_CRORE[row["fy26_unit"]]
    return out


def load_operating() -> dict[str, dict[str, float | None]]:
    out: dict[str, dict[str, float | None]] = {}
    with open(DATA / "operating.csv", newline="") as f:
        for row in csv.DictReader(f):
            out[row["metric"]] = {
                "FY25": float(row["fy25"]) if row["fy25"] else None,
                "FY26": float(row["fy26"]) if row["fy26"] else None,
            }
    return out


def pnl(year: dict[str, float]) -> dict[str, float]:
    rev = year["revenue_from_operations"]
    cogs = year["purchases_of_stock_in_trade"] + year["changes_in_inventories"]
    gross = rev - cogs
    emp = year["employee_benefit_expense"]
    oth = year["other_expenses"]
    ebitda = gross - emp - oth
    pbt = ebitda + year["other_income"] - year["depreciation_and_amortisation"] - year["finance_costs"]
    pat = pbt - year["total_tax_expense"]
    return {
        "revenue": rev, "cogs": cogs, "gross_margin": gross, "employee": emp, "other_expenses": oth,
        "ebitda": ebitda, "other_income": year["other_income"],
        "depreciation": year["depreciation_and_amortisation"], "finance_costs": year["finance_costs"],
        "pbt": pbt, "tax": year["total_tax_expense"], "pat": pat,
        "reported_pbt": year["profit_before_tax"], "reported_pat": year["profit_after_tax"],
        "gross_margin_pct": gross / rev, "employee_pct": emp / rev, "other_expenses_pct": oth / rev,
        "ebitda_pct": ebitda / rev, "pat_pct": pat / rev,
    }


def bridge(p25: dict[str, float], p26: dict[str, float]) -> list[dict[str, float | str]]:
    """FY25 -> FY26 EBITDA bridge. Sums exactly: EBITDA% = GM% - employee% - other%."""
    r26 = p26["revenue"]
    steps = [
        ("FY25 EBITDA", p25["ebitda"]),
        ("Revenue growth at FY25 margin", (r26 - p25["revenue"]) * p25["ebitda_pct"]),
        ("Gross margin rate", (p26["gross_margin_pct"] - p25["gross_margin_pct"]) * r26),
        ("Employee cost rate", -(p26["employee_pct"] - p25["employee_pct"]) * r26),
        ("Other expenses rate", -(p26["other_expenses_pct"] - p25["other_expenses_pct"]) * r26),
        ("FY26 EBITDA", p26["ebitda"]),
    ]
    return [{"step": s, "value": v} for s, v in steps]


def levers(p26: dict[str, float], density: dict[str, float], area26_msqft: float, a: dict[str, float]) -> dict[str, float]:
    r26 = p26["revenue"]
    l1 = a["emp_bps"] / 10_000 * r26
    l2 = a["gm_bps"] / 10_000 * r26
    gap_rs_per_sqft = density["FY25"] - density["FY26"]
    extra_revenue = a["density_recovery"] * gap_rs_per_sqft * area26_msqft * 1e6 * 1e-7
    flow_through = p26["gross_margin_pct"] - a["variable_opex_share"] * (p26["employee_pct"] + p26["other_expenses_pct"])
    l3 = extra_revenue * flow_through
    total = l1 + l2 + l3
    pro_rev = r26 + extra_revenue
    pro_ebitda = p26["ebitda"] + total
    return {
        "employee_productivity": l1, "gross_margin_mix": l2, "new_store_ramp": l3,
        "new_store_extra_revenue": extra_revenue, "flow_through_pct": flow_through,
        "total": total, "proforma_revenue": pro_rev, "proforma_ebitda": pro_ebitda,
        "proforma_margin": pro_ebitda / pro_rev,
        "margin_uplift_bps": (pro_ebitda / pro_rev - p26["ebitda_pct"]) * 10_000,
    }


def compute() -> dict:
    fin = load_financials()
    ops = load_operating()
    p = {y: pnl(fin[y]) for y in ("FY25", "FY26")}
    area = {y: ops["retail_business_area_at_year_end"][y] for y in ("FY25", "FY26")}
    stores = {y: ops["stores_at_year_end"][y] for y in ("FY25", "FY26")}
    density = {y: p[y]["revenue"] * 1e7 / (area[y] * 1e6) for y in ("FY25", "FY26")}
    growth = {k: p["FY26"][k] / p["FY25"][k] - 1 for k in
              ("revenue", "gross_margin", "employee", "other_expenses", "ebitda", "depreciation", "finance_costs", "pat")}
    case = {
        "pnl": p,
        "growth": growth,
        "bridge": bridge(p["FY25"], p["FY26"]),
        "operations": {
            "stores": stores, "area_msqft": area,
            "stores_added_fy26": ops["stores_added_in_year"]["FY26"],
            "stores_added_q4fy26": ops["stores_added_in_q4"]["FY26"],
            "revenue_per_yearend_sqft": density,
            "density_change_pct": density["FY26"] / density["FY25"] - 1,
            "revenue_per_store_cr": {y: p[y]["revenue"] / stores[y] for y in stores},
            "employee_cost_per_store_cr": {y: p[y]["employee"] / stores[y] for y in stores},
            "area_growth_pct": area["FY26"] / area["FY25"] - 1,
        },
        "reported": {
            "ebitda_fy26": ops["reported_ebitda"]["FY26"],
            "ebitda_growth_pct": ops["reported_ebitda_growth"]["FY26"],
            "ebitda_margin_pct": {"FY25": ops["reported_ebitda_margin"]["FY25"], "FY26": ops["reported_ebitda_margin"]["FY26"]},
        },
        "scenarios": {name: {"assumptions": a, **levers(p["FY26"], density, area["FY26"], a)} for name, a in SCENARIOS.items()},
        "sensitivity": {
            "bps": SENSITIVITY_BPS,
            "margin": [[p["FY26"]["ebitda_pct"] + gm / 1e4 - emp / 1e4 for emp in SENSITIVITY_BPS] for gm in SENSITIVITY_BPS],
        },
    }
    return case


def main() -> None:
    BUILD.mkdir(exist_ok=True)
    case = compute()
    (BUILD / "case.json").write_text(json.dumps(case, indent=2))
    p = case["pnl"]
    print(f"EBITDA FY25 {p['FY25']['ebitda']:.2f} cr ({p['FY25']['ebitda_pct']:.2%}), FY26 {p['FY26']['ebitda']:.2f} cr ({p['FY26']['ebitda_pct']:.2%})")
    for b in case["bridge"]:
        print(f"  {b['step']:<32}{b['value']:>10.1f}")
    for n, s in case["scenarios"].items():
        print(f"  {n:<9} levers {s['total']:7.1f} cr  pro-forma margin {s['proforma_margin']:.2%} (+{s['margin_uplift_bps']:.0f} bps)")


if __name__ == "__main__":
    main()

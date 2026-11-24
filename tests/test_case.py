"""Checks that the case reconciles to DMart's filings and that every output shows the same numbers."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))

import build_model  # noqa: E402
import case as case_mod  # noqa: E402

MODEL = ROOT / "model" / "DMart_EBITDA_Case.xlsx"
DECK = ROOT / "deck" / "DMart_EBITDA_Case.pptx"
SITE = ROOT / "site" / "index.html"


@pytest.fixture(scope="module")
def c():
    return case_mod.compute()


# ---------- the numbers reconcile to what DMart reported ----------

def test_ebitda_matches_reported_figures(c):
    assert round(c["pnl"]["FY26"]["ebitda"]) == 5255  # press release: Rs 5,255 crore
    assert round(c["growth"]["ebitda"] * 100, 1) == 15.7  # press release: +15.7%
    assert round(c["pnl"]["FY25"]["ebitda_pct"] * 100, 1) == 7.9
    assert round(c["pnl"]["FY26"]["ebitda_pct"] * 100, 1) == 7.8


def test_revenue_matches_reported_figures(c):
    assert round(c["pnl"]["FY26"]["revenue"], 2) == 66968.03
    assert round(c["pnl"]["FY25"]["revenue"], 2) == 57789.81
    assert round(c["growth"]["revenue"] * 100, 1) == 15.9


@pytest.mark.parametrize("year", ["FY25", "FY26"])
def test_profit_before_and_after_tax_rebuild_exactly(c, year):
    p = c["pnl"][year]
    assert p["pbt"] == pytest.approx(p["reported_pbt"], abs=0.01)
    assert p["pat"] == pytest.approx(p["reported_pat"], abs=0.01)


def test_unit_conversion():
    fin = case_mod.load_financials()
    assert fin["FY26"]["revenue_from_operations"] == pytest.approx(66968.03)
    assert fin["FY25"]["revenue_from_operations"] == pytest.approx(57789.81)


# ---------- the analysis is internally consistent ----------

def test_bridge_adds_up_to_fy26_ebitda(c):
    steps = [b["value"] for b in c["bridge"]]
    assert sum(steps[:-1]) == pytest.approx(steps[-1], abs=1e-6)


def test_bridge_story(c):
    by = {b["step"]: b["value"] for b in c["bridge"]}
    assert by["Employee cost rate"] < 0 < by["Gross margin rate"]
    assert by["Other expenses rate"] > 0
    assert round(-by["Employee cost rate"]) == 178


def test_margin_shares_add_up(c):
    for year in ("FY25", "FY26"):
        p = c["pnl"][year]
        assert p["gross_margin_pct"] - p["employee_pct"] - p["other_expenses_pct"] == pytest.approx(p["ebitda_pct"])


def test_levers(c):
    s, rev = c["scenarios"], c["pnl"]["FY26"]["revenue"]
    assert s["Base"]["employee_productivity"] == pytest.approx(10 / 10_000 * rev)
    assert s["Base"]["gross_margin_mix"] == pytest.approx(10 / 10_000 * rev)
    assert s["Cautious"]["employee_productivity"] == 0
    for v in s.values():
        assert v["total"] == pytest.approx(v["employee_productivity"] + v["gross_margin_mix"] + v["new_store_ramp"])
    assert s["Cautious"]["total"] < s["Base"]["total"] < s["Stretch"]["total"]
    assert round(s["Base"]["total"]) == 258
    assert round(s["Base"]["margin_uplift_bps"]) == 25


def test_density_fell(c):
    assert round(c["operations"]["density_change_pct"] * 100, 1) == -3.2


def test_sensitivity_grid(c):
    grid, bps = c["sensitivity"]["margin"], c["sensitivity"]["bps"]
    mid = bps.index(0)
    assert grid[mid][mid] == pytest.approx(c["pnl"]["FY26"]["ebitda_pct"])
    assert grid[-1][0] > grid[mid][mid] > grid[0][-1]


# ---------- the Excel model computes the same numbers ----------

@pytest.fixture(scope="module")
def xl(tmp_path_factory):
    from pycel import ExcelCompiler

    out = tmp_path_factory.mktemp("model") / "fresh.xlsx"
    original = build_model.OUT
    build_model.OUT = out
    try:
        build_model.build()
    finally:
        build_model.OUT = original
    return ExcelCompiler(filename=str(out)), out


@pytest.mark.parametrize("ref,path", [
    ("P&L!C10", ("pnl", "FY26", "ebitda")),
    ("P&L!B10", ("pnl", "FY25", "ebitda")),
    ("P&L!C16", ("pnl", "FY26", "pat")),
    ("P&L!C23", ("pnl", "FY26", "ebitda_pct")),
    ("Levers!D23", ("scenarios", "Base", "total")),
    ("Levers!C26", ("scenarios", "Cautious", "proforma_margin")),
    ("Levers!E27", ("scenarios", "Stretch", "margin_uplift_bps")),
    ("Operations!C10", ("operations", "revenue_per_yearend_sqft", "FY26")),
])
def test_model_formulas_match_python(xl, c, ref, path):
    compiler, _ = xl
    expected = c
    for k in path:
        expected = expected[k]
    assert compiler.evaluate(ref) == pytest.approx(expected, rel=1e-9)


def test_model_bridge_matches_python(xl, c):
    compiler, _ = xl
    for i, step in enumerate(c["bridge"]):
        assert compiler.evaluate(f"Bridge!B{5 + i}") == pytest.approx(step["value"], rel=1e-9)


def test_model_checks_all_pass(xl):
    compiler, _ = xl
    assert compiler.evaluate("Checks!C15") is True


def test_committed_model_is_the_generator_output(xl):
    _, fresh = xl
    a, b = load_workbook(MODEL), load_workbook(fresh)
    assert a.sheetnames == b.sheetnames
    for name in a.sheetnames:
        for row_a, row_b in zip(a[name].iter_rows(), b[name].iter_rows()):
            for ca, cb in zip(row_a, row_b):
                assert ca.value == cb.value, f"{name}!{ca.coordinate} differs: rebuild with scripts/build_model.py"


def test_committed_model_has_current_cached_values(c):
    wb = load_workbook(MODEL, data_only=True)
    assert wb["P&L"]["C10"].value == pytest.approx(c["pnl"]["FY26"]["ebitda"])
    assert wb["Levers"]["D23"].value == pytest.approx(c["scenarios"]["Base"]["total"])
    assert wb["Checks"]["C15"].value is True


# ---------- the deck and site show the same numbers ----------

def deck_text() -> str:
    from pptx import Presentation

    prs = Presentation(DECK)
    parts = []
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                parts.append(shape.text_frame.text)
            if shape.has_table:
                parts += [cell.text for row in shape.table.rows for cell in row.cells]
    return "\n".join(parts), len(prs.slides)


def test_deck_shows_current_numbers(c):
    text, slides = deck_text()
    assert slides == 10
    for s in ["₹66,968 cr", "₹5,255 cr", "7.85%", "+27 bps", "₹258 cr", "+25 bps", "+47 bps", "58 of 85"]:
        assert s in text, s


def test_site_shows_current_numbers():
    page = SITE.read_text()
    for s in ["₹66,968 cr", "₹5,255 cr", "7.86% → 7.85%", "+27 bps", "₹258", "8.10%", "58 of 85"]:
        assert s in page, s


@pytest.mark.parametrize("src", ["model/DMart_EBITDA_Case.xlsx", "deck/DMart_EBITDA_Case.pptx", "deck/DMart_EBITDA_Case.pdf"])
def test_site_downloads_are_the_latest_files(src):
    assert (ROOT / src).read_bytes() == (ROOT / "site" / Path(src).name).read_bytes()

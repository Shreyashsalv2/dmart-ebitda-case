"""Build the static case page (site/index.html) and copy the downloads next to it.

Run after scripts/case.py, scripts/build_model.py and the deck build:  python scripts/build_site.py
"""
from __future__ import annotations

import html
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from case import ROOT, compute  # noqa: E402

SITE = ROOT / "site"
DOWNLOADS = {
    "model/DMart_EBITDA_Case.xlsx": "DMart_EBITDA_Case.xlsx",
    "deck/DMart_EBITDA_Case.pptx": "DMart_EBITDA_Case.pptx",
    "deck/DMart_EBITDA_Case.pdf": "DMart_EBITDA_Case.pdf",
}
REPO = "https://github.com/Shreyashsalv2/dmart-ebitda-case"


def n0(x: float) -> str:
    """Indian digit grouping, no decimals (66968 -> 66,968; 112345 -> 1,12,345)."""
    neg, s = x < 0, str(int(round(abs(x))))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        s = ",".join(parts + [tail])
    return ("−" if neg else "") + s


def pct(x: float, d: int = 1) -> str:
    return ("−" if x < 0 else "") + f"{abs(x) * 100:.{d}f}%"


def gpct(x: float, d: int = 1) -> str:
    return ("+" if x >= 0 else "−") + f"{abs(x) * 100:.{d}f}%"


def bps(x: float) -> str:
    return ("+" if x >= 0 else "−") + f"{abs(round(x))} bps"


def bridge_svg(bridge: list[dict]) -> str:
    w, h, left, right, top, bottom = 760, 330, 20, 20, 30, 260
    lo_axis, hi_axis = 4000, 5500
    bw = (w - left - right) / len(bridge) * 0.62
    step = (w - left - right) / len(bridge)
    y = lambda v: bottom - (v - lo_axis) / (hi_axis - lo_axis) * (bottom - top)  # noqa: E731
    names = ["FY25 EBITDA", "Revenue growth", "Gross margin rate", "Employee cost rate", "Other expenses rate", "FY26 EBITDA"]
    out = [f'<svg viewBox="0 0 {w} {h}" role="img" aria-label="EBITDA bridge from FY25 to FY26 in INR crore">']
    running = 0.0
    for i, b in enumerate(bridge):
        x = left + i * step + (step - bw) / 2
        total = i in (0, len(bridge) - 1)
        if total:
            lo, hi, cls, running = lo_axis, b["value"], "t", b["value"]
            label = n0(b["value"])
        else:
            before = running
            running += b["value"]
            lo, hi = min(before, running), max(before, running)
            cls = "u" if b["value"] >= 0 else "d"
            label = ("+" if b["value"] >= 0 else "−") + n0(abs(b["value"]))
        out.append(f'<rect class="{cls}" x="{x:.1f}" y="{y(hi):.1f}" width="{bw:.1f}" height="{max(y(lo) - y(hi), 2):.1f}"/>')
        out.append(f'<text class="v {cls}" x="{x + bw / 2:.1f}" y="{y(hi) - 8:.1f}">{label}</text>')
        out.append(f'<text class="c" x="{x + bw / 2:.1f}" y="{bottom + 22}">{html.escape(names[i])}</text>')
    out.append(f'<line x1="{left}" x2="{w - right}" y1="{bottom}" y2="{bottom}" class="ax"/>')
    out.append(f'<text class="note" x="{left}" y="{h - 8}">INR crore · axis starts at 4,000</text></svg>')
    return "".join(out)


def build() -> Path:
    c = compute()
    p25, p26, g, o = c["pnl"]["FY25"], c["pnl"]["FY26"], c["growth"], c["operations"]
    s = c["scenarios"]
    base = s["Base"]
    d = {k: (p26[k] - p25[k]) * 1e4 for k in ("gross_margin_pct", "employee_pct", "other_expenses_pct", "ebitda_pct")}
    rows = "".join(
        f"<tr><td>{label}</td><td>{pct(p25[k], 2)}</td><td>{pct(p26[k], 2)}</td><td class='{cls}'>{bps(d[k])}</td></tr>"
        for label, k, cls in [("Gross margin", "gross_margin_pct", "good"), ("Employee cost", "employee_pct", "bad"),
                              ("Other expenses", "other_expenses_pct", "good"), ("EBITDA margin", "ebitda_pct", "")]
    )
    scen = "".join(
        f"<tr><td>{name}</td><td>{a['assumptions']['emp_bps']} bps</td><td>{a['assumptions']['gm_bps']} bps</td>"
        f"<td>{a['assumptions']['density_recovery']:.0%}</td><td><b>₹{n0(a['total'])} cr</b></td>"
        f"<td><b>{pct(a['proforma_margin'], 2)}</b></td><td>{bps(a['margin_uplift_bps'])}</td></tr>"
        for name, a in s.items()
    )
    levers = [
        ("Staff to the ramp", f"Win back 10 of the {round(d['employee_pct'])} bps employee-cost rise by staffing new stores to a sales ramp curve and sharing staff across nearby stores.", base["employee_productivity"]),
        ("Shift the mix", "Add 10 bps of gross margin by growing general merchandise and apparel, without raising everyday prices.", base["gross_margin_mix"]),
        ("Ramp new stores faster", f"Close half the sales-density gap: ₹{n0(base['new_store_extra_revenue'])} cr extra revenue at {pct(base['flow_through_pct'])} flow-through.", base["new_store_ramp"]),
    ]
    lever_html = "".join(f"<div class='card'><span class='num'>{i}</span><h3>{h}</h3><p>{t}</p><p class='big'>+₹{n0(v)} cr</p></div>"
                         for i, (h, t, v) in enumerate(levers, start=1))
    page = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>DMart EBITDA improvement case · Shreyash Singh Salv</title>
<meta name="description" content="Why DMart's EBITDA margin stayed flat in FY26 while revenue grew 15.9%, and three levers sized at about ₹{n0(base['total'])} crore. Excel model and deck from public filings.">
<style>
:root{{--g:#1B4332;--g2:#2D6A4F;--l:#EEF5F0;--r:#C0392B;--t:#1B1B1B;--m:#5c6670}}
*{{box-sizing:border-box}}body{{margin:0;font-family:Calibri,"Segoe UI",system-ui,sans-serif;color:var(--t);line-height:1.5;background:#fff}}
header{{background:var(--g);color:#fff;padding:56px 20px 44px}}.w{{max-width:980px;margin:0 auto;padding:0 20px}}
h1{{font-family:Cambria,Georgia,serif;font-size:clamp(28px,4.4vw,44px);line-height:1.15;margin:0 0 12px}}h2{{font-family:Cambria,Georgia,serif;color:var(--g);font-size:26px;margin:0 0 14px}}
header p{{color:#B7E4C7;font-size:18px;margin:0 0 22px;max-width:760px}}.btns a{{display:inline-block;margin:6px 10px 0 0;padding:10px 16px;border-radius:8px;background:#fff;color:var(--g);font-weight:700;text-decoration:none}}
.btns a.alt{{background:transparent;color:#fff;border:1px solid #95D5B2}}section{{padding:40px 0;border-bottom:1px solid #e4ebe6}}
.stats{{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px}}.stat,.card{{background:var(--l);border-radius:12px;padding:18px}}
.stat b{{display:block;font-size:28px;color:var(--g)}}.stat span{{color:var(--g2);font-weight:700}}.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}}
.card h3{{margin:6px 0 6px;color:var(--g)}}.card .big{{font-size:26px;font-weight:700;color:var(--g2);margin:8px 0 0}}.num{{display:inline-flex;width:30px;height:30px;border-radius:50%;background:var(--g2);color:#fff;align-items:center;justify-content:center;font-weight:700}}
table{{width:100%;border-collapse:collapse;font-size:16px}}th{{background:var(--g);color:#fff;text-align:left;padding:8px 10px}}td{{padding:8px 10px;border-bottom:1px solid #dfe7e2}}
.good{{color:var(--g2);font-weight:700}}.bad{{color:var(--r);font-weight:700}}.scroll{{overflow-x:auto}}
svg{{width:100%;height:auto;font-family:Calibri,system-ui,sans-serif}}svg .t{{fill:var(--g)}}svg .u{{fill:var(--g2)}}svg .d{{fill:var(--r)}}
svg text{{text-anchor:middle;font-size:13px;fill:var(--t)}}svg text.v{{font-weight:700;font-size:14px}}svg text.v.t,svg text.v.u{{fill:var(--g)}}svg text.v.d{{fill:var(--r)}}svg .ax{{stroke:#8d99ae}}svg .note{{text-anchor:start;font-size:12px;fill:var(--m);font-style:italic}}
.muted{{color:var(--m);font-size:14px}}footer{{padding:30px 0 50px;color:var(--m);font-size:14px}}a{{color:var(--g2)}}
ul{{padding-left:20px}}li{{margin:4px 0}}.two{{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:24px}}
</style></head><body>
<header><div class="w">
<h1>DMart grew {pct(g['revenue'])} in FY26. Its EBITDA margin did not move.</h1>
<p>An EBITDA improvement case built only from DMart's public filings: what held the margin at {pct(p26['ebitda_pct'], 2)}, and three levers worth about ₹{n0(base['total'])} crore in the base case.</p>
<div class="btns"><a href="DMart_EBITDA_Case.xlsx" download>Excel model ↓</a><a href="DMart_EBITDA_Case.pdf" target="_blank" rel="noopener">Deck (PDF) ↗</a><a href="DMart_EBITDA_Case.pptx" download>Deck (PowerPoint) ↓</a><a class="alt" href="{REPO}" target="_blank" rel="noopener">GitHub ↗</a></div>
</div></header>
<main class="w">
<section><h2>FY26 at a glance (standalone)</h2><div class="stats">
<div class="stat"><b>₹{n0(p26['revenue'])} cr</b>Revenue from operations <span>{gpct(g['revenue'])}</span></div>
<div class="stat"><b>₹{n0(p26['ebitda'])} cr</b>EBITDA <span>{gpct(g['ebitda'])}</span></div>
<div class="stat"><b>{pct(p25['ebitda_pct'], 2)} → {pct(p26['ebitda_pct'], 2)}</b>EBITDA margin <span>{bps(d['ebitda_pct'])}</span></div>
<div class="stat"><b>{int(o['stores']['FY25'])} → {int(o['stores']['FY26'])}</b>Stores at year end <span>{int(o['stores_added_fy26'])} added, {int(o['stores_added_q4fy26'])} in Q4</span></div>
</div></section>
<section><h2>A rising employee-cost share cost ₹{n0(-c['bridge'][3]['value'])} crore</h2>
<p>Gross margin and other expenses both improved as a share of revenue. Employee cost grew {pct(g['employee'])} against revenue growth of {pct(g['revenue'])}, which cancelled those gains.</p>
{bridge_svg(c['bridge'])}
<div class="scroll"><table><tr><th>% of revenue</th><th>FY25</th><th>FY26</th><th>Change</th></tr>{rows}</table></div>
<p class="muted">For cost lines, a positive change means a higher share of revenue. Revenue growth in the bridge is valued at the FY25 margin; each other bar is the change in that line's share of revenue × FY26 revenue.</p></section>
<section><h2>Hypothesis: new stores were staffed before they reached full sales</h2>
<p>{int(o['stores_added_q4fy26'])} of {int(o['stores_added_fy26'])} new stores opened in Q4. Revenue per square foot of year-end space fell {pct(-o['density_change_pct'])} (₹{n0(o['revenue_per_yearend_sqft']['FY25'])} → ₹{n0(o['revenue_per_yearend_sqft']['FY26'])}), and employee cost per store rose from ₹{o['employee_cost_per_store_cr']['FY25']:.2f} crore to ₹{o['employee_cost_per_store_cr']['FY26']:.2f} crore. Testing this needs store-level data by opening date, which only DMart has.</p></section>
<section><h2>Three levers, about ₹{n0(base['total'])} crore in the base case</h2><div class="cards">{lever_html}</div>
<h2 style="margin-top:30px">Scenarios (FY26 pro-forma)</h2><div class="scroll"><table><tr><th>Case</th><th>Employee cost recovered</th><th>Gross margin gain</th><th>New-store gap closed</th><th>Lever impact</th><th>EBITDA margin</th><th>Uplift</th></tr>{scen}</table></div>
<p class="muted">Lever sizes are the author's assumptions, applied to the FY26 base. They are not forecasts or DMart guidance. Change them in the yellow cells of the Excel model.</p></section>
<section class="two"><div><h2>Risks</h2><ul><li>Mix changes must not weaken DMart's low-price image</li><li>Leaner staffing can hurt service, shrinkage and availability</li><li>New-store ramp depends on location</li><li>Assumptions are untested against DMart's internal data</li></ul></div>
<div><h2>Next steps</h2><ul><li>Get store-level sales and staff hours by opening month</li><li>Build a staffing-to-sales curve for new stores</li><li>Check gross margin by category and region</li><li>Pilot the staffing model in one city cluster</li></ul></div></section>
<section><h2>Sources and checks</h2><p>All figures are DMart's own, from its audited standalone results filed with NSE for FY25 and FY26, its FY26 results press release, a Q4 FY25 company update and its FY2024-25 annual report. The model reproduces DMart's reported EBITDA of ₹{n0(c['reported']['ebitda_fy26'])} crore and both years' reported profit before and after tax exactly. Full links: <a href="{REPO}/blob/main/data/SOURCES.md" target="_blank" rel="noopener">data/SOURCES.md</a>.</p></section>
</main>
<footer class="w">Independent student analysis by Shreyash Singh Salv (<a href="https://github.com/Shreyashsalv2" target="_blank" rel="noopener">github.com/Shreyashsalv2</a>). Not investment advice; not affiliated with or reviewed by DMart.</footer>
</body></html>
"""
    SITE.mkdir(exist_ok=True)
    (SITE / "index.html").write_text(page)
    for src, name in DOWNLOADS.items():
        shutil.copy2(ROOT / src, SITE / name)
    return SITE / "index.html"


if __name__ == "__main__":
    print(build())

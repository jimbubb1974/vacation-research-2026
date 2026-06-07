#!/usr/bin/env python3
"""
build_costs.py — generates trip_costs.xlsx and updates the Cost Estimate
section in itinerary.md and itinerary.html.

Run: python build_costs.py
Requires: openpyxl  (pip install openpyxl)

To update a cost: edit the COSTS list below and re-run.
Totals are computed by Python — not the LLM, not Excel formulas.
The spreadsheet uses SUM() formulas as a double-check.
"""

import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).parent

# ── Cost line items ────────────────────────────────────────────────────────────
# (category, description, USD_amount, is_estimate)
#   is_estimate=False  confirmed/booked price
#   is_estimate=True   placeholder — research and update before trip

COSTS = [
    # FLIGHTS
    ("Flights",
     "UA 108 IAD→MUC Jul 11 + OS135/UA53 VIE→IAD Jul 21 (4 people, Standard Economy)",
     5422.52, False),

    # ACCOMMODATION
    ("Accommodation", "Munich — Holiday Inn Westpark, 2 nights (Jul 12–14)",   677.36, False),
    ("Accommodation", "Königssee — Villa Alpenrausch, 3 nights (Jul 14–17)", 1500.00, True),
    ("Accommodation", "Salzburg — Numa Vogelweider, 2 nights (Jul 17–19)",     850.66, False),
    ("Accommodation", "Vienna — Hotel Marc Aurel, 2 nights (Jul 19–21)",       609.00, False),

    # TRANSPORT
    ("Transport", "Airport → city: MVV group day ticket",                       16.00, True),
    ("Transport", "Bayern-Ticket: Munich → Berchtesgaden (Jul 14, group)",      33.00, True),
    ("Transport", "ÖBB Railjet: Salzburg → Vienna (Jul 18, 4 people)",         110.00, True),
    ("Transport", "Vienna public transit",                                       22.00, True),

    # ACTIVITIES
    ("Activities", "Mike's Bike Tours Munich (4 people)",                       160.00, True),
    ("Activities", "Nymphenburg Palace admission (4 people)",                    44.00, True),
    ("Activities", "Eagle's Nest Kehlstein bus (4 people)",                      80.00, True),
    ("Activities", "Königssee electric boat (4 people)",                         88.00, True),
    ("Activities", "Hohensalzburg fortress admission (4 people)",                66.00, True),
    ("Activities", "St. Peter Stiftskeller dinner (4 people)",                  220.00, True),
    ("Activities", "Schönbrunn Palace + maze (4 people)",                       132.00, True),

    # MEALS & INCIDENTALS
    ("Meals & Incidentals", "Meals (~$150/day × 10 days)",      1500.00, True),
    ("Meals & Incidentals", "Miscellaneous / buffer",             500.00, True),
]

# ── Helpers ───────────────────────────────────────────────────────────────────

def usd(n: float) -> str:
    return f"${n:,.0f}"

def get_subtotals() -> dict[str, float]:
    result: dict[str, float] = {}
    for cat, _, amt, _ in COSTS:
        result[cat] = result.get(cat, 0.0) + amt
    return result

def get_grand_total() -> float:
    return sum(amt for _, _, amt, _ in COSTS)

# ── Spreadsheet ───────────────────────────────────────────────────────────────

def build_spreadsheet() -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Trip Costs"

    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 54
    ws.column_dimensions["C"].width = 16

    def fill(hex_color: str) -> PatternFill:
        return PatternFill("solid", fgColor=hex_color)

    h_fill = fill("2C5F8A")
    c_fill = fill("D5E8F5")
    t_fill = fill("2C5F8A")
    h_font = Font(bold=True, color="FFFFFF", size=11)
    c_font = Font(bold=True, size=10)
    t_font = Font(bold=True, color="FFFFFF", size=11)
    money_fmt = '"$"#,##0.00'
    center = Alignment(horizontal="center")

    # Header row
    ws.append(["Category", "Item", "Amount (USD)"])
    for col in range(1, 4):
        cell = ws.cell(1, col)
        cell.font = h_font
        cell.fill = h_fill
        cell.alignment = center

    row = 1
    prev_cat = None
    for cat, item, amt, est in COSTS:
        if cat != prev_cat:
            row += 1
            ws.cell(row, 1, cat).font = c_font
            for col in range(1, 4):
                ws.cell(row, col).fill = c_fill
            prev_cat = cat

        row += 1
        label = item + (" *" if est else "")
        item_cell = ws.cell(row, 2, label)
        if est:
            item_cell.font = Font(italic=True)
        amt_cell = ws.cell(row, 3, amt)
        amt_cell.number_format = money_fmt

    last_data_row = row

    # Blank row then TOTAL
    row += 2
    ws.cell(row, 1, "TOTAL").font = t_font
    ws.cell(row, 1).fill = t_fill
    ws.cell(row, 1).alignment = Alignment(horizontal="right")
    ws.cell(row, 2).fill = t_fill
    total_cell = ws.cell(row, 3)
    total_cell.value = f"=SUM(C2:C{last_data_row})"
    total_cell.number_format = money_fmt
    total_cell.font = t_font
    total_cell.fill = t_fill

    row += 2
    ws.cell(row, 1, "* = estimated").font = Font(italic=True, color="888888")

    out = BASE / "trip_costs.xlsx"
    wb.save(out)
    print(f"  Wrote {out.name}")

# ── Sentinel markers (same string used in both MD and HTML) ───────────────────

MARKER_START = "<!-- COSTS_START -->"
MARKER_END   = "<!-- COSTS_END -->"

# ── Markdown section ──────────────────────────────────────────────────────────

def build_md_section() -> str:
    total = get_grand_total()
    lines = [
        "## Cost Estimate",
        "",
        "*Amounts in USD. Rows marked \\* are estimates — confirm before trip. "
        "Total computed by `build_costs.py`; source data: `trip_costs.xlsx`.*",
        "",
        "| Category | Item | Amount |",
        "|---|---|---|",
    ]
    prev_cat = None
    for cat, item, amt, est in COSTS:
        if cat != prev_cat:
            lines.append(f"| **{cat}** | | |")
            prev_cat = cat
        label = item + (" \\*" if est else "")
        lines.append(f"| | {label} | {usd(amt)} |")

    lines += [
        "| | | |",
        f"| **TOTAL** | | **{usd(total)}** |",
    ]
    return "\n".join(lines)

def update_md() -> None:
    path = BASE / "itinerary.md"
    text = path.read_text(encoding="utf-8")
    block = f"{MARKER_START}\n{build_md_section()}\n{MARKER_END}"

    if MARKER_START in text:
        pattern = re.escape(MARKER_START) + r".*?" + re.escape(MARKER_END)
        text = re.sub(pattern, block, text, flags=re.DOTALL)
    else:
        # First run: insert before "## To Add"
        text = text.replace("## To Add", f"{block}\n\n---\n\n## To Add")

    path.write_text(text, encoding="utf-8")
    print(f"  Updated itinerary.md")

# ── HTML section ──────────────────────────────────────────────────────────────

COST_CSS = """\
    .cost-table { width: 100%; border-collapse: collapse; font-size: 0.93em; margin-top: 10px; }
    .cost-table th { background: #2c5f8a; color: #fff; padding: 7px 10px; text-align: left; }
    .cost-table td { padding: 5px 10px; border-bottom: 1px solid #e8e4dc; vertical-align: top; }
    .cost-cat-header td { background: #d5e8f5; font-weight: bold; padding: 6px 10px; }
    .cost-est td:nth-child(2) { font-style: italic; color: #555; }
    .cost-amt { text-align: right !important; font-variant-numeric: tabular-nums; white-space: nowrap; }
    .cost-total-row td { background: #2c5f8a; color: #fff; font-weight: bold; padding: 8px 10px; border: none; }"""

def build_html_section() -> str:
    total = get_grand_total()
    rows: list[str] = []
    prev_cat = None
    for cat, item, amt, est in COSTS:
        if cat != prev_cat:
            rows.append(
                f'        <tr class="cost-cat-header">'
                f'<td colspan="2">{cat}</td><td></td></tr>'
            )
            prev_cat = cat
        cls = ' class="cost-est"' if est else ""
        label = item + (" *" if est else "")
        rows.append(
            f'        <tr{cls}><td style="width:2em"></td>'
            f'<td>{label}</td>'
            f'<td class="cost-amt">{usd(amt)}</td></tr>'
        )

    return (
        f"{MARKER_START}\n"
        f"  <section>\n"
        f"    <h2>Cost Estimate</h2>\n"
        f"    <p style=\"font-size:0.88em;color:#666;margin-bottom:10px\">"
        f"All amounts USD. <em>Italic rows (*) are estimates — confirm before trip.</em> "
        f"Total verified by <code>build_costs.py</code> · source: <code>trip_costs.xlsx</code>.</p>\n"
        f"    <table class=\"cost-table\">\n"
        f"      <thead><tr><th colspan=\"2\">Item</th><th>Amount</th></tr></thead>\n"
        f"      <tbody>\n"
        + "\n".join(rows) + "\n"
        f"        <tr class=\"cost-total-row\">"
        f'<td colspan="2">TOTAL</td>'
        f'<td class="cost-amt">{usd(total)}</td></tr>\n'
        f"      </tbody>\n"
        f"    </table>\n"
        f"  </section>\n"
        f"{MARKER_END}"
    )

def update_html() -> None:
    path = BASE / "itinerary.html"
    text = path.read_text(encoding="utf-8")

    # Inject CSS once
    if ".cost-table" not in text:
        text = text.replace("</style>", COST_CSS + "\n  </style>")

    block = build_html_section()
    if MARKER_START in text:
        pattern = re.escape(MARKER_START) + r".*?" + re.escape(MARKER_END)
        text = re.sub(pattern, block, text, flags=re.DOTALL)
    else:
        # First run: insert before <!-- PHOTO GALLERY -->
        text = text.replace("  <!-- PHOTO GALLERY -->", f"{block}\n\n  <!-- PHOTO GALLERY -->")

    path.write_text(text, encoding="utf-8")
    print(f"  Updated itinerary.html")

# ── Main ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Building cost estimate…")
    build_spreadsheet()
    update_md()
    update_html()

    print("\nSummary:")
    subs = get_subtotals()
    for cat, amt in subs.items():
        print(f"  {cat:<26}  {usd(amt):>9}")
    print(f"  {'-' * 38}")
    print(f"  {'TOTAL':<26}  {usd(get_grand_total()):>9}")

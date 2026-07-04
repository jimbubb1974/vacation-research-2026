"""Generate synchronized itinerary sections and the printable daily handout."""

from __future__ import annotations

import html
import re
from pathlib import Path

from itinerary_data import DAYS, TRIP
from map_data import MAPS


BASE = Path(__file__).resolve().parent
OVERVIEW_START = "<!-- ITINERARY_OVERVIEW_START -->"
OVERVIEW_END = "<!-- ITINERARY_OVERVIEW_END -->"
DETAIL_START = "<!-- ITINERARY_DAILY_DETAIL_START -->"
DETAIL_END = "<!-- ITINERARY_DAILY_DETAIL_END -->"


def inline_html(text: str) -> str:
    """Render the small Markdown subset used by itinerary_data.py."""
    rendered = html.escape(text, quote=False)
    rendered = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", rendered)
    rendered = re.sub(r"\*(.+?)\*", r"<em>\1</em>", rendered)
    return rendered


def md_overview() -> str:
    rows = [
        OVERVIEW_START,
        "## Day-by-Day",
        "",
        "| Day | Date | Location | Overnight | Plan |",
        "|---|---|---|---|---|",
    ]
    for day in DAYS:
        rows.append(
            f"| {day['day']} | {day['date']} | {day['location']} | "
            f"{day['overnight']} | {day['summary']} |"
        )
    rows.extend([OVERVIEW_END])
    return "\n".join(rows)


def md_detail() -> str:
    rows = [DETAIL_START, "## Daily Detail", ""]
    for day in DAYS:
        rows.append(f"### Day {day['day']} — {day['date']} · {day['title']}")
        for item in day["events"]:
            rows.append(f"- **{item['time']}** — {item['text']}")
            if item.get("where"):
                rows.append(f"  - **Where / how:** {item['where']}")
        for note in day.get("notes", []):
            rows.append(f"> **Day note:** {note}")
        rows.append("")
    rows.append(DETAIL_END)
    return "\n".join(rows).rstrip()


def html_overview() -> str:
    rows = [
        f"  {OVERVIEW_START}",
        "  <section>",
        "    <h2>Day-by-Day</h2>",
        "    <table>",
        "      <thead><tr><th>Day</th><th>Date</th><th>Location</th><th>Overnight</th><th>Plan</th></tr></thead>",
        "      <tbody>",
    ]
    for day in DAYS:
        cls = ' class="flight-row"' if day["day"] in (0, 11) else ""
        rows.extend(
            [
                f"        <tr{cls}>",
                f"          <td class=\"day-num\">{day['day']}</td>",
                f"          <td class=\"day-date\">{html.escape(day['date'])}</td>",
                f"          <td class=\"location\">{html.escape(day['location'])}</td>",
                f"          <td class=\"overnight\">{html.escape(day['overnight'])}</td>",
                f"          <td>{inline_html(day['summary'])}</td>",
                "        </tr>",
            ]
        )
    rows.extend(["      </tbody>", "    </table>", "  </section>", f"  {OVERVIEW_END}"])
    return "\n".join(rows)


def html_detail() -> str:
    rows = [f"  {DETAIL_START}", "  <section>", "    <h2>Daily Detail</h2>"]
    for day in DAYS:
        rows.extend([f"    <h3>Day {day['day']} — {html.escape(day['date'])} · {html.escape(day['title'])}</h3>", "    <ul>"])
        for item in day["events"]:
            rows.append(f"      <li><strong>{html.escape(item['time'])}</strong> — {inline_html(item['text'])}")
            if item.get("where"):
                rows.append(f"        <div class=\"event-where\"><strong>Where / how:</strong> {inline_html(item['where'])}</div>")
            rows.append("      </li>")
        rows.append("    </ul>")
        for note in day.get("notes", []):
            rows.append(f"    <div class=\"note-box\"><strong>Day note:</strong> {inline_html(note)}</div>")
    rows.extend(["  </section>", f"  {DETAIL_END}"])
    return "\n".join(rows)


def daily_map(day: dict) -> str:
    """Render a compact, print-safe location map for a daily sheet."""
    map_ids = day.get("map")
    if not map_ids:
        return '<aside class="map-notes"><h2>Map / route notes</h2><div class="writing-lines"></div></aside>'
    if isinstance(map_ids, str):
        map_ids = [map_ids]

    rendered = []
    for map_id in map_ids:
        map_data = MAPS[map_id]
        if not (BASE / map_data["output"]).exists():
            raise RuntimeError(f'Missing generated map {map_data["output"]}; run python3 build_maps.py')
        places = map_data["places"]
        legend = "".join(
            '<li>'
            f'<span>{place["number"]}</span><strong>{html.escape(place["name"])}</strong>'
            f'<small>{html.escape(place["address"])}</small>'
            "</li>"
            for place in places
        )
        rendered.append(
            f'<aside class="location-map{" compact" if map_data.get("compact") else ""}">'
            '<div class="map-layout">'
            f'<img class="map-image" src="{html.escape(map_data["output"])}" alt="{html.escape(map_data["alt"], quote=True)}">'
            '<div class="map-key">'
            f'<h2>{html.escape(map_data["title"])}</h2>'
            f'<ol class="map-legend">{legend}</ol>'
            '</div>'
            '</div>'
            '</aside>'
        )
    return "".join(rendered)


def replace_or_bootstrap(text: str, start: str, end: str, block: str, bootstrap_pattern: str) -> str:
    if start in text and end in text:
        pattern = r"(?m)^[ \t]*" + re.escape(start) + r".*?" + re.escape(end)
    else:
        pattern = bootstrap_pattern
    updated, count = re.subn(pattern, block, text, count=1, flags=re.DOTALL)
    if count != 1:
        raise RuntimeError(f"Could not locate exactly one generated section for {start}")
    return updated


def update_markdown() -> None:
    path = BASE / "itinerary.md"
    text = path.read_text(encoding="utf-8")
    text = replace_or_bootstrap(
        text,
        OVERVIEW_START,
        OVERVIEW_END,
        md_overview(),
        r"## Day-by-Day\n.*?(?=\n---\n\n## Base Summary)",
    )
    text = replace_or_bootstrap(
        text,
        DETAIL_START,
        DETAIL_END,
        md_detail(),
        r"## Daily Detail\n.*?(?=\n---\n\n## Booking Priorities)",
    )
    path.write_text(text, encoding="utf-8")


def update_html() -> None:
    path = BASE / "itinerary.html"
    text = path.read_text(encoding="utf-8")
    if ".event-where" not in text:
        text = text.replace(
            "  .note-box strong { color: #b07800; }",
            "  .note-box strong { color: #b07800; }\n"
            "  .event-where { margin: 5px 0 4px 12px; padding: 6px 9px; background: #eef5fa; "
            "border-left: 3px solid #2e6da4; color: #465563; font-family: sans-serif; font-size: 0.9em; }",
        )
    text = replace_or_bootstrap(
        text,
        OVERVIEW_START,
        OVERVIEW_END,
        html_overview(),
        r"  <!-- DAY BY DAY -->\n.*?(?=\n  <!-- BASE SUMMARY -->)",
    )
    text = replace_or_bootstrap(
        text,
        DETAIL_START,
        DETAIL_END,
        html_detail(),
        r"  <!-- DAILY DETAIL -->\n.*?(?=\n  <!-- BOOKING PRIORITIES -->)",
    )
    path.write_text(text, encoding="utf-8")


def daily_handout() -> str:
    sheets = []
    for day in DAYS:
        sheet_class = "day-sheet long" if len(day["events"]) > 10 else "day-sheet"
        events = []
        for item in day["events"]:
            where = ""
            if item.get("where"):
                where = f'<div class="where">{inline_html(item["where"])}</div>'
            events.append(
                f'<article class="event {html.escape(item["kind"])}">'
                f'<div class="plan">{inline_html(item["text"])}{where}</div>'
                f'<div class="time">{html.escape(item["time"])}</div>'
                "</article>"
            )
        notes = "".join(f"<li>{inline_html(note)}</li>" for note in day.get("notes", []))
        sheets.append(
            f'<section class="{sheet_class}" id="day-{day["day"]}">'
            '<header class="day-header">'
            f'<div><div class="eyebrow">Day {day["day"]} · {html.escape(day["date"])}</div>'
            f'<h1>{html.escape(day["title"])}</h1>'
            f'<div class="location">{html.escape(day["location"])}</div></div>'
            f'<div class="overnight"><span>Tonight</span>{html.escape(day["overnight"])}</div>'
            "</header>"
            f'<main class="timeline">{"".join(events)}</main>'
            f'<aside class="day-notes"><h2>Carry / booking notes</h2><ul>{notes}</ul></aside>'
            f'{daily_map(day)}'
            f'<footer>{html.escape(TRIP["title"])} · {html.escape(TRIP["dates"])} · Day {day["day"]}</footer>'
            "</section>"
        )

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{html.escape(TRIP['title'])} — Daily Handout</title>
<style>
  :root {{ --navy:#173a59; --blue:#2f6f9f; --pale:#edf4f8; --ink:#202832; --muted:#63717d; --rule:#cfdae2; --gold:#b98021; }}
  * {{ box-sizing:border-box; }}
  html, body {{ margin:0; padding:0; }}
  body {{ background:#d9dee2; color:var(--ink); font-family:Arial, Helvetica, sans-serif; }}
  .day-sheet {{ width:8.5in; min-height:11in; margin:.25in auto; padding:.42in .48in .34in; background:white; box-shadow:0 3px 18px #0002; break-after:page; page-break-after:always; position:relative; }}
  .day-sheet:last-child {{ break-after:auto; page-break-after:auto; }}
  .day-header {{ display:flex; justify-content:space-between; gap:.3in; padding-bottom:.14in; border-bottom:3px solid var(--navy); break-inside:avoid; }}
  .eyebrow {{ color:var(--blue); font-size:9pt; font-weight:700; letter-spacing:.12em; text-transform:uppercase; }}
  h1 {{ margin:.04in 0 .02in; color:var(--navy); font-family:Georgia, serif; font-size:23pt; line-height:1.08; }}
  .day-header .location {{ color:var(--muted); font-size:10pt; }}
  .overnight {{ min-width:1.65in; max-width:2.25in; align-self:center; color:var(--navy); font-size:9pt; font-weight:700; text-align:right; }}
  .overnight span {{ display:block; color:var(--muted); font-size:7.5pt; letter-spacing:.1em; text-transform:uppercase; }}
  .timeline {{ margin-top:.13in; border-top:1px solid var(--rule); }}
  .event {{ display:grid; grid-template-columns:minmax(0, 1fr) 1.18in; gap:.18in; padding:.095in 0; border-bottom:1px solid var(--rule); break-inside:avoid; page-break-inside:avoid; }}
  .plan {{ font-family:Georgia, serif; font-size:9.35pt; line-height:1.34; }}
  .time {{ color:var(--navy); font-size:9.3pt; font-weight:800; line-height:1.25; text-align:right; white-space:normal; }}
  .where {{ margin-top:.055in; padding:.055in .075in; background:var(--pale); border-left:3px solid var(--blue); color:#41515e; font-family:Arial, sans-serif; font-size:7.9pt; line-height:1.3; }}
  .day-notes, .map-notes {{ margin-top:.12in; padding:.08in .11in; border:1px solid var(--rule); break-inside:avoid; page-break-inside:avoid; }}
  .day-notes {{ background:#fffaf0; border-left:4px solid var(--gold); }}
  .day-notes h2, .map-notes h2 {{ margin:0 0 .035in; color:var(--navy); font-size:7.2pt; letter-spacing:.09em; text-transform:uppercase; }}
  .day-notes ul {{ margin:0; padding-left:.17in; font-size:7.8pt; line-height:1.3; }}
  .map-notes {{ min-height:.48in; }}
  .writing-lines {{ height:.19in; background:repeating-linear-gradient(to bottom, transparent 0, transparent .085in, #d7e0e6 .09in); }}
  .location-map {{ margin-top:.12in; padding:.08in .11in; border:1px solid var(--rule); break-inside:avoid; page-break-inside:avoid; }}
  .location-map h2 {{ margin:0 0 .06in; color:var(--navy); font-size:7.2pt; line-height:1.15; letter-spacing:.09em; text-transform:uppercase; }}
  .map-layout {{ display:grid; grid-template-columns:minmax(0, 1.55fr) minmax(1.75in, 1fr); gap:.12in; align-items:center; }}
  .map-image {{ display:block; width:100%; height:1.72in; object-fit:cover; border:1px solid #aab8c1; background:#e8e4dc; }}
  .map-key {{ min-width:0; }}
  .map-legend {{ margin:0; padding:0; list-style:none; }}
  .map-legend li {{ display:grid; grid-template-columns:.22in 1fr; align-items:center; margin:.045in 0; font-size:7.8pt; line-height:1.15; }}
  .map-legend span {{ grid-row:1 / 3; display:grid; place-items:center; width:.19in; height:.19in; border-radius:50%; background:var(--navy); color:#fff; font-size:7pt; font-weight:700; }}
  .map-legend strong {{ color:var(--navy); }}
  .map-legend small {{ color:var(--muted); font-size:6.8pt; }}
  .location-map.compact .map-image {{ height:1.35in; }}
  .location-map.compact h2 {{ margin-bottom:.035in; font-size:6.8pt; }}
  .location-map.compact .map-legend li {{ margin:.02in 0; font-size:7.2pt; }}
  .location-map.compact .map-legend small {{ font-size:6.3pt; }}
  .day-sheet.long .event {{ padding:.075in 0; }}
  footer {{ margin-top:.1in; color:#85919a; font-size:6.8pt; letter-spacing:.04em; text-align:center; break-inside:avoid; }}
  @page {{ size:Letter portrait; margin:0; }}
  @media print {{
    body {{ background:white; }}
    .day-sheet {{ width:8.5in; min-height:11in; margin:0; box-shadow:none; }}
    footer {{ display:none; }}
    .day-sheet.long .map-notes {{ display:none; }}
  }}
  @media screen and (max-width:8.5in) {{
    .day-sheet {{ width:100%; min-height:0; margin:0 0 16px; padding:24px 20px; }}
  }}
</style>
</head>
<body>
{''.join(sheets)}
</body>
</html>
'''


def write_daily_handout() -> None:
    (BASE / "daily_itinerary.html").write_text(daily_handout(), encoding="utf-8")


def main() -> None:
    update_markdown()
    update_html()
    write_daily_handout()
    print("Updated itinerary.md")
    print("Updated itinerary.html")
    print("Wrote daily_itinerary.html")


if __name__ == "__main__":
    main()

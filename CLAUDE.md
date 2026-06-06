# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Personal vacation planning project for a family of 4 (2 adults + 2 teens, ages 17 and 15) for summer 2026. The **current itinerary is Bavaria + Salzburg + Vienna, July 11–21, 2026** (`itinerary.md`). Other destination options (Greece, St. Lucia) remain in the research files for reference.

## Key Files

- `itinerary.md` — **canonical working itinerary** (in progress; day-by-day, updated as details are confirmed)
- `bavaria option 2.md` — original planning framework: day-by-day options, lodging recommendations, cost estimates
- `bavaria_itinerary_locked.md` — earlier version of the itinerary; superseded by `itinerary.md`
- `bavaria2_brochure.html` — styled HTML brochure for the trip (PDF export: `bavaria2_brochure.pdf`)
- `build_vacation_spreadsheet.py` — generates `vacation_research_2026.xlsx` with cost comparison tabs
- `brochure_images/` — photos used in brochures (day-by-day filenames like `day4_neuschwanstein.jpg`)
- `archive/` — all non-Bavaria files (Greece, St. Lucia research, Option 1 brochures, working screenshots)

## Planning Research Files (to be created)

As trip details are researched and booked, maintain these focused MD files:

- `flights.md` — Open-jaw: IAD → MUC outbound (Jul 11), VIE → IAD return (Jul 21); airlines, booking status
- `accommodation.md` — Airbnb/hotel options per city (Munich, Berchtesgaden, Salzburg) with links and status
- `trains.md` — Bayern-Ticket logistics, day-trip connections, Railjet booking
- `activities.md` — All bookable items (Eagle's Nest, Mike's Bike Tours, St. Peter Stiftskeller, etc.) with links, prices, booking priority/status

## Running the Spreadsheet Builder

```bash
python build_vacation_spreadsheet.py
```

Requires `openpyxl`. Output is `vacation_research_2026.xlsx`. The script builds multiple tabs: Summary, itinerary tabs per destination, and per-destination cost breakdown tabs. Cell cross-references use named sheets (e.g., `=Bavaria_Costs!F28`).

## Brochure Workflow

Brochures are self-contained HTML files with inline CSS and `<img>` tags referencing `brochure_images/`. To export to PDF, open in a browser and print to PDF (or use Playwright). Screenshots of the rendered output are saved as `*_map_check.png` and `*_brochure.pdf` at the root.

## Conventions

- `*.BACKUP.md` files are gitignored; `*.BACKUP2.md` files are not — treat the non-ignored backups as reference snapshots, not the authoritative version.
- Loose JPEG files at the root (prefixed `preview_`, `check_`, or named by location) are working screenshots used during brochure development, not source assets.
- `_tmp_train_imgs.json` is a large scratch file from image search sessions — ignore it.

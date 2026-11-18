# BC Energy Step Code Dashboard

A filterable dashboard of BC Energy Step Code home evaluation data, built from
the "Conditioned Data" tab of `NRCan Database Analysis.xlsx`.

**Live site:** published via GitHub Pages from this repo (see repo settings for the URL).

## What's here

- `extract_data.py` — streams the Conditioned Data sheet directly from the
  workbook (no Excel/pandas needed) and aggregates it into `data.json`,
  grouped by municipality and year built, with counts/sums for MEUI, TEDI,
  airtightness, GHGs, envelope insulation, and category breakdowns (heating,
  cooling, DHW, step code achieved, etc.).
- `regions.py` — a hand-built lookup from each municipality string found in
  the data (including typos and rural localities) to its BC regional
  district, used to group municipalities in the dashboard.
- `data.json` — the aggregated dataset the page reads. Pre-aggregated by
  (municipality, year) so the page stays small and fast even though the
  source sheet has tens of thousands of rows.
- `index.html` — the dashboard itself: a static page (D3.js + vanilla JS)
  that fetches `data.json` and renders filterable charts. No build step,
  no backend.

## Refreshing the data

Whenever the source workbook has new data:

```bash
python extract_data.py "path to NRCan Database Analysis.xlsx"
```

This regenerates `data.json` in place. Commit and push it to update the
published site. If a municipality is missing from `regions.py`, the script
reports it at the end of the run as "Unmapped / other" so it can be added.

## Privacy

`data.json` contains only counts and sums aggregated by municipality and
year built — no addresses, names, or other record-level identifiers.

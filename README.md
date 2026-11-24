# Cold-calling lead list: UK kitchen companies

Daily list of kitchen design & installation companies (showrooms / supply + fit) with **fewer than 150 Google reviews**, for DNV Leads cold calling.

- `leads/YYYY-MM-DD.xlsx`: that day's call sheet, kept simple: Business name, Owner name, Phone, Reviews, Area (sorted by area). Import into Google Sheets via File → Import → Upload.
- `leads/all_leads.csv`: every lead so far with full detail (website, address, Companies House no., all directors), sorted by city then date.
- `data/towns.txt`: search order (highest potential first). Each run carries on down the list.
- `data/seen_place_ids.json`: companies already listed, so nobody appears twice.
- `data/run_log.json`: leads and Google API requests used per run (keep requests inside the free allowance).

Filters: open businesses only; national chains removed (Howdens, Wren, Magnet, Harvey Jones, Tom Howley, Neptune…); off-niche trades removed (bathrooms, tiles, stone, builders, carpenters, interior designers, fit-only fitters). Franchise showrooms are kept but flagged.

Run: `python3 scripts/build_leads.py` (needs the Places API key saved as a network secret for `places.googleapis.com`, header `X-Goog-Api-Key`).

Owner names: `python3 scripts/add_directors.py` looks up active directors on the public Companies House register. "(verify)" = name-only match, confirm on the call. Blank = no limited company found (likely a sole trader), ask for "the owner".

Before dialling: check each number against TPS/CTPS (sole traders are often registered).

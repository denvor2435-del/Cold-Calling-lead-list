#!/usr/bin/env python3
"""Daily lead list: UK kitchen design & installation companies with < 150 Google reviews.

Searches towns in priority order (data/towns.txt), skips companies already listed on
earlier days (data/seen_place_ids.json), and writes leads/YYYY-MM-DD.xlsx plus a
running master list (leads/all_leads.csv), sorted by city then date.

The Google Places API key is injected by the session's network proxy (never stored here).
"""
import csv, datetime as dt, json, os, re, sys, time, urllib.request, urllib.error
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET = int(os.environ.get("TARGET", 150))
MAX_REVIEWS = 150
MAX_REQUESTS = int(os.environ.get("MAX_REQUESTS", 60))  # hard cap per run, protects the free allowance
QUERIES = ["kitchen showroom {t}", "fitted kitchens {t}", "bespoke kitchens {t}"]
FIELDS = ",".join("places." + f for f in [
    "id", "displayName", "formattedAddress", "nationalPhoneNumber", "websiteUri",
    "rating", "userRatingCount", "businessStatus", "primaryType", "googleMapsUri"]) + ",nextPageToken"

# National chains / corporate stores: no owner to pitch. Fit-only door swaps are off-niche.
EXCLUDE = re.compile(r"howdens|wren|magnet|wickes|b&q|ikea|john lewis|harvey jones|tom howley|neptune|"
                     r"devol|benchmarx|homebase|symphony|topps tiles|dream doors|second nature|"
                     r"jewson|travis perkins|screwfix|selco|currys|marks electrical|plumbase|wolseley", re.I)
# Off-niche: other trades, suppliers, fit-only fitters. Dropped unless the name also says kitchen.
OFF_NICHE = re.compile(r"bathroom|tile|ceramic|stone|marble|quartz|granite|worktop|robert dyas|boconcept|"
                       r"sink|plumbing|handyman|locksmith|architect|construction|builders|build\b|loft|"
                       r"extension|roofing|floor|carpet|curtain|blind|wardrobe|bedroom furniture|garden|"
                       r"outdoor|electrical|appliance|bed\b|sofa|carpent|interior design|handles|\btops\b|"
                       r"home improvement|contruction|conversion|contracts|neatsmith|tinkletorium", re.I)
ALWAYS_OFF = re.compile(r"outdoor kitchen|worktop|commercial kitchen|catering|steelplan|architect|"
                        r"planning permission|kitchen in the garden|fitters?\b|install|fit a kitchen|"
                        r"fit my kitchen|handyman|tradesman|freelance|cookware|kitchen shop|kitchenware", re.I)
KITCHEN = re.compile(r"kitchen|k&b|cuisine", re.I)
POSTCODE = re.compile(r"\s*[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$")

def fit_for_niche(title):
    """'' = keep, otherwise the reason it was dropped."""
    if EXCLUDE.search(title): return "chain"
    if ALWAYS_OFF.search(title): return "off-niche"
    if KITCHEN.search(title): return ""
    if OFF_NICHE.search(title): return "off-niche"
    return ""

def town_of(address, fallback):
    parts = [x.strip() for x in address.split(",") if x.strip() and x.strip() != "UK"]
    if not parts: return fallback
    t = POSTCODE.sub("", parts[-1]).strip()
    return t or fallback

FRANCHISE = re.compile(r"kutchenhaus|schmidt|mobalpa|cuisines|nolte|hacker|poggenpohl|siematic|"
                       r"bulthaup|leicht|roundhouse|mark wilkinson|smallbone|clive christian|"
                       r"masterclass|kitchen makers|sigma 3|moben|kitchens international|optiplan", re.I)

COLS = ["City", "Search area", "Date", "Company", "Phone", "Google reviews", "Rating", "Franchise?",
        "Owner / director", "Companies House", "TPS/CTPS checked", "Website", "Address", "Google Maps", "Call outcome", "Notes"]

def load(path, default):
    p = os.path.join(ROOT, path)
    return json.load(open(p)) if os.path.exists(p) else default

def save(path, obj):
    with open(os.path.join(ROOT, path), "w") as f:
        json.dump(obj, f, indent=1, sort_keys=True)

requests_used = 0
def search(query, page_token=None):
    global requests_used
    if requests_used >= MAX_REQUESTS:
        raise RuntimeError("request cap reached")
    body = {"textQuery": query, "regionCode": "GB", "pageSize": 20}
    if page_token:
        body["pageToken"] = page_token
    req = urllib.request.Request("https://places.googleapis.com/v1/places:searchText",
        data=json.dumps(body).encode(), method="POST",
        headers={"Content-Type": "application/json", "X-Goog-FieldMask": FIELDS})
    requests_used += 1
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            msg = e.read().decode()[:300]
            if e.code in (429, 500, 503) and attempt < 2:
                time.sleep(3 * (attempt + 1)); continue
            raise RuntimeError(f"Places API {e.code}: {msg}")

def main():
    today = os.environ.get("RUN_DATE", dt.date.today().isoformat())
    towns = [l.strip() for l in open(os.path.join(ROOT, "data/towns.txt"))
             if l.strip() and not l.startswith("#")]
    seen = set(load("data/seen_place_ids.json", []))
    done = load("data/towns_done.json", {})
    master = os.path.join(ROOT, "leads/all_leads.csv")
    rows = list(csv.DictReader(open(master))) if os.path.exists(master) else []
    leads = [r for r in rows if r["Date"] == today]          # re-running the same day tops up
    rows = [r for r in rows if r["Date"] != today]
    raw = []
    for town in towns:
        if len(leads) >= TARGET or requests_used >= MAX_REQUESTS:
            break
        if town in done:
            continue
        name = town.split(",")[0]
        try:
            for q in QUERIES:
                token, pages = None, 0
                while pages < 3:
                    d = search(q.format(t=town), token); pages += 1
                    for p in d.get("places", []):
                        pid = p["id"]
                        if pid in seen or p.get("businessStatus", "OPERATIONAL") != "OPERATIONAL":
                            continue
                        title = p.get("displayName", {}).get("text", "")
                        n = p.get("userRatingCount", 0) or 0
                        raw.append(p)
                        if n >= MAX_REVIEWS or fit_for_niche(title) or not p.get("nationalPhoneNumber"):
                            continue
                        seen.add(pid)
                        leads.append({
                            "City": town_of(p.get("formattedAddress", ""), name), "Search area": name, "Date": today, "Company": title,
                            "Phone": p.get("nationalPhoneNumber", ""),
                            "Google reviews": n, "Rating": p.get("rating", ""),
                            "Franchise?": "Yes" if FRANCHISE.search(title) else "",
                            "Owner / director": "", "TPS/CTPS checked": "",
                            "Website": p.get("websiteUri", ""), "Address": p.get("formattedAddress", ""),
                            "Google Maps": p.get("googleMapsUri", ""),
                            "Call outcome": "", "Notes": "", "place_id": pid})
                    token = d.get("nextPageToken")
                    if not token:
                        break
                    time.sleep(2)
            done[town] = today
        except RuntimeError as e:
            print("stopped:", e, file=sys.stderr)
            break

    leads.sort(key=lambda r: (r["City"], r["Date"], r["Company"]))
    cols = COLS
    write_xlsx(os.path.join(ROOT, f"leads/{today}.xlsx"), leads, cols, today)
    if raw:
        rp = os.path.join(ROOT, f"data/raw/{today}.json")
        old = json.load(open(rp)) if os.path.exists(rp) else []
        json.dump(old + raw, open(rp, "w"), indent=0)

    rows += [{k: r.get(k, "") for k in cols} for r in leads]
    rows.sort(key=lambda r: (r["City"], r["Date"]))
    if rows:
        with open(master, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(rows)

    save("data/seen_place_ids.json", sorted(seen))
    save("data/towns_done.json", done)
    log = load("data/run_log.json", [])
    log.append({"date": today, "leads": len(leads), "api_requests": requests_used,
                "search_areas_done": sorted(t for t, d in done.items() if d == today)})
    save("data/run_log.json", log)
    print(json.dumps(log[-1]))

SIMPLE = [("Business name", 38), ("Owner name", 24), ("Phone", 15), ("Reviews", 9), ("Area", 18)]

def owner_short(v):
    """First director only: 'Daniel Butler (+1: ...)' -> 'Daniel Butler'; keeps '(verify)'."""
    v = v or ""
    verify = v.endswith("(verify)")
    name = re.sub(r"\s*\(.*", "", v).strip()
    return f"{name} (verify)" if name and verify else name

def write_xlsx(path, leads, cols, today):
    """Call sheet kept deliberately simple: 5 columns, one tab, sorted by area.
    Full detail (website, address, Companies House no., etc.) stays in leads/all_leads.csv."""
    wb = Workbook()
    ws = wb.active; ws.title = today
    ws.append([c for c, _ in SIMPLE])
    for r in sorted(leads, key=lambda r: (r["City"], r["Company"])):
        try: n = int(r["Google reviews"] or 0)
        except ValueError: n = 0
        ws.append([r["Company"], owner_short(r.get("Owner / director", "")), r["Phone"], n, r["City"]])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
    for i, (_, w) in enumerate(SIMPLE, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(path)

if __name__ == "__main__":
    main()

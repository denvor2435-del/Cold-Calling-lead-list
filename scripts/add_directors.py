#!/usr/bin/env python3
"""Fill in owner/director names from the public Companies House register.

For every lead with an empty "Owner / director", search Companies House by company
name, pick the best live match (postcode match = High confidence, close name match =
Medium), then read the active directors. Rewrites leads/all_leads.csv and the daily
xlsx files that changed. Sole traders / partnerships are not on Companies House, so
those stay blank: ask for "the owner" on the call.
"""
import csv, html, importlib.util, os, re, sys, time, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
spec = importlib.util.spec_from_file_location("b", os.path.join(ROOT, "scripts/build_leads.py"))
b = importlib.util.module_from_spec(spec); spec.loader.exec_module(b)
BASE = "https://find-and-update.company-information.service.gov.uk"
STOP = {"ltd", "limited", "the", "and", "&", "co", "company", "uk", "of", "kitchens", "kitchen",
        "interiors", "design", "designs", "studio", "showroom", "bespoke", "surrey", "herts",
        "hertfordshire", "london", "plc", "llp", "group", "home", "homes", "living"}
PC = re.compile(r"([A-Z]{1,2}\d[A-Z\d]?)\s*(\d[A-Z]{2})")

def get(url):
    for attempt in range(3):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (lead research)"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read().decode("utf-8", "replace")
        except Exception:
            time.sleep(3 * (attempt + 1))
    return ""

def clean(name):
    name = re.split(r"\s[|\-–]\s|\(", name)[0]
    return re.sub(r"\s+", " ", name).strip()

def tokens(s):
    return {t for t in re.findall(r"[a-z0-9]+", s.lower().replace("&", " and ")) if t not in STOP}

def postcode(s):
    m = PC.search(s.upper()); return (m.group(1) + " " + m.group(2)) if m else ""

def search(q):
    page = get(f"{BASE}/search/companies?q={urllib.parse.quote_plus(q)}")
    out = []
    for m in re.finditer(r'href="/company/([A-Z0-9]+)"[^>]*>\s*(.*?)\s*</a>.*?crumbtrail">(.*?)</p>\s*<p>(.*?)</p>',
                         page, re.S):
        out.append({"no": m.group(1), "name": html.unescape(m.group(2)).strip(),
                    "desc": m.group(3), "addr": html.unescape(m.group(4))})
    return out

def best_match(lead):
    q = clean(lead["Company"])
    want, gpc = tokens(q), postcode(lead["Address"])
    best = None
    for c in search(q)[:10]:
        if re.match(r"(SC|SO|NI|NC|R)\d", c["no"]):   # Scottish / NI registrations: wrong country for these leads
            continue
        if "Dissolved" in c["desc"] or "liquidation" in c["desc"].lower():
            continue
        have = tokens(c["name"])
        overlap = len(want & have) / max(1, len(want | have)) if want else 0
        pc = postcode(c["addr"])
        same_pc = bool(gpc and pc == gpc)
        same_area = bool(gpc and pc and pc.split()[0] == gpc.split()[0])
        if same_pc and overlap >= 0.3: conf = "High"
        elif overlap >= 0.5 and same_area: conf = "High"
        elif overlap >= 0.99 and len(want) >= 2: conf = "High"
        elif overlap >= 0.6: conf = "Medium"
        else: continue
        score = (conf == "High", overlap, same_pc)
        if not best or score > best[0]:
            best = (score, conf, c)
    return (best[1], best[2]) if best else (None, None)

def directors(no):
    page = get(f"{BASE}/company/{no}/officers")
    people = []
    for blk in re.split(r'<div class="appointment-\d+">', page)[1:]:
        n = re.search(r'id="officer-name-\d+">\s*<a[^>]*>(.*?)</a>', blk, re.S)
        role = re.search(r'id="officer-role-\d+"[^>]*>\s*(.*?)\s*<', blk, re.S)
        if not n or "Resigned on" in blk:
            continue
        if role and "director" not in role.group(1).lower():
            continue
        raw = html.unescape(n.group(1)).strip()
        if "," in raw:
            sur, first = raw.split(",", 1)
            first = first.strip().split()[0] if first.strip() else ""
            if first.lower() in {"mr", "mrs", "ms", "miss", "dr"} and len(raw.split(",", 1)[1].split()) > 1:
                first = raw.split(",", 1)[1].split()[1]
            name = f"{first.title()} {sur.strip().title()}".strip()
        else:
            name = raw.title()  # corporate director
        people.append(name)
    return people

def main():
    master = os.path.join(ROOT, "leads/all_leads.csv")
    rows = list(csv.DictReader(open(master)))
    cols = b.COLS
    touched = set(); found = 0; tried = 0
    for r in rows:
        if r.get("Owner / director") or r.get("Companies House"):
            continue
        tried += 1
        conf, c = best_match(r); time.sleep(0.5)
        if not c:
            r["Companies House"] = "not found (sole trader?)"
            touched.add(r["Date"]); continue
        ds = directors(c["no"]); time.sleep(0.5)
        r["Companies House"] = f'{c["no"]} {c["name"]} ({conf})'
        if ds:
            r["Owner / director"] = ds[0] + (f" (+{len(ds)-1}: {', '.join(ds[1:3])})" if len(ds) > 1 else "")
            if conf != "High":
                r["Owner / director"] += " (verify)"
            found += 1
        touched.add(r["Date"])
    with open(master, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore"); w.writeheader(); w.writerows(rows)
    for d in sorted(touched):
        day = sorted([r for r in rows if r["Date"] == d], key=lambda r: (r["City"], r["Company"]))
        b.write_xlsx(os.path.join(ROOT, f"leads/{d}.xlsx"), day, cols, d)
    print(f"checked {tried}, directors found for {found}")

if __name__ == "__main__":
    main()

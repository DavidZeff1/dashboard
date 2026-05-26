#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_pniyot.py — read the service-requests Excel and generate data_pniyot.js
for dashboard_pniyot.html.

    python3 build_pniyot.py     # regenerates data_pniyot.js + prints a report

Source: row-level export "ללא שם 14052026 1513 רווחה 2025 -2026.xlsx", Sheet1.
Every metric is derived from the raw rows. The only curated pieces (in the maps below)
are display wording, colors, and the topic classifier — never the counts.

Median convention: the dashboard reports the upper-middle element (sorted[n//2]), so we
match that exactly (matters for small-n departments).
"""

import json, re, datetime
from collections import defaultdict, Counter

import openpyxl

XLSX = "ללא שם 14052026 1513 רווחה 2025 -2026.xlsx"
OUT = "data_pniyot.js"
MONTHS_2026 = 4.42   # Jan 1 – May 13 2026 elapsed-months divisor for YoY normalization

warnings = []

# ── curated presentation (wording / colors only) ──────────────────────────────
DEPT_DISPLAY = {  # raw -> (full name, short)
    "שירותים חברתיים מערב הגוש":        ("שירותים חברתיים מערב הגוש", "מערב הגוש"),
    "שירותים חברתיים - מזכירות אגף":     ("שירותים חברתיים — מזכירות אגף", "מזכירות אגף"),
    "שירותים חברתיים מזרח הגוש":        ("שירותים חברתיים מזרח הגוש", "מזרח הגוש"),
    "מיצוי זכויות - שירותים חברתיים":    ("מיצוי זכויות — שירותים חברתיים", "מיצוי זכויות"),
    'עו"ס משפחות - מערב':               ("עו״ס משפחות — מערב", "עו״ס מערב"),
    'עו"ס משפחות - מזרח':               ("עו״ס משפחות — מזרח", "עו״ס מזרח"),
    "שירותים חברתיים - התמכרויות ונוער": ("שירותים חברתיים — התמכרויות ונוער", "התמכרויות ונוער"),
}
SOURCE_DISPLAY = {
    "מוקד": "מוקד",
    "טופס יצירת קשר ראשוני- האגף לשירותים חברתיים": "טופס · יצירת קשר ראשוני",
    "יצירת קשר - אגף שירותים חברתיים": "טופס · יצירת קשר",
    "וואטסאפ כללי": "וואטסאפ",
}
STATUS_DISPLAY = {
    "הטיפול הסתיים": "הטיפול הסתיים",
    "פניה נפתחה במערכת המוקד": "פניה במערכת המוקד",
    "הפניה נסגרה": "הפניה נסגרה",
    "פניה כפולה": "פניה כפולה",
}
SOURCE_ORDER = ["מוקד", "טופס יצירת קשר ראשוני- האגף לשירותים חברתיים",
                "יצירת קשר - אגף שירותים חברתיים", "וואטסאפ כללי"]
STATUS_ORDER = ["הטיפול הסתיים", "פניה נפתחה במערכת המוקד", "הפניה נסגרה", "פניה כפולה"]

# topic buckets: display order (= dashboard order) + color
TOPIC_ORDER = [
    ("צמיד · ילדים בסיכון", "terra"), ("פנייה למזכירות", "slate"),
    ("אזרחים ותיקים", "ochre"), ("משפחות · כללי", "olive"),
    ("פנייה לעו״ס", "terra"), ("משפחות · בת עין", "olive"),
    ("משפחות · מזרח", "olive"), ("אחר · לא מסווג", "slate"),
    ("מיצוי זכויות", "ochre"), ("נוער / התמכרויות", "terra"),
    ("צרכים מיוחדים", "slate"), ("משפחות · הר גילה", "olive"),
    ("סדרי דין", "slate"), ("כללי · שירותים חברתיים", "slate"),
    ("משפחות · מערב", "olive"),
]
TOPIC_PRIORITY = {name: i for i, (name, _) in enumerate(TOPIC_ORDER)}
TOPIC_COLOR = {name: c for name, c in TOPIC_ORDER}

def classify(s):
    s = (s or "").strip()
    if "צמיד" in s: return "צמיד · ילדים בסיכון"
    if "אזרחים ותיקים" in s: return "אזרחים ותיקים"
    if "מזכירות" in s or "פניה אל אפרת" in s: return "פנייה למזכירות"
    if "סדרי דין" in s: return "סדרי דין"
    if "צרכים מיוחדים" in s: return "צרכים מיוחדים"
    if s == "מיצוי זכויות": return "מיצוי זכויות"
    if "בת עין" in s: return "משפחות · בת עין"
    if "הר גילה" in s: return "משפחות · הר גילה"
    if "משפח" in s and "מזרח" in s: return "משפחות · מזרח"
    if "משפח" in s and "מערב" in s: return "משפחות · מערב"
    if "משפח" in s: return "משפחות · כללי"
    if "נוער" in s or "התמכר" in s: return "נוער / התמכרויות"
    if "עובד" in s or 'עו"ס' in s: return "פנייה לעו״ס"
    if s == "שירותים חברתיים": return "כללי · שירותים חברתיים"
    return "אחר · לא מסווג"

RES_BUCKETS = [  # (label, color, lo_hours_exclusive, hi_hours_inclusive)
    ("≤ 24 שעות",     "olive",   -1,    24),
    ("24ש׳ – 3 ימים", "olive",    24,   72),
    ("3 – 7 ימים",    "ochre",    72,  168),
    ("7 – 30 ימים",   "terra",   168,  720),
    ("30+ ימים",      "crimson", 720,  1e12),
]
PEAK_NOTES = ["היום העמוס בשנה", "סוף שבוע ארוך", "ראש שבוע", ""]
HEB_MON = ["ינו׳","פבר׳","מרץ","אפר׳","מאי","יוני","יולי","אוג׳","ספט׳","אוק׳","נוב׳","דצמ׳"]
HEB_DOW = {6:"ראשון",0:"שני",1:"שלישי",2:"רביעי",3:"חמישי",4:"שישי",5:"שבת"}  # py weekday->Heb
DOW_ORDER = ["ראשון","שני","שלישי","רביעי","חמישי","שישי","שבת"]
EN_MONTHS = ["January","February","March","April","May","June","July","August",
             "September","October","November","December"]

def pdt(v):
    if isinstance(v, datetime.datetime): return v
    if v is None: return None
    s = str(v).strip()
    for f in ("%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None

def umed(lst):  # upper-middle median (matches dashboard)
    s = sorted(lst); return s[len(s) // 2]

import math
def jr(x, d=0):  # JS Math.round-style half-up (the dashboard was authored in JS)
    f = 10 ** d
    return math.floor(x * f + 0.5) / f

# ── load ──────────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, data_only=True)
rows = list(wb["Sheet1"].iter_rows(min_row=3, values_only=True))
# cols: 3=dept 4=topic 6=status 8=source 9=urgency 10=open 11=year 12=month(en) 13=dow 14=hour 15=close 16=target
N = len(rows)
D = {"total": N}

# monthly
mc = defaultdict(lambda: defaultdict(int))
for r in rows:
    mc[r[11]][r[12]] += 1
m2025 = [mc[2025][m] for m in EN_MONTHS]
max26 = max((EN_MONTHS.index(r[12]) for r in rows if r[11] == 2026), default=-1)
m2026 = [mc[2026][m] if i <= max26 else None for i, m in enumerate(EN_MONTHS)]
D["count2025"] = sum(m2025)
D["count2026"] = sum(v for v in m2026 if v)
D["monthly"] = {"labels": HEB_MON, "m2025": m2025, "m2026": m2026, "months2026": MONTHS_2026}

# hours (only hours that actually occur, ascending — matches the dashboard's label set)
hc = Counter(int(str(r[14]).split(":")[0]) for r in rows if r[14])
hours = sorted(hc)
D["hours"] = {"labels": [f"{h:02d}" for h in hours], "data": [hc[h] for h in hours]}
D["hour9to12"] = sum(hc[h] for h in (9, 10, 11, 12))

# day of week
dc = Counter(r[13] for r in rows)
D["dow"] = {"labels": DOW_ORDER, "data": [dc[d] for d in DOW_ORDER]}
D["friTotal"] = dc.get("שישי", 0); D["satTotal"] = dc.get("שבת", 0)

# resolution times (per dept + overall)
deptres = defaultdict(list); allres = []
for r in rows:
    o, c = pdt(r[10]), pdt(r[15])
    if o and c:
        h = (c - o).total_seconds() / 3600
        deptres[r[3]].append(h); allres.append(h)
D["resMeasured"] = len(allres)
D["overallMedianH"] = round(umed(allres), 1)
D["share24h"] = round(sum(1 for h in allres if h <= 24) / len(allres) * 100, 1)
D["share7d"]  = round(sum(1 for h in allres if h <= 168) / len(allres) * 100, 1)

# departments (count desc)
dept_count = Counter(r[3] for r in rows)
depts = []
for raw, cnt in dept_count.most_common():
    name, short = DEPT_DISPLAY.get(raw, (raw, raw))
    if raw not in DEPT_DISPLAY: warnings.append(f"no dept display for {raw!r}")
    depts.append({"name": name, "short": short, "count": cnt,
                  "medianH": round(umed(deptres[raw]), 1) if deptres[raw] else None})
D["depts"] = depts

# resolution histogram
resb = []
for label, color, lo, hi in RES_BUCKETS:
    cnt = sum(1 for h in allres if lo < h <= hi)
    resb.append({"label": label, "color": color, "count": cnt,
                 "pct": round(cnt / len(allres) * 100, 1)})
D["resBuckets"] = resb

# source / urgency / status
sc = Counter(r[8] for r in rows)
D["source"] = {"labels": [SOURCE_DISPLAY[k] for k in SOURCE_ORDER],
               "data": [sc.get(k, 0) for k in SOURCE_ORDER]}
uc = Counter(r[9] for r in rows)
D["urgency"] = {"labels": ["רגיל", "דחוף"], "data": [uc.get("רגיל", 0), uc.get("דחוף", 0)]}
stc = Counter(r[6] for r in rows)
D["status"] = {"labels": [STATUS_DISPLAY[k] for k in STATUS_ORDER],
               "data": [stc.get(k, 0) for k in STATUS_ORDER]}
D["closedCount"] = stc.get("הטיפול הסתיים", 0)
D["closedRate"] = round(stc.get("הטיפול הסתיים", 0) / N * 100, 1)

# SLA (close vs target)
before = after = 0
for r in rows:
    c, t = pdt(r[15]), pdt(r[16])
    if c and t:
        if c <= t: before += 1
        else: after += 1
D["slaBefore"], D["slaAfter"], D["slaTotal"] = before, after, before + after
D["slaPct"] = round(before / (before + after) * 100, 1)

# topics (count by classifier) + YoY trends
tot = defaultdict(int); y25 = defaultdict(int); y26 = defaultdict(int)
for r in rows:
    b = classify(r[4]); tot[b] += 1
    if r[11] == 2025: y25[b] += 1
    elif r[11] == 2026: y26[b] += 1
for b in tot:
    if b not in TOPIC_PRIORITY: warnings.append(f"unexpected topic bucket {b!r}")
topics = sorted(tot.items(), key=lambda kv: (-kv[1], TOPIC_PRIORITY.get(kv[0], 99)))
D["topics"] = [{"label": b, "count": c, "pct": round(c / N * 100, 1),
                "color": TOPIC_COLOR.get(b, "slate")} for b, c in topics]
trends = []
for b, _ in TOPIC_ORDER:
    if b == "אחר · לא מסווג":   # the dashboard does not trend the unclassified bucket
        continue
    r25 = y25[b] / 12; r26 = y26[b] / MONTHS_2026
    delta = int(jr((r26 - r25) / r25 * 100)) if r25 else (100 if r26 else 0)
    trends.append({"name": b, "r25": jr(r25, 1), "r26": jr(r26, 1), "delta": delta})
trends.sort(key=lambda t: -t["delta"])
D["trends"] = trends

# open/close flow (Jan'25 .. last month present)
def ym_list():
    out = []
    for y in (2025, 2026):
        for i, mo in enumerate(EN_MONTHS):
            if y == 2026 and i > max26: break
            out.append((y, i))
    return out
yms = ym_list()
openc = Counter((r[11], EN_MONTHS.index(r[12])) for r in rows)
closec = Counter()
for r in rows:
    c = pdt(r[15])
    if c and c.year in (2025, 2026):
        closec[(c.year, c.month - 1)] += 1
flow_labels = [f"{HEB_MON[i]}{str(y)[2:]}" for (y, i) in yms]
opened = [openc.get(k, 0) for k in yms]
closed = [closec.get(k, 0) for k in yms]
D["flow"] = {"labels": flow_labels, "opened": opened, "closed": closed}

# peak days (top 4 by open-date count)
datec = Counter()
for r in rows:
    o = pdt(r[10])
    if o: datec[o.date()] += 1
top = sorted(datec.items(), key=lambda kv: (-kv[1], kv[0]))[:4]
peaks = []
for i, (d, c) in enumerate(top):
    peaks.append({"day": d.day, "mon": HEB_MON[d.month - 1], "dow": HEB_DOW[d.weekday()],
                  "date": d.strftime("%d.%m.%Y"), "count": c,
                  "note": PEAK_NOTES[i] if i < len(PEAK_NOTES) else ""})
D["peaks"] = peaks

# hero / misc derived
dates = [pdt(r[10]) for r in rows if pdt(r[10])]
D["spanDays"] = (max(dates) - min(dates)).days
D["activeDays"] = len({d.date() for d in dates})
D["missingOpenDate"] = sum(1 for r in rows if not pdt(r[10]))
D["perActiveDay"] = round(N / D["activeDays"], 1)
D["mokedShare"] = round(sc.get("מוקד", 0) / N * 100, 1)

D["meta"] = {"generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
             "sourceFile": XLSX, "warnings": warnings}

with open(OUT, "w", encoding="utf-8") as fh:
    fh.write("/* AUTO-GENERATED by build_pniyot.py — do not edit by hand. */\n")
    fh.write("window.DASHDATA = ")
    json.dump(D, fh, ensure_ascii=False, indent=2)
    fh.write(";\n")

print(f"wrote {OUT}  ({N} tickets; {D['count2025']} in 2025, {D['count2026']} in 2026)")
print(f"  closed={D['closedRate']}%  medianH={D['overallMedianH']}  SLA={D['slaPct']}% ({before}/{after})")
print(f"  resMeasured={D['resMeasured']}  perActiveDay={D['perActiveDay']}  span={D['spanDays']}d  missing={D['missingOpenDate']}")
if warnings:
    print("WARNINGS:"); [print("  •", w) for w in warnings]

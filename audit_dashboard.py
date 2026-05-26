#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_dashboard.py — verify that the numbers hardcoded in dashboard.html
match the source Excel workbook, and check the source's internal consistency.

Run after every Excel refresh:
    python3 audit_dashboard.py

Exit code 0 = all checks pass. Non-zero = at least one mismatch (see report).

Why this exists: dashboard.html has its data baked into HTML/JS. The numbers were
transcribed accurately (audited 2026-05), but there is no automatic link back to
the source, so a future hand-edit could drift. This script is that link.
"""

import sys
from collections import defaultdict

import openpyxl

XLSX = "מכלול_אקסל_לרשות_אפריל_עם_שמות.xlsx"   # the "with names" workbook (named sheets)

# ── expected values currently hardcoded in dashboard.html (update if the design changes) ──
DASH = {
    "population":        26503,   # line 730
    "recipients":         4415,   # line 735  (anchored to the age breakdown / comparison sheet)
    "active_families":    1363,   # line 740
    "individuals_coded":  1821,   # line 741
    "uncoded":            2605,   # line 741
    "gross_alloc_millions": 35.77, # line 745
    "age":    [1832, 1804, 390, 389],                 # line 1138  (0-17,18-49,50-64,65+)
    "gender": [2214, 2201, 1],                        # line 1169  (F,M,other)
    "marital":[1347, 931, 196, 72, 26, 11],           # line 1203
    "trend_cur":  [21,15,12,23,25,20,35,21,20,12,11,5],  # line 1235  (May'25..Apr'26)
    "trend_prev": [9,11,22,18,31,24,22,24,26,28,10,8],   # line 1249  (May'24..Apr'25)
    # towns: name -> (pop, rec, fam, ses)   lines 1497-1511
    "towns": {
        "מעלה עמוס": (906,353,75,2), "בת עין": (1730,662,203,3), "אספר": (1228,342,64,3),
        "תקוע": (4326,767,266,6), "נוקדים": (3093,561,173,5), "כרמי צור": (979,162,48,5),
        "אלון שבות": (3046,470,191,7), "נווה דניאל": (2355,325,99,7), "אלעזר": (2615,324,94,7),
        "מגדל עוז": (572,53,14,3), "ראש צורים": (978,104,31,7), "קדר": (1649,131,44,7),
        "הר גילה": (1656,111,43,7), "כפר עציון": (1370,79,19,4),
    },
    # framework-type totals (top 9), line 1359
    "frames": {"מרכז טיפולי":182,"מועדון/מרכז קהילתי":177,"מעון יום":89,"מערך דיור":44,
               "טיפול יום":43,"מעטפת":40,"תעסוקה מוגנת":37,"סביבה תומכת":31,"פנימיה":29},
    # budget utilization (T36, Apr-2026 cumulative), lines 1583-1593: domain -> (alloc, used)
    "budget_util": {
        "שירותים לילד ולנוער":(3307658,1429641), "מנהל משה":(2671282,1113651),
        "שרותי שיקום":(2157121,480444), "רווחת הפרט והמשפחה":(1041640,364857),
        "שרותי תקון":(1027104,553909), "אזרחים ותיקים":(890386,243740),
        "פעילויות בקהילה":(410082,97686), "מינהל הרווחה":(114835,61448),
        "שירותים לעולים":(36133,0),
    },
}

results = []  # (ok: bool, message: str)
def check(ok, msg):
    results.append((bool(ok), msg))

def num(x):
    if isinstance(x, str): x = x.replace(",", "").strip()
    try: return float(x)
    except (TypeError, ValueError): return 0.0

def i(x):
    """Coerce a cell (possibly text like '1,730') to int."""
    return int(num(x))

def rows(ws, start=1):
    return list(ws.iter_rows(min_row=start, values_only=True))

wb = openpyxl.load_workbook(XLSX, data_only=True)

# ── 1. headline scalars ──────────────────────────────────────────────────────
comp = rows(wb["השוואה_קבוצות_אוכלוסייה"], 2)
total_pop = comp[0][4]; total_rec = comp[0][5]
check(total_pop == DASH["population"], f"population: dash {DASH['population']:,} vs excel {total_pop:,}")
check(total_rec == DASH["recipients"], f"recipients: dash {DASH['recipients']:,} vs excel {total_rec:,}")

fam = rows(wb["נזקקויות_משפחה_כלל"], 2)
active_fam = fam[0][5]
check(active_fam == DASH["active_families"], f"active families: dash {DASH['active_families']:,} vs excel {active_fam:,}")

ind = rows(wb["נזקקויות_פרט_כלל"], 2)
coded = ind[0][5]
check(coded == DASH["individuals_coded"], f"individuals w/ coded need: dash {DASH['individuals_coded']:,} vs excel {coded:,}")

bud_year = rows(wb["תקציב_שנתי_לפי_תחום"], 2)
gross = sum(num(r[1]) for r in bud_year)   # הקצאה כוללת ברוטו
check(abs(gross/1e6 - DASH["gross_alloc_millions"]) < 0.01,
      f"gross allocation: dash {DASH['gross_alloc_millions']}M vs excel {gross/1e6:.2f}M ({gross:,.0f})")

# ── 2. simple breakdown charts ───────────────────────────────────────────────
age = [r[1] for r in rows(wb["מקבלי_שירות_לפי_גיל"], 2)]
check(age == DASH["age"], f"age chart: dash {DASH['age']} vs excel {age}")

gender = [r[2] for r in rows(wb["מגדר"], 2)]
check(gender == DASH["gender"], f"gender chart: dash {DASH['gender']} vs excel {gender}")

marital = [r[2] for r in rows(wb["מצב_משפחתי"], 2)]
check(marital == DASH["marital"], f"marital chart: dash {DASH['marital']} vs excel {marital}")

cur = [r[4] for r in rows(wb["פתיחת_תיקים_2025_2026"], 2)]
check(cur == DASH["trend_cur"], f"trend current-year: dash {DASH['trend_cur']} vs excel {cur}")
prev = [r[4] for r in rows(wb["פתיחת_תיקים_2024_2025"], 2)]
check(prev == DASH["trend_prev"], f"trend previous-year: dash {DASH['trend_prev']} vs excel {prev}")

# ── 3. towns (pop, rec, fam, ses) ────────────────────────────────────────────
# columns: ...|fam(11)|rec(12)|pop(13)|ses(14)|name(15)
town_xl = {}
for r in rows(wb["ישובים_14_השוואה"], 2):
    town_xl[r[15]] = (i(r[13]), i(r[12]), i(r[11]), i(r[14]))
for name, exp in DASH["towns"].items():
    got = town_xl.get(name)
    check(got == exp, f"town {name}: dash {exp} vs excel {got}")

# ── 4. framework-type totals (C122) ──────────────────────────────────────────
ALIAS = {"מערך דיור/דירת לוין":"מערך דיור", "תעסוקה מוגנת )מעש(\"":"תעסוקה מוגנת"}
ftot = defaultdict(int)
for r in rows(wb["השמות_לפי_סוג_מסגרת"], 2):
    label = ALIAS.get(r[1], r[1])
    ftot[label] += int(num(r[0]))
for label, exp in DASH["frames"].items():
    check(ftot.get(label) == exp, f"framework {label}: dash {exp} vs excel {ftot.get(label)}")

# ── 5. budget utilization (T36, latest month) ────────────────────────────────
t36 = rows(wb["תקציב_מצטבר_חודשי"], 3)
target = max(r[19] for r in t36 if r[19] is not None)
alloc = defaultdict(float); used = defaultdict(float)
for r in t36:
    if r[19] == target:
        alloc[r[18]] += num(r[14]); used[r[18]] += num(r[11])
# excel domain spellings differ slightly from the dashboard labels; match by used+alloc
def find_domain(a, u):
    for d in alloc:
        if round(alloc[d]) == a and round(used[d]) == u:
            return d
    return None
for label, (a, u) in DASH["budget_util"].items():
    check(find_domain(a, u) is not None,
          f"budget {label}: dash alloc={a:,} used={u:,} — {'found in T36' if find_domain(a,u) else 'NO MATCH in T36'}")

# ── 6. internal consistency of the SOURCE (these are expected to disagree) ────
print("\nSOURCE internal-consistency (informational — known to differ in the raw data):")
age_total = sum(age); gender_total = sum(gender)
town_total = sum(v[1] for v in town_xl.values())
print(f"  total recipients via age    : {age_total:,}")
print(f"  total recipients via gender : {gender_total:,}  (Δ {gender_total-age_total:+})")
print(f"  total recipients via towns  : {town_total:,}  (Δ {town_total-age_total:+})")
print("  → dashboard headline uses the age/comparison anchor (4,415); the +29 town gap")
print("    and +1 gender gap originate in the welfare-system exports, not the dashboard.")

# ── report ───────────────────────────────────────────────────────────────────
print("\n" + "=" * 64)
passed = sum(1 for ok, _ in results if ok)
for ok, msg in results:
    if not ok:
        print(f"  FAIL  {msg}")
print(f"\n{passed}/{len(results)} checks passed.")
sys.exit(0 if passed == len(results) else 1)

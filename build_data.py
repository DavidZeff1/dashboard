#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — read the source Excel and generate data.js for dashboard.html.

Workflow after each monthly Excel refresh:
    python3 build_data.py        # regenerates data.js (+ prints a consistency report)

dashboard.html loads data.js (`<script src="data.js">`, which works under file://)
and reads everything from window.DASHDATA. No number is hand-typed in the HTML.

Presentation note: the dashboard uses hand-polished short labels and curated sub-captions
that don't exist verbatim in the Excel (e.g. "זקנה · אזרח ותיק" for the Excel's
"זקנה אזרח ותיק"). Those curated strings live in the *_DISPLAY / *_LABEL maps below — the
NUMBERS always come from the Excel; only the wording is curated here. If a future Excel
introduces a category not in a map, the raw Excel label is used and a warning is emitted.
"""

import json
import datetime
from collections import defaultdict

import openpyxl

XLSX = "מכלול_אקסל_לרשות_אפריל_עם_שמות.xlsx"
OUT = "data.js"
PERIOD = "אפריל 2026"

warnings = []

def num(x):
    if isinstance(x, str): x = x.replace(",", "").strip()
    try: return float(x)
    except (TypeError, ValueError): return 0.0

def i(x):
    return int(num(x))

# ── curated presentation maps (wording only; numbers come from Excel) ──────────
FAM_NEED_LABEL = {
    "מוגבלות": "מוגבלות",
    "זקנה אזרח ותיק": "זקנה · אזרח ותיק",
    "קשיים בין הורים לבין ילדיהם": "קשיים בין הורים לילדיהם",
    "אירוע ביטחון /חירום /אסונות טבע /מגיפה": "אירוע ביטחון / חירום",
    "קשיים ביחסים בין כלל בני המשפחה": "קשיים ביחסים בין בני המשפחה",
    "קשיים ביחסים שבין בני הזוג": "קשיים ביחסים בין בני הזוג",
    "מצוקה כלכלית זמנית": "מצוקה כלכלית זמנית",
    "אי בטחון תזונתי": "אי־ביטחון תזונתי",
    "ליקוי בתפקוד הורי לרבות קשיים בהשגחה והגנה": "ליקוי בתפקוד הורי",
    "קונפליקט בעוצמה גבוהה על רקע פרוד/גרושין": "קונפליקט גירושין בעוצמה",
}
IND_NEED_LABEL = {
    "מצוקה רגשית": "מצוקה רגשית",
    "קושי בטיפול אישי ובמיומנויות יומיומיות": "קושי בטיפול אישי ומיומנויות יומיומיות",
    "קשיים בצריכת שירותים ומיצוי זכויות": "קשיים בצריכת שירותים ומיצוי זכויות",
    "קשיים בהתפתחות הילד": "קשיים בהתפתחות הילד",
    "בדידות ו/או העדר מערכות תמיכה": "בדידות / היעדר מערכות תמיכה",
    "קשיים במסגרת תעסוקה (מגיל 18 ומעלה)": "קשיים במסגרת תעסוקה (18+)",
    "קשיים באינטראקציה בין אישית/זוגית": "קשיים באינטראקציה בין־אישית/זוגית",
    "קשיים במסגרת חינוכית (עד גיל 21)": "קשיים במסגרת חינוכית (עד 21)",
    "התנהגות סיכונית אחרת": "התנהגות סיכונית אחרת",
    "חשד או נתון/ה לפגיעה מינית": "חשד או נתון/ה לפגיעה מינית",
}
# excel domain -> (alloc-chart label, util name, util sub-caption)
BUDGET_DISPLAY = {
    "משה\"":                ("מנהל משה (מוגבלויות)", "מנהל משה",           "מוגבלויות"),
    "שרותים לילד ולנוער":   ("שירותים לילד ולנוער",  "שירותים לילד ולנוער", "תיקי טיפול"),
    "מינהל הרווחה":         ("מינהל הרווחה",         "מינהל הרווחה",        "תקורה"),
    "שרותי שיקום":          ("שרותי שיקום",          "שרותי שיקום",         "מסגרות שיקום"),
    "שרותי תקון":           ("שרותי תקון",           "שרותי תקון",          "תיקון, סדנאות"),
    "רווחת הפרט והמשפחה":   ("רווחת הפרט והמשפחה",   "רווחת הפרט והמשפחה",  "תחנות, מרכזים"),
    "אזרחים ותיקים":        ("אזרחים ותיקים",        "אזרחים ותיקים",       "גיל שלישי"),
    "פעליות בקהילה":        ("פעילויות בקהילה",      "פעילויות בקהילה",     "פעולות שגרתיות"),
    "שרותים לעולים":        ("שירותים לעולים",       "שירותים לעולים",      "תוכניות עליה"),
}
COMPARE_DISPLAY = {
    "מקבלי שירות": ("מקבלי שירות (כלל)", "default"),
    "אזרחים ותיקים": ("אזרחים ותיקים", "crimson"),
    "ילדים 0-18": ("ילדים 0–18", "default"),
    "עולים חדשים": ("עולים חדשים", "olive"),
}
FRAME_LABEL = {
    "מערך דיור/דירת לוין": "מערך דיור",
    "תעסוקה מוגנת )מעש(\"": "תעסוקה מוגנת",
}
# explicit unit order + display for the inside/outside chart (curated order)
INOUT_ORDER = [
    ("ילד ונוער", "ילד ונוער"), ("אזרחים ותיקים", "אזרחים ותיקים"),
    ("רווחת הפרט ומש`", "רווחת הפרט"), ("שיקום ונכים", "שיקום ונכים"),
    ("אגף משה", "אגף משה"), ("נפגעי התמכרויות", "נפגעי התמכרויות"),
    ("נוער וצעירים", "נוער וצעירים"), ("אוטיסטים", "אוטיסטים"),
    ("נשים ונערות", "נשים ונערות"),
]

def relabel(raw, mp, what):
    if raw in mp: return mp[raw]
    # normalize stray cp1252 dashes/control chars that appear in some exports
    norm = " ".join((raw or "").replace("\x96", " ").replace("\x97", " ").split())
    if norm in mp: return mp[norm]
    warnings.append(f"no curated label for {what}: {raw!r} (using raw)")
    return raw

# ── load ──────────────────────────────────────────────────────────────────────
wb = openpyxl.load_workbook(XLSX, data_only=True)
def rows(name, start): return list(wb[name].iter_rows(min_row=start, values_only=True))

D = {}

# headline scalars + comparison table
comp = rows("השוואה_קבוצות_אוכלוסייה", 2)
D["population"]   = i(comp[0][4])
D["recipients"]   = i(comp[0][5])
fam = rows("נזקקויות_משפחה_כלל", 2)
D["activeFamilies"] = i(fam[0][5])
ind = rows("נזקקויות_פרט_כלל", 2)
D["individualsCoded"] = i(ind[0][5])
# uncoded note: " ...לא זוהתה נזקקות' - 2,605  (59.00% "
import re
note = ind[0][6] or ""
m = re.search(r"([\d,]+)\s*\(([\d.]+)%", note)
D["individualsUncoded"]    = i(m.group(1)) if m else 0
D["individualsUncodedPct"] = float(m.group(2)) if m else 0.0
chnote = rows("מאפייני_פרט_כלל", 2)[0][5] or ""
mc = re.search(r"\(([\d.]+)%\)\s*([\d,]+)", chnote)
D["charsUncodedPct"] = float(mc.group(1)) if mc else 0.0
D["charsUncoded"]    = i(mc.group(2)) if mc else 0

# budget — annual gross by domain (T33)
bud_year = rows("תקציב_שנתי_לפי_תחום", 2)
gross = sum(num(r[1]) for r in bud_year)
D["grossAlloc"] = round(gross, 2)
alloc_list = []
for r in bud_year:
    label = relabel(r[3], {k: v[0] for k, v in BUDGET_DISPLAY.items()}, "budget domain")
    alloc_list.append([label, int(num(r[1]))])
alloc_list.sort(key=lambda x: -x[1])
D["budgetAlloc"] = alloc_list

# comparison rows
cmp_rows = []
for r in comp:
    disp, color = COMPARE_DISPLAY.get(r[1], (r[1], "default"))
    cmp_rows.append({"group": disp, "pop": i(r[4]), "rec": i(r[5]),
                     "national": float(r[7]), "similar": float(r[8]), "color": color})
D["compare"] = cmp_rows

# demographics
age = rows("מקבלי_שירות_לפי_גיל", 2)
D["age"] = {"labels": ["0—17", "18—49", "50—64", "65+"], "data": [i(r[1]) for r in age]}
gen = rows("מגדר", 2)
D["gender"] = {"labels": ["נשים", "גברים", "אחר"], "data": [i(r[2]) for r in gen]}
mar = rows("מצב_משפחתי", 2)
D["marital"] = {"labels": ["נשוי/ה","רווק/ה","גרוש/ה","אלמן/ה","חי/ה בנפרד","ידוע/ה בציבור"],
                "data": [i(r[2]) for r in mar]}

# towns
town = []
for r in rows("ישובים_14_השוואה", 2):
    town.append({"name": r[15], "pop": i(r[13]), "rec": i(r[12]), "fam": i(r[11]), "ses": i(r[14])})
D["towns"] = town

# needs (top 10 by count, relabeled)
def top_needs(sheet, denom, labelmap, n=10):
    data = [(r[2], i(r[3])) for r in rows(sheet, 2)]
    data.sort(key=lambda x: -x[1])
    out = []
    for raw, c in data[:n]:
        out.append([relabel(raw, labelmap, sheet), c, round(c / denom * 100, 2)])
    return out
D["famNeeds"] = top_needs("נזקקויות_משפחה_כלל", D["activeFamilies"], FAM_NEED_LABEL)
D["indNeeds"] = top_needs("נזקקויות_פרט_כלל", D["individualsCoded"], IND_NEED_LABEL)

# trend (case openings)
tcur = rows("פתיחת_תיקים_2025_2026", 2)
tprev = rows("פתיחת_תיקים_2024_2025", 2)
D["trend"] = {
    "labels": ["מאי","יוני","יולי","אוג׳","ספט׳","אוק׳","נוב׳","דצמ׳","ינו׳","פבר׳","מרץ","אפר׳"],
    "cur":  [i(r[4]) for r in tcur],
    "prev": [i(r[4]) for r in tprev],
}

# budget utilization (T36 latest month, by domain)
t36 = rows("תקציב_מצטבר_חודשי", 3)
target = max(r[19] for r in t36 if r[19] is not None)
a_sum = defaultdict(float); u_sum = defaultdict(float)
for r in t36:
    if r[19] == target:
        a_sum[r[18]] += num(r[14]); u_sum[r[18]] += num(r[11])
# T36 spells some domains differently from the annual sheet (e.g. "משה" vs 'משה"'),
# so match on a normalized key (strip quotes + collapse whitespace).
def dkey(s): return " ".join((s or "").replace('"', "").split())
disp_by_norm = {dkey(dom): (name, sub) for dom, (cl, name, sub) in BUDGET_DISPLAY.items()}
util = []
for dom in a_sum:
    nd = dkey(dom)
    if nd in disp_by_norm:
        name, sub = disp_by_norm[nd]
        util.append({"name": name, "sub": sub, "alloc": round(a_sum[dom]), "used": round(u_sum[dom])})
    else:
        warnings.append(f"no curated budget display for T36 domain: {dom!r}")
util.sort(key=lambda x: -x["alloc"])
D["budgetRows"] = util

# frameworks: top-9 types (C122) + inside/outside by unit
c122 = rows("השמות_לפי_סוג_מסגרת", 2)
ftot = defaultdict(int)
for r in c122:
    ftot[relabel(r[1], FRAME_LABEL, "frame type") if r[1] in FRAME_LABEL else r[1]] += i(r[0])
ftop = sorted(ftot.items(), key=lambda x: -x[1])[:9]
D["frames"] = {"labels": [k for k, _ in ftop], "data": [v for _, v in ftop]}

inside = defaultdict(int); outside = defaultdict(int)
for r in c122:
    inside[r[3]] += i(r[5]); outside[r[3]] += i(r[6])
D["inout"] = {
    "labels":  [disp for _, disp in INOUT_ORDER],
    "inside":  [inside.get(raw, 0) for raw, _ in INOUT_ORDER],
    "outside": [outside.get(raw, 0) for raw, _ in INOUT_ORDER],
}

# in-brief stats
goals = rows("יעדי_טיפול", 2)
D["topGoal"] = max(i(r[2]) for r in goals)
D["goalsCount"] = len(goals)
t30 = rows("מסגרות_ברשות_הנבחרת", 2)
D["frameworksCount"] = len(t30)
D["fosterFamilies"] = sum(1 for r in t30 if r[9] == "מש` אומנה")
D["crossAuthPlacements"] = len(rows("מסגרות_שמות_מלאים", 2))

# meta
D["meta"] = {
    "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
    "sourceFile": XLSX,
    "period": PERIOD,
    "warnings": warnings,
}

# ── internal-consistency check (source data is known to disagree slightly) ─────
age_t = sum(D["age"]["data"]); gen_t = sum(D["gender"]["data"])
town_t = sum(t["rec"] for t in D["towns"])
consistency = []
if gen_t != D["recipients"]:
    consistency.append(f"gender total {gen_t:,} ≠ headline recipients {D['recipients']:,} (Δ{gen_t-D['recipients']:+})")
if town_t != D["recipients"]:
    consistency.append(f"town recipients sum {town_t:,} ≠ headline recipients {D['recipients']:,} (Δ{town_t-D['recipients']:+})")
D["meta"]["consistency"] = consistency

# ── write data.js ─────────────────────────────────────────────────────────────
with open(OUT, "w", encoding="utf-8") as fh:
    fh.write("/* AUTO-GENERATED by build_data.py — do not edit by hand. */\n")
    fh.write("/* Re-run `python3 build_data.py` after updating the Excel.   */\n")
    fh.write("window.DASHDATA = ")
    json.dump(D, fh, ensure_ascii=False, indent=2)
    fh.write(";\n")

# ── report ────────────────────────────────────────────────────────────────────
print(f"wrote {OUT}  (source: {XLSX}, period {PERIOD})")
print(f"  population={D['population']:,}  recipients={D['recipients']:,}  "
      f"families={D['activeFamilies']:,}  gross={D['grossAlloc']/1e6:.2f}M")
if warnings:
    print("\nLABEL WARNINGS (curation maps may need updating):")
    for w in warnings: print("  •", w)
print("\nSOURCE internal-consistency (these gaps live in the welfare-system export, not the dashboard):")
print(f"  recipients via age={age_t:,}  gender={gen_t:,}  towns={town_t:,}")
for c in consistency: print("  ⚠ ", c)
if not consistency: print("  ✓ all recipient totals agree")

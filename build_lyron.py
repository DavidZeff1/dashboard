#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_lyron.py — generate data_lyron.js for dashboard_lyron.html (the strategy memo).

    python3 build_lyron.py

Source: מכלול_אקסל_לרשות_אפריל_עם_שמות.xlsx. Derived sections:
  • cost-per-case by unit + the community/residential ratio  → T26 (תקציב_לפי_מסגרת_וסעיף)
  • SES clusters                                              → ישובים_14_השוואה (grouped by cluster)
  • prevention goals                                          → יעדי_טיפול (curated 9-goal subset)
  • need cards                                                → מאפייני_פרט_כלל + נזקקויות_פרט + comparison
  • individual-assistance categories                          → תשלומי_עזרה_לפרטים (T39, curated top set)
  • footer KPIs                                               → comparison + budget + frameworks sheets
Curated wording/colors live in the maps below; numbers come from the Excel.

ONE EXCEPTION — the "projects" panel (PROJS) is a hand-picked illustrative sample of individual
line items from תשלומי_פרויקטים (the dashboard labels them "דוגמאות בולטות"/sample). Those rows
are editorial selections with custom labels, so they are kept as a curated constant here rather
than auto-aggregated.
"""

import json, datetime, math
from collections import defaultdict

import openpyxl

XLSX = "מכלול_אקסל_לרשות_אפריל_עם_שמות.xlsx"
OUT = "data_lyron.js"
warnings = []

def num(x):
    if isinstance(x, str): x = x.replace(",", "").strip()
    try: return float(x)
    except (TypeError, ValueError): return 0.0
def i(x): return int(num(x))
def jr(x, d=0):
    f = 10 ** d; return math.floor(x * f + 0.5) / f

# ── curated presentation ──────────────────────────────────────────────────────
COST_KIND = {  # unit -> kind label (residential vs community character)
    "חסות הנוער": "חוץ ביתי", "אוטיסטים": "חוץ ביתי", "אגף משה": "חוץ ביתי",
    "השרות לעיוור": "חוץ ביתי", "שיקום ונכים": "חוץ ביתי", "שיקום הנוער": "בקהילה",
    "נשים ונערות": "חוץ ביתי", "נוער וצעירים": "בקהילה", "ילד ונוער": "חוץ ביתי",
    "רווחת הפרט ומש`": "בקהילה", "נפגעי התמכרויות": "בקהילה", "אזרחים ותיקים": "חוץ ביתי",
}
SES_CLS = {2: "lo", 3: "mid-lo", 4: "mid-lo", 5: "mid", 6: "mid-hi", 7: "hi"}

GOALS_SELECT = [  # full label in יעדי_טיפול -> short display (curated subset, in display order)
    ("הגברת הנגישות לשירותים ומשאבים", "הגברת הנגישות לשירותים ומשאבים"),
    ("הגברת הידע והמידע אודות שירותים ומשאבים", "הגברת הידע על שירותים ומשאבים"),
    ("הקטנת חסמים/ קשיים/ שילוב במסגרת תעסוקה", "שילוב במסגרת תעסוקה"),
    ("הקטנת חסמים/ קשיים/ שילוב במסגרת חינוכית", "שילוב במסגרת חינוכית"),
    ("הגברת מצאי המשאבים והקשרים החברתיים", "הגברת מצאי משאבים וקשרים חברתיים"),
    ("הקטנת רמת סיכון בהתמודדות על רקע התנהגויות סיכוניות", "הפחתת התנהגויות סיכוניות"),
    ("הפחתת הקונפליקטים עם מערכות חברתיות", "הפחתת קונפליקטים עם מערכות חברתיות"),
    ("שיפור והקטנת רמת סיכון בהתמודדות עם קשיים על רקע בעיית התמכרויות בהיבטים הפיזיים, רגשיים, קוגניטיביים וחברתיים", "התמודדות עם התמכרויות"),
    ("שילוב ועידוד לתוכניות לבוגרים צעירים (צבא, שירות לאומי, מכינות,ישיבות)", "תוכניות לבוגרים צעירים — צבא, ש״ל, מכינות"),
]
ASSIST_SELECT = [  # raw סוג עזרה in T39 -> display (curated top set, in display order)
    ("הוצאות נסיעה/הסעה", "הוצאות נסיעה והסעה"),
    ("הוצאות נסיעה/הסעה לעולה", "הוצאות נסיעה והסעה לעולים"),
    ("הוצאות נסיעה למסגרת", "הוצאות נסיעה למסגרות"),
    ("יתד-שיפור תפקוד, בריאות ומשק בית", "יתד · שיפור תפקוד ובריאות"),
    ("ליווי", "ליווי במסגרות"),
    ("אבחונים לילדים", "אבחונים לילדים"),
    ("סיוע לימודי לילדים ובני נוער", "סיוע לימודי לילדים ובני נוער"),
    ("פעילות פנאי והעשרה לילדים", "פעילות פנאי והעשרה לילדים"),
    ("טיפולים לילדים", "טיפולים לילדים"),
    ("יתד-דמי כיס וצורכי קיום", "יתד · דמי כיס וצרכי קיום"),
]
# curated illustrative sample (see module docstring) — [name, amount, source-label]
PROJS = [
    ["הסעות לפעילות קבוצתית", 82881, "הסעות במועצות איזורית"],
    ["סדנאות לכידות חברתית", 85634, "תקציב לכידות"],
    ["הסעות למסגרות", 16693, "הסעות במועצות"],
    ["הרצאות לכידות", 45822, "תקציב לכידות"],
    ["טיפול קבוצתי ילד ונוער", 49025, "טיפול בילד בקהילה"],
    ["הרצאות ילד ונוער", 23174, "טיפול בילד בקהילה"],
    ["הסעות מ.יום שיקומי", 53681, "הסעות שיקום"],
]

wb = openpyxl.load_workbook(XLSX, data_only=True)
def rows(name, start): return list(wb[name].iter_rows(min_row=start, values_only=True))
D = {}

# ── cost per case by unit (T26) ───────────────────────────────────────────────
t26 = rows("תקציב_לפי_מסגרת_וסעיף", 2)  # 0=amount 4=placements 9=arrangement 10=unit
amt_u = defaultdict(float); plc_u = defaultdict(float)
amt_a = defaultdict(float); plc_a = defaultdict(float)
for r in t26:
    amt_u[r[10]] += num(r[0]); plc_u[r[10]] += num(r[4])
    amt_a[r[9]]  += num(r[0]); plc_a[r[9]]  += num(r[4])
cost = []
for u in amt_u:
    n = round(plc_u[u]); tot = round(amt_u[u])
    if u not in COST_KIND: warnings.append(f"no kind for cost unit {u!r}")
    cost.append([u, COST_KIND.get(u, "בקהילה"), round(tot / n) if n else 0, n, tot])
cost.sort(key=lambda c: -c[2])
D["costData"] = cost
# community / residential ratio (by arrangement)
comm_n = round(plc_a.get("בקהילה", 0)); comm_amt = round(amt_a.get("בקהילה", 0))
resi_n = round(plc_a.get("חוץ ביתי", 0)); resi_amt = round(amt_a.get("חוץ ביתי", 0))
D["ratio"] = {
    "commN": comm_n, "commAmt": comm_amt, "commPer": round(comm_amt / comm_n) if comm_n else 0,
    "resiN": resi_n, "resiAmt": resi_amt, "resiPer": round(resi_amt / resi_n) if resi_n else 0,
    "totalPlacements": comm_n + resi_n,
}
D["ratio"]["x"] = jr(D["ratio"]["resiPer"] / D["ratio"]["commPer"], 1) if D["ratio"]["commPer"] else 0

# ── SES clusters (towns grouped by cluster) ───────────────────────────────────
towns = []
for r in rows("ישובים_14_השוואה", 2):
    towns.append({"name": r[15], "pop": i(r[13]), "rec": i(r[12]), "ses": i(r[14])})
clusters = defaultdict(lambda: {"towns": [], "pop": 0, "rec": 0})
for t in towns:
    c = clusters[t["ses"]]; c["towns"].append(t["name"]); c["pop"] += t["pop"]; c["rec"] += t["rec"]
ses = []
for s in sorted(clusters):
    c = clusters[s]
    ses.append({"ses": s, "towns": c["towns"], "pop": c["pop"], "rec": c["rec"], "cls": SES_CLS.get(s, "mid")})
D["sesData"] = ses
totPop = sum(t["pop"] for t in towns); totRec = sum(t["rec"] for t in towns)

# ── prevention goals (curated subset of יעדי_טיפול) ───────────────────────────
goalcnt = {}
for r in rows("יעדי_טיפול", 2):
    goalcnt[(r[1] or "").strip()] = i(r[2])
goals = []
for full, short in GOALS_SELECT:
    if full not in goalcnt: warnings.append(f"goal not found: {full!r}")
    goals.append([goalcnt.get(full, 0), short])
goals.sort(key=lambda g: -g[0])
D["goals"] = goals
D["goalsTotalAll"] = sum(goalcnt.values())   # all 28 goals (section IV meta context)

# ── need cards (characteristics + needs) ──────────────────────────────────────
char = {}
for r in rows("מאפייני_פרט_כלל", 2):
    if r[2]: char[(r[2] or "").strip()] = i(r[3])
indneed = {}
for r in rows("נזקקויות_פרט_כלל", 2):
    indneed[(r[2] or "").strip()] = i(r[3])
comp = rows("השוואה_קבוצות_אוכלוסייה", 2)
vatikim_rec = i(comp[1][5])  # אזרחים ותיקים recipients
D["needCards"] = {
    "vatikim": vatikim_rec,
    "asd": char.get("הפרעות מאובחנות ברצף האוטיסטי (ASD)", 0),
    "intellectual": char.get("מוגבלות שכלית התפתחותית מאובחנת", 0),
    "bereavement": char.get("אובדן ושכול", 0),
    "bereavementSource": char.get("מקור השכול", 0),
    "sexual": indneed.get("חשד או נתון/ה לפגיעה מינית", 0),
    "sexualChar": char.get("נפגעי תקיפה מינית", 0),
    "complexFamilies": char.get("משפחה מורכבת", 0),
}

# ── assistance (T39, curated top set) ─────────────────────────────────────────
a_amt = defaultdict(float); a_rec = defaultdict(float)
for r in rows("תשלומי_עזרה_לפרטים", 2):  # 0=amount 2=recipients 3=type
    k = (r[3] or "").strip(); a_amt[k] += num(r[0]); a_rec[k] += num(r[2])
assist = []
for raw, disp in ASSIST_SELECT:
    if raw not in a_amt: warnings.append(f"assist type not found: {raw!r}")
    assist.append([disp, round(a_rec.get(raw, 0)), round(a_amt.get(raw, 0))])
D["assist"] = assist
D["projs"] = PROJS

# ── footer KPIs ───────────────────────────────────────────────────────────────
gross = sum(num(r[1]) for r in rows("תקציב_שנתי_לפי_תחום", 2))
# budget utilization (T36 latest month) — total used / total alloc
t36 = rows("תקציב_מצטבר_חודשי", 3)
target = max(r[19] for r in t36 if r[19] is not None)
u_alloc = u_used = 0.0
for r in t36:
    if r[19] == target: u_alloc += num(r[14]); u_used += num(r[11])
# new family cases over last 4 months
trend = [i(r[4]) for r in rows("פתיחת_תיקים_2025_2026", 2)]
D["kpi"] = {
    "recipients": i(comp[0][5]), "population": i(comp[0][4]),
    "ratePct": jr(i(comp[0][5]) / i(comp[0][4]) * 100, 2),
    "nationalPct": jr(comp[0][7] * 100, 2), "similarPct": jr(comp[0][8] * 100, 2),
    "activeFamilies": i(rows("נזקקויות_משפחה_כלל", 2)[0][5]),
    "newCases4mo": sum(trend[-4:]),
    "frameworks": len(rows("מסגרות_ברשות_הנבחרת", 2)),
    "crossPlacements": len(rows("מסגרות_שמות_מלאים", 2)),
    "grossM": jr(gross / 1e6, 2),
    "utilUsedM": jr(u_used / 1e6, 2), "utilAllocM": jr(u_alloc / 1e6, 2),
    "utilPct": jr(u_used / u_alloc * 100, 1),
    "vatikimRatePct": jr(vatikim_rec / i(comp[1][4]) * 100, 2),
    "vatikimSimilarPct": jr(comp[1][8] * 100, 2),
    "individualsCoded": i(rows("נזקקויות_פרט_כלל", 2)[0][5]),
}
# SES cluster 7 highlight
ses7 = next(s for s in ses if s["ses"] == 7)
D["ses7"] = {"pop": ses7["pop"], "rec": ses7["rec"],
             "popPct": jr(ses7["pop"] / totPop * 100, 1),
             # dashboard expresses cluster-7 share against the headline recipients (4,415),
             # not the town-breakdown sum (4,444 — the known source gap), so match that.
             "recPct": jr(ses7["rec"] / D["kpi"]["recipients"] * 100, 1),
             "ratePct": jr(ses7["rec"] / ses7["pop"] * 100, 1)}

D["meta"] = {"generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
             "sourceFile": XLSX, "warnings": warnings}

with open(OUT, "w", encoding="utf-8") as fh:
    fh.write("/* AUTO-GENERATED by build_lyron.py — do not edit by hand. */\n")
    fh.write("window.DASHDATA = ")
    json.dump(D, fh, ensure_ascii=False, indent=2)
    fh.write(";\n")

print(f"wrote {OUT}")
print(f"  cost units={len(cost)}  ratio={D['ratio']['commPer']}/{D['ratio']['resiPer']} (×{D['ratio']['x']})  placements={D['ratio']['totalPlacements']}")
print(f"  ses clusters={len(ses)}  goals={len(goals)}  assist={len(assist)}  goalsTotalAll={D['goalsTotalAll']}")
print(f"  needCards={D['needCards']}")
print(f"  kpi recipients={D['kpi']['recipients']} util={D['kpi']['utilPct']}% newCases4mo={D['kpi']['newCases4mo']}")
if warnings:
    print("WARNINGS:"); [print("  •", w) for w in warnings]

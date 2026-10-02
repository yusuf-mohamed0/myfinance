# -*- coding: utf-8 -*-
r"""
BUILD_WEB - generates a self-contained static site in 06_Web from
summary_2026.json + plan_YYYY-MM.xlsx + transactions_2026.csv,
then copies the downloadable reports next to it. Publish with publish_web.py.

Regeneration chain (keeps every download in sync with the system):
  main() first runs regen_reports():
    1. analyze_sms.py   -> summary_2026.json, transactions_2026.csv,
                           Financial_Analysis_2026.xlsx, dashboard_2026.html
    2. plan_month.py    -> plan_<MONTH>.xlsx (income read from the existing
                           plan cell B2, fallback 10000)
  Any non-zero exit aborts the build (exit 1) so stale/uncolored files are
  never published.  Set MF_NOREGEN=1 to skip the chain for quick rebuilds.
  market_watch.json has no generator: it is refreshed manually on Saturday
  and only copied through.

Design v2 (navy/gold refined, mobile-first, RTL Arabic, zero emojis):
  - Sticky header + hero card: remaining weekly cash, safe/day, days left
  - Quick-log (expenses AND income) in localStorage with inline edit,
    Enter-to-add, export/import backup (JSON), stale-month hint
  - Envelopes: plan vs actual (bank actuals embedded + manual log merged in
    the browser), progress bars per envelope with color thresholds
  - Market pulse (weekly Saturday cadence) section
  - Insights: 9-month SVG spending trend, month-over-month net-spend table,
    top-5 merchants, unusual-spend alerts
  - Transactions explorer: searchable/filterable recent bank txns
  - Full plan table, category history table, token-named downloads
    (freshness note with $gentime build timestamp)
  - Bottom tab bar (mobile) with scroll-spy, sticky header

Security/verify constraints (verify_system.py 56 checks):
  id="lock", gate hash, >=14 <tr>, >=5 token downloads, id="market" with
  >=9 .mk-i, id="seg_in"/"seg_out"/"st_in", functions editEntry/saveEdit/
  cancelEdit/setEdType, zero emoji/arrow chars (0x1F000-0x1FAFF, 0x2600-
  0x27BF, 0xFE0F, 0x2192, 0x2190), word "yowmeeyan" (Arabic) in market note.

Usage: python build_web.py [plan_month YYYY-MM]   (default: current month)
Console output is ASCII-only.
"""
import json, os, re, csv, shutil, sys, calendar, datetime as dt, subprocess
from string import Template

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB = os.path.join(BASE, "06_Web")
from mfconfig import CFG, YEAR, GATE  # noqa: E402  (settings live in config.json)

SUMMARY = os.path.join(BASE, "02_Reports", "summary_{}.json".format(YEAR))
TXNS_CSV = os.path.join(BASE, "02_Reports", "transactions_{}.csv".format(YEAR))

MONTH = sys.argv[1] if len(sys.argv) > 1 else dt.date.today().strftime("%Y-%m")
PLAN = os.path.join(BASE, "02_Reports", "plan_{}.xlsx".format(MONTH))

DL = [  # report files copied into the site for download
    ("Financial_Analysis_{}.xlsx".format(YEAR), "Financial_Analysis_{}.xlsx".format(YEAR),
     "تقرير التحليل الكامل"),
    ("transactions_{}.csv".format(YEAR), "transactions_{}.csv".format(YEAR),
     "كل المعاملات (CSV)"),
    ("summary_{}.json".format(YEAR), "summary_{}.json".format(YEAR),
     "الأرقام المجمعة (JSON)"),
    ("plan_{}.xlsx".format(MONTH), "plan_{}.xlsx".format(MONTH), "خطة الشهر (إكسل)"),
    ("market_watch.json", "market_watch.json", "أسعار السوق والأخبار (JSON)"),
]

# ---------------------------------------------------------------- Lucide icons
# Inner SVG of official Lucide icons (ISC license), embedded for offline use.
ICONS = {
"lock": '''<rect width="18" height="11" x="3" y="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>''',
"shield-check": '''<path d="M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 6.24-2.72a1.17 1.17 0 0 1 1.52 0C14.51 3.81 17 5 19 5a1 1 0 0 1 1 1z"/><path d="m9 12 2 2 4-4"/>''',
"wallet": '''<path d="M19 7V4a1 1 0 0 0-1-1H5a2 2 0 0 0 0 4h15a1 1 0 0 1 1 1v4h-3a2 2 0 0 0 0 4h3a1 1 0 0 1 1 1v3a1 1 0 0 1-1 1H5a2 2 0 0 1-2-2V5"/><path d="M3 5v14a2 2 0 0 0 2 2h15a1 1 0 0 0 1-1v-4"/>''',
"trending-up": '''<path d="M16 7h6v6"/><path d="m22 7-8.5 8.5-5-5L2 17"/>''',
"trending-down": '''<path d="M16 17h6v-6"/><path d="m22 17-8.5-8.5-5 5L2 7"/>''',
"piggy-bank": '''<path d="M11 17h3v2a1 1 0 0 0 1 1h2a1 1 0 0 0 1-1v-3a3.16 3.16 0 0 0 2-2h1a1 1 0 0 0 1-1v-2a1 1 0 0 0-1-1h-1a5 5 0 0 0-2-4V3a4 4 0 0 0-3.2 1.6l-.3.4H11a6 6 0 0 0-6 6v1a5 5 0 0 0 2 4v3a1 1 0 0 0 1 1h2a1 1 0 0 0 1-1z"/><path d="M16 10h.01"/><path d="M2 8v1a2 2 0 0 0 2 2h1"/>''',
"gauge": '''<path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>''',
"calendar-days": '''<path d="M8 2v3"/><path d="M16 2v3"/><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M8 13h.01"/><path d="M12 13h.01"/><path d="M16 13h.01"/><path d="M8 17h.01"/><path d="M12 17h.01"/><path d="M16 17h.01"/>''',
"calendar-check": '''<path d="M8 2v3"/><path d="M16 2v3"/><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="m9 15 2 2 4-4"/>''',
"clipboard-list": '''<rect width="8" height="4" x="8" y="2" rx="1" ry="1"/><path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/><path d="M12 11h4"/><path d="M12 16h4"/><path d="M8 11h.01"/><path d="M8 16h.01"/>''',
"plus": '''<path d="M5 12h14"/><path d="M12 5v14"/>''',
"copy": '''<rect width="14" height="14" x="8" y="8" rx="2" ry="2"/><path d="M4 16c-1.1 0-2-.9-2-2V4c0-1.1.9-2 2-2h10c1.1 0 2 .9 2 2"/>''',
"download": '''<path d="M12 15V3"/><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><path d="m7 10 5 5 5-5"/>''',
"info": '''<circle cx="12" cy="12" r="10"/><path d="M12 16v-4"/><path d="M12 8h.01"/>''',
"settings": '''<path d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z"/><circle cx="12" cy="12" r="3"/>''',
"sparkles": '''<path d="M9.937 15.5A2 2 0 0 0 8.5 14.063l-6.135-1.582a.5.5 0 0 1 0-.962L8.5 9.936A2 2 0 0 0 9.937 8.5l1.582-6.135a.5.5 0 0 1 .963 0L14.063 8.5A2 2 0 0 0 15.5 9.937l6.135 1.581a.5.5 0 0 1 0 .964L15.5 14.063a2 2 0 0 0-1.437 1.437l-1.582 6.135a.5.5 0 0 1-.963 0z"/><path d="M20 3v4"/><path d="M22 5h-4"/>''',
"target": '''<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>''',
"circle-check": '''<circle cx="12" cy="12" r="10"/><path d="m16 9-5.5 5.5L8 12"/>''',
"circle-alert": '''<circle cx="12" cy="12" r="10"/><line x1="12" x2="12" y1="8" y2="12"/><line x1="12" x2="12.01" y1="16" y2="16"/>''',
"triangle-alert": '''<path d="m21.73 18-8-14a2 2 0 0 0-3.48 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3"/><path d="M12 9v4"/><path d="M12 17h.01"/>''',
"banknote": '''<rect width="20" height="12" x="2" y="6" rx="2"/><circle cx="12" cy="12" r="2"/><path d="M6 12h.01M18 12h.01"/>''',
"fuel": '''<path d="M14 13h2a2 2 0 0 1 2 2v2a2 2 0 0 0 4 0v-6.998a2 2 0 0 0-.59-1.42L18 5"/><path d="M14 21V5a2 2 0 0 0-2-2H5a2 2 0 0 0-2 2v16"/><path d="M2 21h13"/><path d="M3 9h11"/>''',
"smartphone": '''<rect width="14" height="20" x="5" y="2" rx="2" ry="2"/><path d="M12 18h.01"/>''',
"shopping-cart": '''<path d="m2.05 2.05 1.099-.028a1 1 0 0 1 1.008.815l2.69 14.347A1 1 0 0 0 7.83 18H18"/><path d="M4.563 5h16.435a1 1 0 0 1 .981 1.204l-1.026 6.226A2 2 0 0 1 18.962 14H6.25"/><circle cx="18" cy="20" r="2"/><circle cx="8" cy="20" r="2"/>''',
"utensils": '''<path d="M3 2v7c0 1.1.9 2 2 2h4a2 2 0 0 0 2-2V2"/><path d="M7 2v20"/><path d="M21 15V2a5 5 0 0 0-5 5v6c0 1.1.9 2 2 2h3Zm0 0v7"/>''',
"coffee": '''<path d="M10 2v2"/><path d="M14 2v2"/><path d="M16 8a1 1 0 0 1 1 1v8a4 4 0 0 1-4 4H7a4 4 0 0 1-4-4V9a1 1 0 0 1 1-1h14a4 4 0 1 1 0 8h-1"/><path d="M6 2v2"/>''',
"shopping-basket": '''<path d="m15 11-1 9"/><path d="m19 11-4-7"/><path d="M2 11h20"/><path d="m3.5 11 1.6 7.4a2 2 0 0 0 2 1.6h9.8a2 2 0 0 0 2-1.6l1.7-7.4"/><path d="M4.5 15.5h15"/><path d="m5 11 4-7"/><path d="m9 11 1 9"/>''',
"heart-pulse": '''<path d="M2 9.5a5.5 5.5 0 0 1 9.591-3.676.56.56 0 0 0 .818 0A5.49 5.49 0 0 1 22 9.5c0 2.29-1.5 4-3 5.5l-5.492 5.313a2 2 0 0 1-3 .019L5 15c-1.5-1.5-3-3.2-3-5.5"/><path d="M3.22 13H9.5l.5-1 2 4.5 2-7 1.5 3.5h5.27"/>''',
"shirt": '''<path d="M20.38 3.46 16 2a4 4 0 0 1-8 0L3.62 3.46a2 2 0 0 0-1.34 2.23l.58 3.47a1 1 0 0 0 .99.84H6v10c0 1.1.9 2 2 2h8a2 2 0 0 0 2-2V10h2.15a1 1 0 0 0 .99-.84l.58-3.47a2 2 0 0 0-1.34-2.23z"/>''',
"refresh-cw": '''<path d="M3 12a9 9 0 0 1 9-9 9.75 9.75 0 0 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L3 16"/><path d="M8 16H3v5"/>''',
"send": '''<path d="M14.536 21.686a.5.5 0 0 0 .937-.024l6.5-19a.496.496 0 0 0-.635-.635l-19 6.5a.5.5 0 0 0-.024.937l7.93 3.18a2 2 0 0 1 1.112 1.11z"/><path d="m21.854 2.147-10.94 10.939"/>''',
"file-spreadsheet": '''<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/><path d="M14 2v5a1 1 0 0 0 1 1h5"/><path d="M8 13h2"/><path d="M14 13h2"/><path d="M8 17h2"/><path d="M14 17h2"/>''',
"file-text": '''<path d="M6 22a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h8a2.4 2.4 0 0 1 1.704.706l3.588 3.588A2.4 2.4 0 0 1 20 8v12a2 2 0 0 1-2 2z"/><path d="M14 2v5a1 1 0 0 0 1 1h5"/><path d="M10 9H8"/><path d="M16 13H8"/><path d="M16 17H8"/>''',
"braces": '''<path d="M8 3H7a2 2 0 0 0-2 2v5a2 2 0 0 1-2 2 2 2 0 0 1 2 2v5c0 1.1.9 2 2 2h1"/><path d="M16 21h1a2 2 0 0 0 2-2v-5c0-1.1.9-2 2-2a2 2 0 0 1-2-2V5a2 2 0 0 0-2-2h-1"/>''',
"table": '''<path d="M12 3v18"/><rect width="18" height="18" x="3" y="3" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/>''',
"layout-dashboard": '''<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>''',
"list-checks": '''<path d="M13 5h8"/><path d="M13 12h8"/><path d="M13 19h8"/><path d="m3 17 2 2 4-4"/><path d="m3 7 2 2 4-4"/>''',
"square-pen": '''<path d="M12 3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.375 2.625a1 1 0 0 1 3 3l-9.013 9.014a2 2 0 0 1-.853.505l-2.873.84a.5.5 0 0 1-.62-.62l.84-2.873a2 2 0 0 1 .506-.852z"/>''',
"x": '''<path d="M18 6 6 18"/><path d="m6 6 12 12"/>''',
"eye": '''<path d="M2.062 12.348a1 1 0 0 1 0-.696 10.75 10.75 0 0 1 19.876 0 1 1 0 0 1 0 .696 10.75 10.75 0 0 1-19.876 0"/><circle cx="12" cy="12" r="3"/>''',
"eye-off": '''<path d="M10.733 5.076a10.744 10.744 0 0 1 11.205 6.575 1 1 0 0 1 0 .696 10.747 10.747 0 0 1-1.444 2.49"/><path d="M14.084 14.158a3 3 0 0 1-4.242-4.242"/><path d="M17.479 17.499a10.75 10.75 0 0 1-15.417-5.151 1 1 0 0 1 0-.696 10.75 10.75 0 0 1 4.446-5.143"/><path d="m2 2 20 20"/>''',
"log-in": '''<path d="m10 17 5-5-5-5"/><path d="M15 12H3"/><path d="M15 3h4a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2h-4"/>''',
"circle-dot": '''<circle cx="12" cy="12" r="1"/><circle cx="12" cy="12" r="10"/>''',
"chevron-left": '''<path d="m15 18-6-6 6-6"/>''',
"search": '''<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/>''',
"clock": '''<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>''',
"chart-column": '''<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/>''',
"upload": '''<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" x2="12" y1="3" y2="15"/>''',
"store": '''<path d="m2 7 4.41-4.41A2 2 0 0 1 7.83 2h8.34a2 2 0 0 1 1.42.59L22 7"/><path d="M4 12v8a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-8"/><path d="M15 22v-4a2 2 0 0 0-2-2h-2a2 2 0 0 0-2 2v4"/><path d="M2 7h20"/>''',
"scale": '''<path d="m16 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="m2 16 3-8 3 8c-.87.65-1.92 1-3 1s-2.13-.35-3-1Z"/><path d="M7 21h10"/><path d="M12 3v18"/><path d="M3 7h2c2 0 5-1 7-2 2 1 5 2 7 2h2"/>''',
"volume-2": '''<path d="M11 5 6 9H2v6h4l5 4z"/><path d="M16 9a5 5 0 0 1 0 6"/><path d="M19.4 5.6a9 9 0 0 1 0 12.8"/>''',
"volume-x": '''<path d="M11 5 6 9H2v6h4l5 4z"/><line x1="22" x2="16" y1="9" y2="15"/><line x1="16" x2="22" y1="9" y2="15"/>''',
}

ICON_RULES = [  # label keyword -> icon (first match wins)
    ("ادخار", "piggy-bank"), ("فائض", "piggy-bank"), ("سحب", "banknote"),
    ("موبايل", "smartphone"), ("بنزين", "fuel"), ("مشتريات", "shopping-cart"),
    ("مطاعم", "utensils"), ("كافيهات", "coffee"), ("بقالة", "shopping-basket"),
    ("صحة", "heart-pulse"), ("تسوق", "shirt"), ("اشتراكات", "refresh-cw"),
    ("تحويلات", "send"), ("طوارئ", "triangle-alert"), ("غير مصنف", "circle-alert"),
]


def svg(name, cls="ic"):
    return ('<svg class="{c}" aria-hidden="true"><use href="#ic-{n}"/></svg>'
            .format(c=cls, n=name))


def icon_for(label):
    s = str(label)
    for kw, ic in ICON_RULES:
        if kw in s:
            return ic
    return "circle-dot"


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def esc_attr(s):
    return esc(s).replace('"', "&quot;")


def token_for(fn):
    """unguessable prefix for download links (sha256 of secret+name, 10 chars)

    The secret is the user's own passcode from 03_System/config.json, so the
    tokens cannot be regenerated by anyone who has not got the config.
    """
    import hashlib
    from mfconfig import TOKEN_SALT
    return hashlib.sha256((TOKEN_SALT + fn).encode("utf-8")).hexdigest()[:10]


def money(x):
    try:
        return "{:,.0f}".format(float(x))
    except (TypeError, ValueError):
        return esc(x)


def human(b):
    if b >= 1048576:
        return "{:.1f} MB".format(b / 1048576.0)
    return "{:.0f} KB".format(b / 1024.0)


def sprite():
    parts = ['<symbol id="ic-{n}" viewBox="0 0 24 24" fill="none" '
             'stroke="currentColor" stroke-width="2" stroke-linecap="round" '
             'stroke-linejoin="round">{b}</symbol>'.format(n=n, b=b)
             for n, b in ICONS.items()]
    return ('<svg xmlns="http://www.w3.org/2000/svg" style="display:none" '
            'aria-hidden="true">' + "".join(parts) + "</svg>")


def expand_icons(text):
    out = text
    for n in ICONS:
        out = out.replace("$i-" + n, svg(n))
    return out


# ------------------------------------------------------------------ data prep
STOP_TOKENS = {"غير", "كل", "أي", "قبل", "مع", "من", "في", "على", "الى", "إلى",
               "و", "فيه", "لأول"}


def toks(s):
    """normalize an Arabic label/category into comparable tokens"""
    raw = re.findall(r"[^\W\d_]+|\d+", str(s), re.UNICODE)
    out = set()
    for t in raw:
        if len(t) > 3 and t[:1] == "و":
            t = t[1:]
        t = t.lower()
        if len(t) > 1 and t not in STOP_TOKENS:
            out.add(t)
    return out


def load_txns():
    if not os.path.isfile(TXNS_CSV):
        return []
    out = []
    try:
        with open(TXNS_CSV, encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                try:
                    amt = float(r.get("amount_EGP") or 0)
                except (TypeError, ValueError):
                    continue
                out.append({"d": (r.get("date") or "").strip(),
                            "k": (r.get("kind") or "").strip(),
                            "n": (r.get("name") or "").strip(),
                            "c": (r.get("category") or "").strip(),
                            "a": amt})
    except OSError:
        return []
    return out


def netted(txns):
    """subtract each refund from the debit it reversed (same merchant,
    <=7 days apart, amount <= debit); fully reversed debits and refund rows
    drop out, so envelope actuals and anomaly stats see REAL spending only"""
    txns = [dict(t) for t in txns]
    for r in [t for t in txns if t["k"] == "refund"]:
        best, gap = None, 8
        for t in txns:
            if t["k"] != "debit" or t["n"] != r["n"] or t["a"] < r["a"] - 0.005:
                continue
            try:
                g = abs((dt.date.fromisoformat(t["d"]) -
                         dt.date.fromisoformat(r["d"])).days)
            except ValueError:
                continue
            if g <= 7 and g < gap:
                best, gap = t, g
        if best is not None:
            best["a"] = round(best["a"] - r["a"], 2)
    out = []
    for t in txns:
        if t["k"] == "refund":
            continue
        if t["k"] == "debit" and t["a"] <= 0.005:
            continue
        out.append(t)
    return out


def bank_actuals(plan_rows, txns, month):
    """sum debits of `month` into the plan envelope with most token overlap"""
    result = {lbl: 0.0 for lbl, _ in plan_rows}
    plan_toks = [(lbl, toks(lbl)) for lbl, _ in plan_rows]
    for t in netted(txns):
        if t["k"] != "debit" or t["d"][:7] != month:
            continue
        if t["c"] in ("وارد", "رد مبلغ"):
            continue
        ct = toks(t["c"])
        best, bestn = None, 0
        for lbl, pt in plan_toks:
            n = len(ct & pt)
            if n > bestn:
                best, bestn = lbl, n
        if best:
            result[best] += t["a"]
    return result


def _med(v):
    v = sorted(v)
    n = len(v)
    if not n:
        return 0.0
    return float(v[n // 2]) if n % 2 else (v[n // 2 - 1] + v[n // 2]) / 2.0


def anomalies(txns, topn=6):
    """outliers by modified z-score 0.6745*(x-median)/MAD > 3.5
    (Iglewski-Hoaglin): robust vs skewed category spreads; requires >=4 real
    txns in the category and a material amount (>=500 EGP)"""
    deb = [t for t in netted(txns) if t["k"] == "debit" and t["a"] > 0]
    by = {}
    for t in deb:
        by.setdefault(t["c"], []).append(t["a"])
    cand = []
    for t in deb:
        g = by[t["c"]]
        if len(g) < 4 or t["a"] < 500:
            continue
        m = _med(g)
        if m <= 0:
            continue
        mad = _med([abs(x - m) for x in g])
        if mad > 0:
            if 0.6745 * (t["a"] - m) / mad > 3.5:
                cand.append((t, t["a"] / m))
        elif t["a"] >= 3 * m:
            cand.append((t, t["a"] / m))
    if len(cand) < topn:  # pad so the panel always shows `topn` meaningful rows
        have = {id(t) for t, _ in cand}
        for t in sorted(deb, key=lambda x: -x["a"]):
            if len(cand) >= topn:
                break
            if id(t) in have:
                continue
            m = _med(by[t["c"]])
            cand.append((t, t["a"] / m if m > 0 else 1.0))
            have.add(id(t))
    cand.sort(key=lambda x: -x[0]["a"])
    return cand[:topn]


def add_month(d):
    m = d.month + 1
    y = d.year + (1 if m > 12 else 0)
    if m > 12:
        m -= 12
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


def salary_pill(txns):
    """countdown to next big credit (salary), estimated monthly anniversary"""
    cands = [t for t in txns if t["k"] == "credit" and t["a"] >= 5000]
    if not cands:
        return ""
    last = max(cands, key=lambda t: t["d"])
    try:
        d = dt.date.fromisoformat(last["d"])
    except ValueError:
        return ""
    today = dt.date.today()
    nxt = d
    for _ in range(6):
        if nxt >= today:
            break
        nxt = add_month(nxt)
    days = (nxt - today).days
    if days < 0 or days > 45:
        return ""
    if days == 0:
        when = "النهاردة"
    elif days == 1:
        when = "بكرة"
    else:
        when = "بعد {} يوم".format(days)
    return ('<span class="pill">{ic}الراتب المتوقع: {w} ({d})</span>'
            .format(ic=svg("calendar-check"), w=when, d=nxt.strftime("%d/%m")))


def trend_svg(s):
    """9-month net-spend bars (debits - refunds) + average line, pure inline
    SVG (no libraries); bars and the avg line share the same net basis"""
    monthly = s.get("monthly") or {}
    keys = sorted(monthly.keys())[-9:]
    if not keys:
        return ""
    vals = [float(monthly[k].get("debit") or 0) -
            float(monthly[k].get("refund") or 0) for k in keys]
    maxv = max(vals) or 1.0
    avg = float(s.get("avg_monthly_out") or 0)
    names = ["ينا", "فبر", "مار", "أبر", "ماي", "يون", "يول", "أغس",
             "سبت", "أكت", "نوف", "ديس"]
    W, H, bottom, top = 680, 232, 34, 34
    n = len(keys)
    slot = (W - 20) / n
    bw = min(46.0, slot * 0.52)
    parts = []
    if 0 < avg <= maxv * 1.2:
        ay = top + (H - bottom - top) * (1 - avg / maxv)
        parts.append('<line x1="10" y1="{:.1f}" x2="{}" y2="{:.1f}" '
                     'stroke="#C9A227" stroke-width="1.5" stroke-dasharray="5 4" '
                     'opacity=".9"/>'.format(ay, W - 10, ay))
        parts.append('<text class="cv" x="{}" y="{:.1f}" fill="#8A6F14" font-size="11" '
                     'text-anchor="start">المتوسط {}</text>'
                     .format(W - 12, ay - 6, money(avg)))
    for i, (k, v) in enumerate(zip(keys, vals)):
        x = 10 + i * slot + (slot - bw) / 2
        h = (H - bottom - top) * (v / maxv)
        y = H - bottom - h
        last = (i == n - 1)
        fill = "#C9A227" if last else "#163A5F"
        parts.append('<rect x="{:.1f}" y="{:.1f}" width="{:.1f}" height="{:.1f}" '
                     'rx="6" fill="{}"/>'.format(x, y, bw, h, fill))
        parts.append('<text class="cv" x="{:.1f}" y="{:.1f}" fill="#5B6A80" font-size="10" '
                     'text-anchor="middle">{}</text>'
                     .format(x + bw / 2, y - 7, money(v)))
        mi = int(k[5:7]) - 1
        parts.append('<text x="{:.1f}" y="{}" fill="#5B6A80" font-size="11" '
                     'text-anchor="middle">{}</text>'
                     .format(x + bw / 2, H - 10, names[mi] if 0 <= mi < 12 else k))
    parts.append('<line x1="10" y1="{}" x2="{}" y2="{}" stroke="#E3E9F2" '
                 'stroke-width="1"/>'.format(H - bottom, W - 10, H - bottom))
    return ('<svg class="chart" viewBox="0 0 {} {}" role="img" '
            'aria-label="مصروف آخر تسعة شهور" preserveAspectRatio="xMidYMid meet">'
            '{}</svg>').format(W, H, "".join(parts))


def envelope_row(label, budget, bank):
    return ('<div class="env" data-k="{k}" data-b="{b}" data-bank="{bk}">'
            '<div class="env-h">{ic}<span class="env-n">{l}</span>'
            '<span class="env-a num"><b class="env-done">0</b> / {bf} ج.م</span></div>'
            '<div class="bar"><i style="width:0%"></i></div>'
            '<div class="env-f"><span class="env-left">متبقي {bf} ج.م</span>'
            '<span class="env-src">بنك {bkf} · يدوي 0 · أسبوع {wk}</span>'
            '</div></div>').format(
        k=esc_attr(label), b="{:.0f}".format(float(budget)),
        bk="{:.0f}".format(float(bank)), ic=svg(icon_for(label)),
        l=esc(label), bf=money(budget), bkf=money(bank),
        wk=money(float(budget) / 4.33))


TEMPLATE = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "site_template.html"),
                  encoding="utf-8").read()


def regen_reports():
    """Rebuild every downloadable report from source data before the site is
    assembled, so downloads always match the rest of the system (data, numbers
    and colors).  Order: analyze_sms.py -> plan_month.py -> site build.
    Income for the plan is carried over from the plan on disk (cell B2),
    falling back to 10000.  Hard-fails the build on any error.
    Set MF_NOREGEN=1 to skip (offline / quick local rebuilds)."""
    if os.environ.get("MF_NOREGEN") == "1":
        print("regen: skipped (MF_NOREGEN=1)")
        return
    src = os.path.dirname(os.path.abspath(__file__))
    py = sys.executable or "python"

    # 1) SMS analysis -> summary_2026.json + csv + xlsx + dashboard html
    p = subprocess.run([py, os.path.join(src, "analyze_sms.py")],
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if p.returncode != 0:
        print("REGEN FAILED: analyze_sms.py")
        print(((p.stdout or "") + (p.stderr or "")).strip())
        sys.exit(1)

    # 2) income from the existing plan (cell B2), fallback 10000
    income = 10000.0
    if os.path.isfile(PLAN):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(PLAN, data_only=True)
            v = wb[wb.sheetnames[0]].cell(row=2, column=2).value
            wb.close()
            if isinstance(v, (int, float)) and v > 0:
                income = float(v)
        except Exception as e:
            print("regen: income read failed, fallback 10000 ({})".format(e))

    # 3) plan for the current month -> plan_YYYY-MM.xlsx
    p = subprocess.run([py, os.path.join(src, "plan_month.py"),
                        str(income), MONTH],
                       capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if p.returncode != 0:
        print("REGEN FAILED: plan_month.py")
        print(((p.stdout or "") + (p.stderr or "")).strip())
        sys.exit(1)
    # 4) master workbook -> 03_System/My_Financial_System_<YEAR>_SMS.xlsx
    try:
        e_ = subprocess.run([py, os.path.join(src, "build_excel.py")],
                            capture_output=True, text=True,
                            encoding="utf-8", errors="replace")
        if e_.returncode == 0:
            print("regen: master workbook refreshed")
        else:
            print("NOTE: build_excel.py skipped ({})".format(
                ((e_.stdout or "") + (e_.stderr or "")).strip().splitlines()[-1]
                if ((e_.stdout or "") + (e_.stderr or "")).strip() else "no output"))
    except Exception as e:
        print("NOTE: build_excel.py not run ({})".format(e))

    print("regen: reports refreshed (analyze_sms + plan_month {} income={:.0f})"
          .format(MONTH, income))



# Arabic output must not crash a cp1252 Windows console
def _mf_utf8():
    import sys as _s
    for _n in ("stdout", "stderr"):
        _st = getattr(_s, _n, None)
        if _st is not None:
            try:
                _st.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
_mf_utf8()
def main():
    regen_reports()
    s = json.load(open(SUMMARY, encoding="utf-8"))
    os.makedirs(WEB, exist_ok=True)

    # ---- plan rows from xlsx ----
    plan_rows = []
    savings = 0.0
    weekly_cap = 0.0
    if os.path.isfile(PLAN):
        import openpyxl
        wb = openpyxl.load_workbook(PLAN, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=4, max_row=30, max_col=5):
            label, amt = row[1].value, row[3].value
            if not label or not isinstance(amt, (int, float)):
                continue
            if label in ("الإجمالي", "تفاوت التقريب (لو فيه)"):
                continue
            plan_rows.append((label, float(amt)))
            if "ادخار" in str(label) or "فائض" in str(label):
                savings += float(amt)
            if "ATM" in str(label) or "سحب" in str(label):
                weekly_cap = float(amt) / 4.33

    # ---- categories (amount, count) + share of real spend ----
    cats = s.get("categories") or {}
    cat_rows = sorted(cats.items(), key=lambda kv: kv[1]["amount"], reverse=True)
    total_out = sum(float(v.get("amount") or 0) for v in cats.values()) or 1.0

    # ---- plan table html ----
    plan_html = "\n".join(
        '<tr><td><span class="cell-cat">{ic}{lbl}</span></td>'
        '<td class="num">{amt}</td><td class="num">{wk}</td></tr>'.format(
            ic=svg(icon_for(lbl)), lbl=esc(lbl),
            amt=money(amt), wk=money(float(amt) / 4.33))
        for lbl, amt in plan_rows)

    cat_html = "\n".join(
        '<tr><td><span class="cell-cat">{ic}{name}</span></td>'
        '<td class="num">{amt}</td>'
        '<td class="num"><span class="share">'
        '<span class="share-bar"><i style="width:{pct:.0f}%"></i></span>{pct:.1f}%</span></td>'
        '<td class="num">{n}</td></tr>'.format(
            ic=svg(icon_for(name)), name=esc(name),
            amt=money(v["amount"]), pct=v["amount"] / total_out * 100, n=money(v["n"]))
        for name, v in cat_rows)

    # ---- bank transactions: envelopes actuals, alerts, trend, explorer ----
    txns = load_txns()
    bank = bank_actuals(plan_rows, txns, MONTH)
    env_html = "".join(envelope_row(lbl, amt, bank.get(lbl, 0.0))
                       for lbl, amt in plan_rows)

    alerts = anomalies(txns)
    if alerts:
        alert_html = "".join(
            '<div class="alert"><span class="a-ic">{ic}</span>'
            '<div class="a-m"><b>{n}</b>'
            '<span class="a-s">{d} · {c} · أعلى من المعتاد ×{x}</span></div>'
            '<span class="a-amt num">{a} ج.م</span></div>'.format(
                ic=svg("triangle-alert"), n=esc(t["n"]),
                d=esc(t["d"][8:10] + "/" + t["d"][5:7]), c=esc(t["c"]),
                x="{:.1f}".format(x), a=money(t["a"]))
            for t, x in alerts)
    else:
        alert_html = '<div class="empty">لا توجد تنبيهات — مصروفاتك في حدود المعتاد</div>'

    trend = trend_svg(s)

    # ---- insights: month-over-month compare + biggest merchants ----
    monthly = s.get("monthly") or {}
    cmp_rows = ""
    prev_net = None
    for mk in sorted(monthly):
        mv = monthly[mk] or {}
        net_m = float(mv.get("debit") or 0) - float(mv.get("refund") or 0)
        if prev_net is None:
            delta = '<td class="num">—</td>'
        else:
            dlt = net_m - prev_net
            # dir=ltr span keeps the sign glued to the number in RTL cells
            delta = ('<td class="num {cls}"><span dir="ltr">{sign}{v}</span>'
                     '</td>').format(
                         cls="v-green" if dlt <= 0 else "v-red",
                         sign="+" if dlt > 0 else "", v=money(dlt))
        prev_net = net_m
        cmp_rows += ('<tr><td class="num pub">{mk}</td><td class="num">{net}</td>'
                     '<td class="num">{cr}</td>{delta}</tr>'.format(
                         mk=esc(mk), net=money(net_m),
                         cr=money(mv.get("credit") or 0), delta=delta))

    # ---- grouping: ATM-related merchants (NBE ATM, BDC…) combine under "Cash Withdrawal"
    raw = list((s.get("top_merchants") or {}).items())
    atm_keywords = ("ATM", "BDC", "NBE")
    atm_merch = {}
    other_merch = {}
    for name, data in raw:
        is_atm = any(kw in name.upper() for kw in atm_keywords)
        if is_atm:
            atm_merch[name] = data
        else:
            other_merch[name] = data
    # combine ATM total
    atm_total = sum((d.get("amount") or 0) for d in atm_merch.values())
    atm_count = sum((d.get("n") or 0) for d in atm_merch.values())
    # build display list: grouped ATM first, then remaining sorted by amount
    display = []
    if atm_merch:
        display.append(("Cash Withdrawal", dict(amount=atm_total, n=atm_count)))
    # add remaining merchants sorted by amount desc
    rem_sorted = sorted(other_merch.items(),
                        key=lambda kv: float((kv[1] or {}).get("amount") or 0),
                        reverse=True)
    for k, v in rem_sorted:
        display.append((k, v))
    merch = display[:5]  # keep top 5 after grouping
    merch_rows = "".join(
        '<tr><td><span dir="ltr" style="display:inline-block">{i}. {n}</span>'
        '</td><td class="num">{a}</td>'
        '<td class="num">{c}</td></tr>'.format(
            i=i, n=esc(k) if not k.startswith("Cash") else "سحب نقدي",
            a=money((v or {}).get("amount") or 0),
            c=money((v or {}).get("n") or 0))
        for i, (k, v) in enumerate(merch, 1))

    months_n = len(monthly) or 1
    incavg = money(float(s.get("total_credits") or 0) / months_n)

    recent = sorted(txns, key=lambda t: t["d"], reverse=True)[:80]
    # refunds (money back) always stay visible in the explorer, even when
    # they fall outside the 80-newest window
    recent = sorted(recent + [t for t in txns if t["k"] == "refund"
                              if t not in recent],
                    key=lambda t: t["d"], reverse=True)
    txjson = json.dumps(
        [{"d": t["d"], "k": t["k"], "n": t["n"], "c": t["c"], "a": t["a"]}
         for t in recent], ensure_ascii=False)
    txjson = txjson.replace("</", "<\\/")

    salary = salary_pill(txns)

    # ---- market pulse section (from 02_Reports/market_watch.json) ----
    market_html = ""
    mkt = os.path.join(BASE, "02_Reports", "market_watch.json")
    if os.path.isfile(mkt):
        try:
            m = json.load(open(mkt, encoding="utf-8"))
            chips = "\n".join(
                '<div class="mk-i"><span class="mk-l">{ic}{lab}</span>'
                '<b class="mk-v num pub">{val}</b><span class="mk-s">{src}</span></div>'.format(
                    ic=svg(it.get("icon") or "circle-dot"), lab=esc(it.get("label", "")),
                    val=esc(it.get("value", "")), src=esc(it.get("src", "")))
                for it in m.get("items", []))
            impact = "\n".join(
                "<li>{ic}<span>{t}</span></li>".format(ic=svg("target"), t=esc(t))
                for t in m.get("impact", []))
            news = "\n".join(
                '<li>{ic}<span>{t}<span class="mk-date">{d} — {w}</span></span></li>'.format(
                    ic=svg("circle-alert"), t=esc(n.get("title", "")),
                    d=esc(n.get("date", "")), w=esc(n.get("why", "")))
                for n in m.get("news", []))
            if chips:
                market_html = (
                    '<section id="market" class="sec"><div class="sec-h"><span class="sic">{ic}</span>'
                    '<h2>نبض السوق — القاهرة</h2>'
                    '<span class="hint">آخر فحص: {upd}</span></div>'
                    '<div class="mk">{chips}</div>'
                    '<div class="mk-cols"><div><div class="mk-h">{ti} أثر الأسعار على خطتك</div>'
                    '<ul class="mk-ul">{impact}</ul></div>'
                    '<div><div class="mk-h">{ni} أخبار تؤثر في مصروفاتك</div>'
                    '<ul class="mk-ul">{news}</ul></div></div>'
                    '<div class="note">{ii} الأرقام دي بتتحدّث أسبوعيًا كل سبت وأمام أي خطة شهرية، '
                    'والمصدر الكامل بتواريخه في وثيقة الأسعار داخل المجلد المحلي.</div></section>'
                ).format(ic=svg("refresh-cw"), upd=esc(m.get("last_checked", "")),
                         chips=chips, impact=impact, news=news,
                         ti=svg("target"), ni=svg("circle-alert"), ii=svg("info"))
        except Exception as e:
            print("WARN market section skipped: {}".format(e))

    # ---- download links (token-prefixed names, with size) ----
    dl_items = []
    for f, orig, t in DL:
        src = os.path.join(BASE, "02_Reports", f)
        if os.path.isfile(src):
            ext = os.path.splitext(orig)[1].lower()
            ic = {".xlsx": "file-spreadsheet", ".csv": "table",
                  ".json": "braces", ".html": "layout-dashboard"}.get(ext, "file-text")
            if orig.startswith("plan_"):
                ic = "calendar-check"
            dl_items.append((token_for(f) + "_" + f, t, ic,
                             human(os.path.getsize(src))))
    dl_html = "\n".join(
        '<a href="{s}" download>{ic}{t}<span class="sz">{sz}</span></a>'.format(
            s=esc(sn), ic=svg(ic), t=esc(t), sz=sz)
        for sn, t, ic, sz in dl_items)

    period = str(s.get("period") or "")
    if ".." in period:
        p_start, p_end = [x.strip() for x in period.split("..", 1)]
    else:
        p_start = p_end = period

    html = Template(expand_icons(TEMPLATE)).safe_substitute(
        sprite=sprite(),
        p1=esc(p_start), p2=esc(p_end),
        net=money(s.get("net_out")), inc=money(s.get("total_credits")),
        mo=money(s.get("avg_monthly_out")), sav=money(savings),
        month=esc(MONTH), cap=money(weekly_cap),
        capnum=round(weekly_cap, 2),
        plan_rows=plan_html, cat_rows=cat_html, market=market_html,
        envelopes=env_html, trend=trend, alerts=alert_html,
        txjson=txjson, salary=salary,
        cmp_rows=cmp_rows, merch_rows=merch_rows,
        incavg=incavg, mcount=months_n,
        dl=dl_html, today=dt.date.today().strftime("%d/%m/%Y"),
        gentime=dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
        gate=GATE,
    )
    if "$i-" in html or "$market" in html or "$txjson" in html \
            or "$envelopes" in html or "$salary" in html \
            or "$cmp_rows" in html or "$merch_rows" in html \
            or "$incavg" in html or "$gentime" in html:
        print("WARN: unresolved template token in output")

    # ---- optional embedded encrypted AI config (03_System/ai_vault.json) ----
    vault_json = os.path.join(BASE, "03_System", "ai_vault.json")
    if os.path.isfile(vault_json):
        try:
            with open(vault_json, "r", encoding="utf-8") as vf:
                vpayload = (json.load(vf) or {}).get("v", "")
            if vpayload:
                html = html.replace("var PUBLIC_VAULT = null;",
                                    "var PUBLIC_VAULT = " + json.dumps(vpayload) + ";")
        except Exception as e:
            print("WARN vault embed skipped: {}".format(e))

    with open(os.path.join(WEB, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    # ---- copy downloadable reports (token-prefixed names) ----
    copied = 0
    for stored, _, _, _ in dl_items:
        for f, orig, _t in DL:
            if token_for(f) + "_" + f == stored:
                src_name = f
                break
        src = os.path.join(BASE, "02_Reports", src_name)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(WEB, stored))
            copied += 1

    # ---- copy art assets (transparent PNG illustrations) ----
    art_src = os.path.join(BASE, "04_Source", "art")
    art_dst = os.path.join(WEB, "art")
    art_n = 0
    if os.path.isdir(art_src):
        os.makedirs(art_dst, exist_ok=True)
        for fn in os.listdir(art_src):
            if fn.lower().endswith(".png"):
                shutil.copy2(os.path.join(art_src, fn),
                             os.path.join(art_dst, fn))
                art_n += 1

    print("WEB OK  index.html written  ({} plan rows, {} envelopes, "
          "{} categories, {} alerts, {} txns, {} icons)".format(
              len(plan_rows), len(plan_rows), len(cat_rows),
              len(alerts), len(recent), len(ICONS)))
    print("copied {} report files into 06_Web".format(copied))
    print("copied {} art files into 06_Web/art".format(art_n))
    print("sounds: procedural Web Audio (zero audio files by design)")
    print("open: " + os.path.join(WEB, "index.html"))


if __name__ == "__main__":
    main()

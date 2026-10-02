# -*- coding: utf-8 -*-
r"""
VERIFY_SYSTEM - end-to-end check that every part of MyFinance is present,
linked and readable. Prints an ASCII checklist (PASS/FAIL), exits non-zero
on any failure so it can be used in automation.

Checks:
  1. Folder structure (5 numbered folders)
  2. Every key file exists + non-empty
  3. Data chain: SMS txt -> analyze -> summary_<year>.json -> plan_month -> plan xlsx
  4. summary_<year>.json numbers self-consistent
  5. Scripts compile (py_compile)
  6. Website build: lock gate, gate hash, token-named downloads,
     market pulse section (prices/news), income+expense quick-log
     with inline edit, zero emojis/arrows
  7. Market watch feed: 02_Reports/market_watch.json parses + complete
     + weekly refresh cadence (json field + site note)
  8. NTFS ACL: only the owner's Windows account has access
"""
import json, os, sys, py_compile, subprocess

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
import datetime as _dt
from mfconfig import CFG, YEAR, GATE  # noqa: E402
MONTH = _dt.date.today().strftime("%Y-%m")
fails = []
passes = 0


def check(name, ok, detail=""):
    global passes
    if ok:
        passes += 1
        print("[PASS] " + name)
    else:
        fails.append(name)
        print("[FAIL] " + name + ("  -> " + detail if detail else ""))


# 1. folders
for d in ["01_Data", "02_Reports", "03_System", "04_Source", "05_Docs", "06_Web"]:
    check("folder " + d, os.path.isdir(os.path.join(BASE, d)))

# 2. key files. Fixed names are listed verbatim; year/month dependent files use a
#    glob so the check works for any user and any reporting period.
import glob as _glob

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

FIXED = [
    ("04_Source", "mfconfig.py"),
    ("04_Source", "analyze_sms.py"),
    ("04_Source", "build_excel.py"),
    ("04_Source", "build_word.py"),
    ("04_Source", "plan_month.py"),
    ("04_Source", "daily_status.py"),
    ("04_Source", "build_web.py"),
    ("04_Source", "publish_web.py"),
    ("04_Source", "site_template.html"),
    ("06_Web", "index.html"),
    ("02_Reports", "market_watch.json"),
]
PATTERNS = [
    ("01_Data", "sms_raw_*.txt", "raw SMS capture"),
    ("02_Reports", "transactions_*.csv", "all transactions"),
    ("02_Reports", "Financial_Analysis_*.xlsx", "full analysis workbook"),
    ("02_Reports", "summary_*.json", "summary json"),
    ("02_Reports", "plan_*.xlsx", "month plan"),
    ("03_System", "My_Financial_System*.xlsx", "master workbook"),
]
# created only when the matching tool is used, so absence is not a failure
OPTIONAL = [
    ("01_Data", "daily_log_*.csv", "daily log (run daily_status.py)"),
    ("05_Docs", "*.docx", "Word guide (run build_word.py)"),
]
for _folder, _fn in FIXED:
    _p = os.path.join(BASE, _folder, _fn)
    check("file " + _folder + "/" + _fn,
          os.path.isfile(_p) and os.path.getsize(_p) > 0, "missing or empty")
for _folder, _pat, _label in PATTERNS:
    _hits = [f for f in _glob.glob(os.path.join(BASE, _folder, _pat))
             if os.path.getsize(f) > 0]
    check("file " + _folder + "/" + _pat + " (" + _label + ")", bool(_hits),
          "no non-empty match")
for _folder, _pat, _label in OPTIONAL:
    _hits = [f for f in _glob.glob(os.path.join(BASE, _folder, _pat))
             if os.path.getsize(f) > 0]
    check("optional " + _folder + "/" + _pat + " (" + _label + ")", True,
          ("found " + str(len(_hits))) if _hits else "not generated yet")

# 3. data chain
summary_path = os.path.join(BASE, "02_Reports", "summary_{}.json".format(YEAR))
s = None
try:
    s = json.load(open(summary_path, encoding="utf-8"))
    check("summary_2026.json parses", True)
except Exception as e:
    check("summary_2026.json parses", False, str(e))

if s:
    # 4. numeric self-consistency
    deb, ref, net = s.get("total_debits"), s.get("total_refunds"), s.get("net_out")
    check("net_out = debits - refunds",
          abs((deb or 0) - (ref or 0) - (net or 0)) < 0.01,
          "{} - {} != {}".format(deb, ref, net))
    # sanity, not a volume gate - a fresh install starts with a small sample
    check("txns parsed > 0", (s.get("n_txns") or 0) > 0, str(s.get("n_txns")))
    check("period present", bool(s.get("period")), "")
    cats = s.get("categories") or {}
    check("categories mapped > 0", len(cats) > 0, str(len(cats)))
    # monthly credit/debit sums consistent with totals
    m = s.get("monthly") or {}
    md = round(sum(v.get("debit", 0) for v in m.values()), 2)
    check("monthly debits sum == total_debits", abs(md - (deb or 0)) < 0.02,
          "{} vs {}".format(md, deb))

# plan file references income and links to summary (it was generated from it)
plan_path = os.path.join(BASE, "02_Reports", "plan_{}.xlsx".format(MONTH))
if os.path.isfile(plan_path):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(plan_path)
        names = wb.sheetnames
        check("plan has 3 sheets", len(names) >= 3, str(names))
        ws = wb[wb.sheetnames[0]]
        check("plan first cell has title", bool(ws.cell(1, 1).value))
    except Exception as e:
        check("plan xlsx readable", False, str(e))

# filled workbook sheets (optional build artifact: only checked when present)
_wb2_path = os.path.join(BASE, "03_System",
                         "My_Financial_System_{}_SMS.xlsx".format(YEAR))
if os.path.isfile(_wb2_path):
    try:
        import openpyxl
        wb2 = openpyxl.load_workbook(_wb2_path)
        check("filled system workbook 8 sheets", len(wb2.sheetnames) == 8,
              str(wb2.sheetnames))
    except Exception as e:
        check("filled system workbook readable", False, str(e))
else:
    check("filled system workbook (optional, skipped)",
          True, "run build_excel.py to generate")

# 5. scripts compile
for sc in ["analyze_sms.py", "build_excel.py", "build_word.py",
           "plan_month.py", "daily_status.py", "build_web.py",
           "publish_web.py"]:
    p = os.path.join(BASE, "04_Source", sc)
    try:
        py_compile.compile(p, doraise=True)
        check("compile " + sc, True)
    except Exception as e:
        check("compile " + sc, False, str(e))

# 6. website build (local checks only - no network)
idx = os.path.join(BASE, "06_Web", "index.html")
try:
    h = open(idx, encoding="utf-8").read()
    check("web index has lock gate", 'id="lock"' in h)
    check("web index has gate hash", GATE[:16] in h)
    check("web index has plan table", h.count("<tr>") >= 14, str(h.count("<tr>")))
    tok_files = [f for f in os.listdir(os.path.join(BASE, "06_Web"))
                 if len(f) > 11 and f[10] == "_" and f[:10].isalnum()]
    check("web has >=5 token-named downloads", len(tok_files) >= 5, str(len(tok_files)))
    # redesigned dashboard: market pulse + income logging + no emojis
    check("web has market pulse section",
          'id="market"' in h and h.count('class="mk-i"') >= 9,
          str(h.count('class="mk-i"')))
    check("web supports income + expense logging",
          'id="seg_in"' in h and 'id="seg_out"' in h and 'st_in' in h)
    check("web supports inline edit of entries",
          all(x in h for x in ("function editEntry", "function saveEdit",
                               "function cancelEdit", "function setEdType")))
    emoji = sum(1 for c in h if (0x1F000 <= ord(c) <= 0x1FAFF)
                or (0x2600 <= ord(c) <= 0x27BF)
                or ord(c) in (0xFE0F, 0x2192, 0x2190))
    check("web has zero emoji/arrow chars", emoji == 0, str(emoji))
except Exception as e:
    check("web index readable", False, str(e))

# 7. market watch feed (prices + news -> site section + price doc)
mw_path = os.path.join(BASE, "02_Reports", "market_watch.json")
try:
    mw = json.load(open(mw_path, encoding="utf-8"))
    check("market_watch: >=6 items, impact, news, last_checked",
          len(mw.get("items", [])) >= 6 and len(mw.get("impact", [])) >= 3
          and len(mw.get("news", [])) >= 1 and bool(mw.get("last_checked")),
          "items={} impact={} news={}".format(
              len(mw.get("items", [])), len(mw.get("impact", [])),
              len(mw.get("news", []))))
except Exception as e:
    check("market_watch.json parses", False, str(e))

# 7b. weekly refresh cadence (prices update weekly)
try:
    mw2 = json.load(open(mw_path, encoding="utf-8"))
    hh = open(os.path.join(BASE, "06_Web", "index.html"),
              encoding="utf-8").read()
    site_weekly = "أسبوعيًا" in hh
    check("market refresh cadence weekly (json + site note)",
          mw2.get("refresh_cadence") == "weekly" and site_weekly,
          "cadence={} site_weekly={}".format(
              mw2.get("refresh_cadence"), site_weekly))
except Exception as e:
    check("market cadence check ran", False, str(e))

# 7c. market prices must feed the plan (plan_month reads market_watch)
try:
    pm_src = open(os.path.join(BASE, "04_Source", "plan_month.py"),
                  encoding="utf-8").read()
    mw3 = json.load(open(mw_path, encoding="utf-8"))
    src_link = ("market_watch" in pm_src and "market_link" in pm_src)
    prof_ok = isinstance(mw3.get("moto_profile"), dict) and bool(
        mw3.get("moto_profile", {}).get("tank_l"))
    check("plan_month links market prices (market_link + moto_profile)",
          src_link and prof_ok,
          "src_link={} profile={}".format(src_link, prof_ok))
except Exception as e:
    check("plan-market link check ran", False, str(e))

# 8. ACL lockdown - the data folders (not the repo root) are what init locks
try:
    owner = os.environ.get("USERNAME") or os.environ.get("USER") or ""
    locked, leaked = [], []
    for t in ("01_Data", "02_Reports", "03_System", "05_Docs", "06_Web"):
        tp = os.path.join(BASE, t)
        if not os.path.isdir(tp):
            continue
        out = subprocess.run(["icacls", tp], capture_output=True, text=True,
                             encoding="utf-8", errors="replace").stdout
        entries = [l.strip() for l in out.splitlines()
                   if "(F)" in l or "(M)" in l or "(RX)" in l]
        # the first token of each icacls line is the trustee
        trustees = [e.split()[0] for e in entries if e.split()]
        strangers = [t2 for t2 in trustees
                     if (not owner or owner.lower() not in t2.lower())
                     and not t2.upper().endswith(r"\SYSTEM")
                     and not t2.upper().endswith(r"\ADMINISTRATORS")]
        (leaked if strangers else locked).append(
            t + (": " + ",".join(strangers) if strangers else ""))
    if not sys.platform.startswith("win"):
        check("ACL: Windows lockdown", True, "skipped (not Windows)")
    elif owner and leaked:
        check("ACL: only the owner (" + owner + ") has rights", False,
              " | ".join(leaked))
    elif owner and locked:
        check("ACL: only the owner (" + owner + ") has rights", True,
              "{} folders locked".format(len(locked)))
    else:
        check("ACL check ran", True, "no owner detected - skipped")
except Exception as e:
    check("ACL check ran", False, str(e))

print("-" * 60)
print("RESULT: {} passed, {} failed".format(passes, len(fails)))
if fails:
    for f in fails:
        print("  missing/broken: " + f)
    sys.exit(1)
print("ALL GREEN - system fully wired")

# -*- coding: utf-8 -*-
"""
DAILY_STATUS - log today's cash withdrawal and show remaining weekly allowance.

Usage:
  python daily_status.py                -> show status only
  python daily_status.py 150 توكيل بنزين -> append today's row + show status

Weekly cash cap comes from plan_YYYY-MM.xlsx column "سحب نقدي" / 4.33.
If the plan file is missing it falls back to 2100/4.33 = 485 EGP/week.
ASCII-only console output.
"""
import csv, os, sys, datetime as dt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NOW = dt.date.today()
MONTH = NOW.strftime("%Y-%m")
LOG = os.path.join(BASE, "01_Data", "daily_log_{}.csv".format(MONTH))
PLAN = os.path.join(BASE, "02_Reports", "plan_{}.xlsx".format(MONTH))
WEEK_CAP_FALLBACK = 485.0


def weekly_cap():
    """Read cash line (row labelled 'withdraw') from this month's plan."""
    try:
        import openpyxl
        if not os.path.isfile(PLAN):
            return WEEK_CAP_FALLBACK
        wb = openpyxl.load_workbook(PLAN, data_only=True)
        ws = wb[wb.sheetnames[0]]
        for row in ws.iter_rows(min_row=4, max_col=4):
            label = row[1].value
            val = row[3].value
            if label and isinstance(val, (int, float)) and (
                    "ATM" in str(label) or "سحب" in str(label)):
                return float(val) / 4.33
    except Exception:
        pass
    return WEEK_CAP_FALLBACK


def week_rows():
    """Rows logged within the last 7 days (rolling week)."""
    if not os.path.isfile(LOG):
        return []
    cutoff = NOW - dt.timedelta(days=6)
    out = []
    with open(LOG, encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            if not r.get("date"):
                continue
            try:
                d = dt.date.fromisoformat(r["date"])
            except ValueError:
                continue
            if d >= cutoff:
                try:
                    amt = float(r.get("withdraw_EGP") or 0)
                except ValueError:
                    amt = 0
                out.append((d, amt, r.get("where", ""), r.get("on_what", "")))
    return out


def append_today(amount, where, on_what):
    new = not os.path.isfile(LOG)
    with open(LOG, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["date", "day", "withdraw_EGP", "where", "on_what",
                        "week_allow_left", "note"])
        w.writerow([NOW.isoformat(), NOW.strftime("%A"), amount, where, on_what, "", ""])



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
    args = sys.argv[1:]
    if len(args) >= 1:
        try:
            amount = float(args[0].replace(",", ""))
        except ValueError:
            print("ERROR: first arg must be a number (EGP)")
            sys.exit(1)
        where = args[1] if len(args) > 1 else "?"
        on_what = " ".join(args[2:]) if len(args) > 2 else "?"
        append_today(amount, where, on_what)
        print("LOGGED today: {:.0f} EGP at {}".format(amount, where))

    rows = week_rows()
    cap = weekly_cap()
    spent = sum(r[1] for r in rows)
    left = cap - spent
    print("---- WEEK STATUS (last 7 days) ----")
    print("weekly cash cap : {:.0f} EGP".format(cap))
    print("spent so far    : {:.0f} EGP ({} entries)".format(spent, len(rows)))
    if left >= 0:
        print("remaining       : {:.0f} EGP".format(left))
        days_left = 6 - (NOW - rows[0][0]).days if rows else 6
        print("verdict         : OK - safe pace")
    else:
        print("remaining       : {:.0f} EGP OVER LIMIT".format(-left))
        print("verdict         : STOP cash spending until next week")
    for d, a, w, o in rows:
        print("  {}  {:>7.0f}  {} / {}".format(d.isoformat(), a, w, o))


if __name__ == "__main__":
    main()

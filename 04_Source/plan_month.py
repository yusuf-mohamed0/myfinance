# -*- coding: utf-8 -*-
"""
PLAN_MONTH - Monthly budget generator (envelope system, EGP)

Usage : python plan_month.py <income_EGP> [YYYY-MM]
Example: python plan_month.py 10000 2026-10
Reads : 02_Reports/summary_2026.json  (real spending baseline from SMS)
Writes: 02_Reports/plan_YYYY-MM.xlsx  (Arabic RTL, 3 sheets)
Console output is ASCII-only to avoid cp1252 crashes.
"""
import json, os, sys, re, datetime as dt

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUMMARY = os.path.join(BASE, "02_Reports", "summary_2026.json")

# ---------- envelope design: (order, plan_label, pct_of_income, actual_category) ----------
# pct must sum to 100. "leftover" row absorbs rounding.
# Re-planned 27/09/2026 (user approved): motorcycle real running cost
# 1,129-1,685 EGP/month (ex-oil) vs old 600 envelope -> raised to 22% and
# rebalanced discretionary/surplus lines. Sum MUST stay 100.
ENVELOPES = [
    (1,  "ادخار وطوارئ (وزّعه قبل أي صرف)", 15.0,  None),
    (2,  "سحب نقدي - مصروف يومي مسجّل",     21.0,  "سحب نقدي (ATM)"),
    (3,  "موبايل وإنترنت",                    7.0,  "موبايل وإنترنت"),
    (4,  "بنزين وصيانة الموتوسيكل",          22.0,  "بنزين ووقود"),
    (5,  "مشتريات ومحلات",                    7.0,  "مشتريات ومحلات"),
    (6,  "مطاعم وتوصيل",                      3.0,  "مطاعم وتوصيل"),
    (7,  "كافيهات وترفيه",                    1.0,  "كافيهات وترفيه"),
    (8,  "بقالة وسوبر ماركت",                 4.0,  "بقالة وسوبر ماركت"),
    (9,  "صحة وصيدليات",                      3.0,  "صحة وصيدليات"),
    (10, "تسوق أونلاين",                      1.5,  "تسوق أونلاين"),
    (11, "اشتراكات ومحتوى",                   2.0,  "اشتراكات ومحتوى"),
    (12, "تحويلات ومدفوعات (أصدقاء)",        3.0,  "مدفوعات وتحويلات"),
    (13, "طوارئ غير متوقع (كل شهر)",         3.5,  None),
    (14, "فائض حر -> ادخار إضافي",            7.0,  None),
]
assert abs(sum(e[2] for e in ENVELOPES) - 100.0) < 1e-6, "ENVELOPES pct must sum to 100"

# derived percentages (single source of truth = ENVELOPES above)
PCT_SAVE = next(e[2] for e in ENVELOPES if "ادخار" in e[1])
PCT_SURPLUS = next(e[2] for e in ENVELOPES if "فائض" in e[1])

RULES = [
    "ادخر اول حاجة: احوّل خانة الادخار لحساب تاني يوم ما المرتب ينزل - قبل ما تصرف جنيه.",
    "كل سحبة ATM لازم تتسجل في نفس اليوم: رقم، مبلغ، فين صرفته - من غير كده الخطة باظت.",
    "الخانات سقف مش هدف: لو خلصت خانة بدري متحرمش نفسك - استلف من الفائض الحر مش من الادخار.",
    "لو المرتب أقل من المتوقع: اقطع من سحب النقدي والمطاعم والتسوق الاول - متقربش من الادخار والبنزين.",
    "لو المرتب أعلى: النص زيادة ادخار + النص زيادة فائض حر.",
    "أي مبلغ وارد كبير مش مرتب (هدية/رد) -> 100% لصندوق الطوارئ.",
    "مراجعة كل جمعة: قارن المصاريف بالمخطط - 15 دقيقة بس.",
]

# ---- wealth projection params (annual net return, conservative for Egypt) ----
ANN_RET = 0.12


def money(x):
    return "{:,.0f}".format(round(x))


def build(income, month_str, actual_monthly, days_observed):
    wb = Workbook()

    # ================= SHEET 1: the plan =================
    ws = wb.active
    ws.title = "الخطة"
    ws.sheet_view.rightToLeft = True
    NAVY, GOLD, GREEN, LIGHT = "0E2A47", "C9A227", "15803D", "EEF2F9"
    thin = Side(style="thin", color="B7C4D6")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cur = '#,##0 "ج.م"'

    def cell(r, c, v, bold=False, fill=None, color="000000", size=11, align="center", fmt=None):
        k = ws.cell(row=r, column=c, value=v)
        k.font = Font(name="Segoe UI", bold=bold, size=size, color=color)
        k.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        k.border = border
        if fill:
            k.fill = PatternFill("solid", fgColor=fill)
        if fmt:
            k.number_format = fmt
        return k

    cell(1, 1, "الخطة الشهرية - شهر " + month_str, True, NAVY, "FFFFFF", 14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    cell(2, 1, "الدخل المتوقع", True, GOLD)
    cell(2, 2, income, True, None, fmt=cur)
    cell(2, 3, "الفترة المرصودة", True, GOLD)
    cell(2, 4, "{} يوم = {:.1f} شهر".format(days_observed, days_observed / 30.44), fill=LIGHT)
    ws.merge_cells(start_row=2, start_column=4, end_row=2, end_column=5)

    hdr = ["#", "الخانة", "نسبة", "المبلغ بالجنيه", "سقف الأسبوع"]
    for i, h in enumerate(hdr, 1):
        cell(3, i, h, True, NAVY, "FFFFFF")

    r = 4
    total = 0
    rows = []
    for order, label, pct, cat in ENVELOPES:
        amt = round(income * pct / 100.0, -1)  # to nearest 10 EGP
        total += amt
        rows.append((order, label, pct, amt, cat))
        cell(r, 1, order, fill=LIGHT)
        cell(r, 2, label, align="right")
        cell(r, 3, str(pct) + "%", fill=LIGHT)
        cell(r, 4, amt, True, fmt=cur,
             color=GREEN if label.startswith("ادخار") else "000000")
        cell(r, 5, round(amt / 4.33, -1), fmt=cur, fill=LIGHT)
        r += 1
    # remainder line
    diff = income - total
    cell(r, 1, len(ENVELOPES) + 1, fill=LIGHT)
    cell(r, 2, "تفاوت التقريب (لو فيه)", align="right")
    cell(r, 3, "-", fill=LIGHT)
    cell(r, 4, diff, fmt=cur)
    cell(r, 5, round(diff / 4.33, -1), fmt=cur, fill=LIGHT)
    r += 1
    cell(r, 1, "", fill=GOLD)
    cell(r, 2, "الإجمالي", True, GOLD)
    cell(r, 3, "100%", True, GOLD)
    cell(r, 4, income, True, GOLD, fmt=cur)
    cell(r, 5, round(income / 4.33, -1), True, GOLD, fmt=cur)
    r += 2

    cell(r, 1, "القواعد الذهبية", True, NAVY, "FFFFFF", 13)
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    r += 1
    for i, rule in enumerate(RULES, 1):
        cell(r, 1, i, True, LIGHT)
        cell(r, 2, rule, align="right")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        for c in range(3, 6):
            ws.cell(row=r, column=c).border = border
        r += 1

    for col, w in zip("ABCDE", [5, 42, 9, 17, 15]):
        ws.column_dimensions[col].width = w

    # ================= SHEET 2: plan vs reality =================
    ws2 = wb.create_sheet("مقارنة_بالواقع")
    ws2.sheet_view.rightToLeft = True

    def cell2(rr, cc, v, bold=False, fill=None, color="000000", size=11, align="center", fmt=None):
        k = ws2.cell(row=rr, column=cc, value=v)
        k.font = Font(name="Segoe UI", bold=bold, size=size, color=color)
        k.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        k.border = border
        if fill:
            k.fill = PatternFill("solid", fgColor=fill)
        if fmt:
            k.number_format = fmt
        return k

    cell2(1, 1, "المخطط vs الواقع (متوسط شهري من رسائلك الفعلية)", True, NAVY, "FFFFFF", 13)
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    for i, h in enumerate(["الخانة", "متوسط صرفك الفعلي", "المخطط", "الفرق", "التقييم"], 1):
        cell2(2, i, h, True, NAVY, "FFFFFF")

    rr = 3
    reality_total = 0
    plan_total = 0
    for order, label, pct, amt, cat in rows:
        act = actual_monthly.get(cat) if cat else None
        reality_total += (act or 0)
        plan_total += amt
        cell2(rr, 1, label, align="right")
        if act is None:
            cell2(rr, 2, "-", fill=LIGHT)
            cell2(rr, 4, "-", fill=LIGHT)
            cell2(rr, 5, "جديد", fill=LIGHT)
        else:
            d = amt - act
            cell2(rr, 2, round(act), fmt=cur, fill=LIGHT)
            cell2(rr, 3, amt, True, fmt=cur)
            cell2(rr, 4, d, fmt=cur, color=GREEN if d <= 0 else "B91C1C")
            if d <= -300:
                ev = "قصاص كبير - نصايح احتياطية"
            elif d <= 0:
                ev = "ناقص شوية - واقعي"
            elif d <= 300:
                ev = "زيادة بسيطة"
            else:
                ev = "زيادة كبيرة - راجعها"
            cell2(rr, 5, ev, fill=LIGHT, size=10)
        rr += 1

    # unmapped actual categories (e.g. أخرى غير مصنف)
    mapped = set(c for *_, c in ENVELOPES if c)
    for cat, v in actual_monthly.items():
        if cat not in mapped:
            cell2(rr, 1, cat + " (غير مخطط له)", align="right")
            cell2(rr, 2, round(v), fmt=cur, fill="FDECEA")
            cell2(rr, 3, 0, fmt=cur)
            cell2(rr, 4, -round(v), fmt=cur, color="B91C1C")
            cell2(rr, 5, "صنّفه في الخطة القادمة", fill="FDECEA", size=10)
            rr += 1

    cell2(rr, 1, "الإجمالي", True, GOLD)
    cell2(rr, 2, round(reality_total), True, GOLD, fmt=cur)
    cell2(rr, 3, plan_total, True, GOLD, fmt=cur)
    cell2(rr, 4, plan_total - round(reality_total), True, GOLD, fmt=cur)
    cell2(rr, 5, "", fill=GOLD)

    for col, w in zip("ABCDE", [34, 20, 16, 14, 26]):
        ws2.column_dimensions[col].width = w

    # ================= SHEET 3: wealth path =================
    ws3 = wb.create_sheet("مسار_الثروة")
    ws3.sheet_view.rightToLeft = True

    def cell3(rr, cc, v, bold=False, fill=None, color="000000", size=11, align="center", fmt=None):
        k = ws3.cell(row=rr, column=cc, value=v)
        k.font = Font(name="Segoe UI", bold=bold, size=size, color=color)
        k.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        k.border = border
        if fill:
            k.fill = PatternFill("solid", fgColor=fill)
        if fmt:
            k.number_format = fmt
        return k

    # saving pool = ادخار % + فائض % (derived from ENVELOPES so it never drifts)
    sav = round(income * PCT_SAVE / 100.0, -1) + round(income * PCT_SURPLUS / 100.0, -1)
    cell3(1, 1, "إيه لو ادّخرت شهرياً وركّبتها في استثمار {:.0f}% سنوي صافي?".format(ANN_RET * 100),
          True, NAVY, "FFFFFF", 13)
    ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    for i, h in enumerate(["سنة", "ادخرت بس (بدون فائدة)", "مع الفائدة المركّبة", "المليون باقي كام"], 1):
        cell3(2, i, h, True, NAVY, "FFFFFF")

    monthly = sav
    rr = 3
    bal = 0.0
    for y in range(1, 16):
        for m in range(12):
            bal = bal * (1 + ANN_RET / 12.0) + monthly
        cell3(rr, 1, y, fill=LIGHT)
        cell3(rr, 2, monthly * 12 * y, fmt=cur)
        cell3(rr, 3, bal, True, fmt=cur, color=GREEN if bal >= 1_000_000 else "000000")
        cell3(rr, 4, max(0, 1_000_000 - bal), fmt=cur, fill=LIGHT,
              color=GREEN if bal >= 1_000_000 else "B91C1C")
        rr += 1
    cell3(rr, 1, "ملاحظة", True, GOLD)
    cell3(rr, 2, "الأرقام تقديرية تعليمية - مش ضمان. الفائض الحر بيتحط في نفس الحساب فبان أعلى.",
          align="right")
    ws3.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)

    for col, w in zip("ABCD", [8, 26, 24, 24]):
        ws3.column_dimensions[col].width = w

    out = os.path.join(BASE, "02_Reports", "plan_{}.xlsx".format(month_str))
    wb.save(out)
    return out, rows, sav


def market_link(income, rows):
    """MARKET -> PLAN coupling: compare real moto running costs (from
    02_Reports/market_watch.json) against the fuel/motorcycle envelope so
    market prices actually drive the budget. Prints ASCII-only lines (EGP)."""
    mw_path = os.path.join(BASE, "02_Reports", "market_watch.json")
    if not os.path.isfile(mw_path):
        print("market link: SKIP (market_watch.json not found)")
        return
    try:
        mw = json.load(open(mw_path, encoding="utf-8"))
        prof = mw.get("moto_profile") or {}
        # single source of truth for the price: the gasoline chip itself
        price = None
        for it in mw.get("items", []):
            if str(it.get("label", "")).startswith("بنزين 92"):
                mnum = re.search(r"(\d+(?:\.\d+)?)", str(it.get("value", "")))
                if mnum:
                    price = float(mnum.group(1))
        if price is None:
            price = float(prof.get("fuel_price_l", 0) or 0)
        tank = float(prof.get("tank_l", 20))
        i_lo, i_hi = prof.get("fill_interval_days", [10, 14])
        s_cost = float(prof.get("service_cost", 350))
        s_lo, s_hi = prof.get("service_interval_months", [1, 2])
        oil = float(prof.get("oil_per_1000km", 620))
        fills_lo, fills_hi = 30.0 / float(i_hi), 30.0 / float(i_lo)
        fuel_lo, fuel_hi = fills_lo * tank * price, fills_hi * tank * price
        srv_lo, srv_hi = s_cost / float(s_hi), s_cost / float(s_lo)
        env = None
        for order, label, pct, amt, cat in rows:
            if cat == "بنزين ووقود" or "بنزين" in str(label):
                env = float(amt)
        print("MARKET->PLAN link (checked {}, cadence {}):".format(
            mw.get("last_checked", "?"), mw.get("refresh_cadence", "?")))
        print("  fuel {:.2f} EGP/L x {:.0f} L = {:.0f} EGP per fill".format(
            price, tank, tank * price))
        print("  fills/month every {}-{} d: {:.1f}-{:.1f}  ->  fuel {:.0f}-{:.0f} EGP".format(
            i_lo, i_hi, fills_lo, fills_hi, fuel_lo, fuel_hi))
        print("  periodic service {:.0f} EGP every {}-{} mo  ->  {:.0f}-{:.0f} EGP".format(
            s_cost, s_lo, s_hi, srv_lo, srv_hi))
        print("  oil {:.0f} EGP / 1000 km (PENDING: km per fill from user)".format(oil))
        run_lo, run_hi = fuel_lo + srv_lo, fuel_hi + srv_hi
        print("  bike running cost ex-oil: {:.0f}-{:.0f} EGP/month".format(run_lo, run_hi))
        if env is None:
            print("WARN: fuel/moto envelope not found in plan rows")
            return
        print("  plan envelope: {:.0f} EGP/month ({:.1f}% of income)".format(
            env, env * 100.0 / income if income else 0))
        gap_lo, gap_hi = run_lo - env, run_hi - env
        if gap_hi > 0:
            print("WARN: market gap {:.0f} - {:.0f} EGP/month vs envelope - rebalance ENVELOPES".format(
                max(0.0, gap_lo), gap_hi))
        else:
            print("OK: envelope covers market-based running cost")
    except Exception as e:
        print("market link skipped: {}".format(e))


def main():
    argv = sys.argv[1:]
    if not argv:
        print("usage: plan_month.py <income_EGP> [YYYY-MM]")
        sys.exit(1)
    income = float(argv[0].replace(",", ""))
    month_str = argv[1] if len(argv) > 1 else dt.date.today().strftime("%Y-%m")

    if not os.path.exists(SUMMARY):
        print("ERROR: summary_2026.json not found - run analyze_sms.py first")
        sys.exit(1)
    s = json.load(open(SUMMARY, encoding="utf-8"))
    days = float(s.get("days", 235))
    months = days / 30.44
    actual_monthly = {k: v["amount"] / months for k, v in s.get("categories", {}).items()}

    out, rows, sav = build(income, month_str, actual_monthly, int(days))
    print("PLAN OK  month={}  income={:.0f}  saving_pool={:.0f}".format(month_str, income, sav))
    for order, label, pct, amt, cat in rows:
        print("  {:>2} {:<45} {:>7.0f} EGP".format(order, label.encode("ascii", "ignore").decode() or "item%d" % order, amt))
    print("saved: plan_{}.xlsx".format(month_str))
    market_link(income, rows)


if __name__ == "__main__":
    main()

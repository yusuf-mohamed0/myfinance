# -*- coding: utf-8 -*-
"""
ANALYZER - NBE debit card 3390 SMS transactions
Period: derived dynamically from first/last SMS in the raw file
Input : 01_Data/sms_raw_<start>_to_<end>.txt  (picks the newest file)
Output: 02_Reports/transactions_<year>.csv
        02_Reports/Financial_Analysis_<year>.xlsx   (Arabic sheets)
        02_Reports/dashboard_<year>.html            (Plotly, offline, RTL)
        02_Reports/summary_<year>.json
        03_System/My_Financial_System_<year>_SMS.xlsx  (filled copy)
Console output is ASCII-only to avoid cp1252 crashes.
"""
import re, os, csv, json, sys, unicodedata
from datetime import date, datetime
from collections import defaultdict, OrderedDict

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# scripts live in MyFinance\04_Source -> project root is one level up
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
from mfconfig import YEAR  # noqa: E402

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
def _newest_raw():
    """pick the most recent sms_raw_*.txt from 01_Data (sorted: newest last)"""
    import glob as _g
    hits = sorted(_g.glob(os.path.join(BASE, "01_Data", "sms_raw_*.txt")))
    if not hits:
        raise SystemExit("no 01_Data/sms_raw_*.txt found - put your SMS there first")
    return hits[-1]


RAW = _newest_raw()
REP = os.path.join(BASE, "02_Reports")
SRC_XLSX = os.path.join(BASE, "03_System", "My_Financial_System.xlsx")
FILLED_XLSX = os.path.join(BASE, "03_System",
                          "My_Financial_System_{}_SMS.xlsx".format(YEAR))
os.makedirs(REP, exist_ok=True)

YEAR_ASSUMED = YEAR

# ---------------------------------------------------------------- regex
R_DEBIT = re.compile(r"^\s*تم\s+خصم\s+([\d.]+)\s*EGP", re.U)
R_CREDIT = re.compile(r"^\s*تم\s+إضافة\s+تحويل", re.U)
R_REFUND = re.compile(r"^\s*تم\s+رد\s+مبلغ\s+([\d.]+)\s*EGP", re.U)
R_FAILED = re.compile(r"^\s*ناسف\s+لعدم\s+إتمام\s+المعاملة\s+ب?مبلغ\s+([\d.]+)\s*EGP", re.U)
R_DATE3 = re.compile(r"يوم\s*(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})")
R_DATE2 = re.compile(r"يوم\s*(\d{1,2})[/-](\d{1,2})(?!\d)")
R_MERCHANT = re.compile(r"عند\s*(.*?)\s*يوم\s*\d")
R_WHERE = re.compile(r"من\s+(.+?)\s*$")
R_BAL = re.compile(r"المتاح\s*([\d.]+)")
R_AMT_CREDIT = re.compile(r"مبلغ\s*([\d.]+)")
R_FROM = re.compile(r"من\s+(.+?)\s*رقم\s*مرجعي")
R_FAILED_AMT = R_FAILED


def clean(s):
    s = unicodedata.normalize("NFKC", s or "")
    return re.sub(r"\s+", " ", s).strip()


def parse_date(line, line_is_credit):
    m = R_DATE3.search(line)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), m.group(3)
        y = int(y) if len(y) == 4 else 2000 + int(y)
        return date(y, mo, d), "dmy"
    m = R_DATE2.search(line)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if line_is_credit:          # credits use mm-dd
            return date(YEAR_ASSUMED, a, b), "md"
        return date(YEAR_ASSUMED, b, a), "dm"  # debits use dd/mm
    return None, None


def categorize(merchant):
    u = merchant.upper()
    if "ATM" in u or "BDC" in u:
        return "سحب نقدي (ATM)"
    if "TALABAT" in u or "UBER" in u:
        return "مطاعم وتوصيل"
    if "VODAFONE" in u or "ORANGE" in u:
        return "موبايل وإنترنت"
    if any(k in u for k in ("WATANIYA", "ALOTHAIM", "KHIER WA BARAKA",
                            "THE STAR SUPER", "BYT ALRYADH", "CHOCO")):
        return "بقالة وسوبر ماركت"
    if any(k in u for k in ("OLA ENERGY", "TOTAL TERSA", "CIRCLE K")):
        return "بنزين ووقود"
    if "EL EZABY" in u or "LAMIAS" in u:
        return "صحة وصيدليات"
    if "CHILL OUT" in u or "STRELLA" in u:
        return "كافيهات وترفيه"
    if "AMAZON" in u:
        return "تسوق أونلاين"
    if "EC H AL AHRAM" in u or "AL AHRAM" in u:
        return "اشتراكات ومحتوى"
    if any(k in u for k in ("CIB", "NBK", "KFH", "FIBE", "QNBS",
                            "PAYMOB", "OPAY", "FAWRY")):
        return "مدفوعات وتحويلات"
    if any(k in u for k in ("BM ", "KASHIER", "85 HARAM", "SENSE",
                            "BESO", "AL NASSER", "FLASH", "ANIMALIA",
                            "GEIDEAE", "TAHA HUSS", "WALKER", "ANAS ELDEMESH",
                            "ROMIO", "ALQRYTY", "RUMBLE", "AWBE")):
        return "مشتريات ومحلات"
    return "أخرى (غير مصنف)"


# workbook category list (must match build_excel.py dropdown)
WB_CAT = {
    "سحب نقدي (ATM)": "أخرى",
    "مطاعم وتوصيل": "مطاعم وكافيهات",
    "موبايل وإنترنت": "فواتير ومرافق",
    "بقالة وسوبر ماركت": "بقالة وسوبر ماركت",
    "بنزين ووقود": "مواصلات ووقود",
    "صحة وصيدليات": "صحة",
    "كافيهات وترفيه": "مطاعم وكافيهات",
    "تسوق أونلاين": "أخرى",
    "اشتراكات ومحتوى": "ترفيه واشتراكات",
    "مدفوعات وتحويلات": "أخرى",
    "مشتريات ومحلات": "أخرى",
    "أخرى (غير مصنف)": "أخرى",
}

# ---------------------------------------------------------------- parse
txns = []
failed = []
otp_lines = 0

with open(RAW, "r", encoding="utf-8") as fh:
    lines = fh.readlines()

for idx, raw in enumerate(lines):
    line = raw.rstrip("\n")
    if not line.strip():
        continue
    if "NBE OTP" in line:
        otp_lines += 1
        continue
    if R_FAILED.search(line):
        m = R_FAILED.search(line)
        failed.append({"amount": float(m.group(1)), "line": clean(line)})
        continue

    if R_DEBIT.search(line):
        amount = float(R_DEBIT.search(line).group(1))
        dt, _ = parse_date(line, False)
        mm = R_MERCHANT.search(line)
        merchant = clean(mm.group(1)) if mm else "?"
        bal = None
        bm = R_BAL.search(line)
        if bm:
            bal = float(bm.group(1))
        txns.append(dict(kind="debit", date=dt, amount=amount,
                         name=merchant, category=categorize(merchant),
                         balance_after=bal, idx=idx))
        continue

    if R_CREDIT.search(line):
        m = R_AMT_CREDIT.search(line)
        amount = float(m.group(1))
        dt, _ = parse_date(line, True)
        fm = R_FROM.search(line)
        sender = clean(fm.group(1)) if fm else "?"
        txns.append(dict(kind="credit", date=dt, amount=amount,
                         name=sender, category="وارد", balance_after=None,
                         idx=idx))
        continue

    if R_REFUND.search(line):
        amount = float(R_REFUND.search(line).group(1))
        dt, _ = parse_date(line, False)
        wm = R_WHERE.search(line)
        where = clean(wm.group(1)) if wm else "?"
        txns.append(dict(kind="refund", date=dt, amount=amount,
                         name=where, category="رد مبلغ", balance_after=None,
                         idx=idx))
        continue

# chronological order (paste is newest-first)
txns = [t for t in txns if t["date"]]
txns.sort(key=lambda t: (t["date"], -t["idx"]))

undated = [t for t in txns if t["date"] is None]

# ---------------------------------------------------------------- validate
first = txns[0]
running = None
mismatches = []
for t in txns:
    if running is None and t["kind"] == "debit" and t["balance_after"] is not None:
        running = t["balance_after"]                 # balance right after first debit
        opening = t["balance_after"] + t["amount"]   # implied opening balance
        t["calc_balance"] = t["balance_after"]
        continue
    if t["kind"] in ("debit", "credit"):
        running = running + (t["amount"] if t["kind"] == "credit" else -t["amount"])
    elif t["kind"] == "refund":
        running = running + t["amount"]
    t["calc_balance"] = running
    if t["kind"] == "debit" and t["balance_after"] is not None and running is not None:
        if abs(running - t["balance_after"]) > 0.01:
            mismatches.append((t["date"].isoformat(), t["name"],
                               round(running, 2), t["balance_after"],
                               round(running - t["balance_after"], 2)))
closing_calc = running

# dynamic period + last reported balance (no hardcoded dates/amounts)
first_date = txns[0]["date"]
last_date = txns[-1]["date"]
_with_bal = [t for t in txns if t.get("balance_after") is not None]
closing_reported = _with_bal[-1]["balance_after"] if _with_bal else None

# ---------------------------------------------------------------- aggregates
debits = [t for t in txns if t["kind"] == "debit"]
credits = [t for t in txns if t["kind"] == "credit"]
refunds = [t for t in txns if t["kind"] == "refund"]


def _net_refunds(debits, refunds):
    """Return debits with their reversals subtracted (same merchant, <=7 days
    apart, amount <= debit). Fully reversed debits drop out, so every
    category/merchant/cash figure below reflects REAL spending only."""
    net = {id(t): t["amount"] for t in debits}
    for r in refunds:
        best, gap = None, 8
        for t in debits:
            if id(t) not in net or t["name"] != r["name"]:
                continue
            if net[id(t)] < r["amount"] - 0.005:
                continue
            g = abs((r["date"] - t["date"]).days)
            if g <= 7 and g < gap:
                best, gap = t, g
        if best is not None:
            net[id(best)] = round(net[id(best)] - r["amount"], 2)
    out = []
    for t in debits:
        amt = net.get(id(t), 0.0)
        if amt <= 0.005:
            continue
        t2 = dict(t)
        t2["amount"] = amt
        out.append(t2)
    return out


real_debits = _net_refunds(debits, refunds)

total_debits = sum(t["amount"] for t in debits)
total_refunds = sum(t["amount"] for t in refunds)
total_credits = sum(t["amount"] for t in credits)
net_out = total_debits - total_refunds
cash_out = sum(t["amount"] for t in real_debits
               if t["category"] == "سحب نقدي (ATM)")
card_out = net_out - cash_out

monthly = OrderedDict()
for t in txns:
    key = t["date"].strftime("%Y-%m")
    m = monthly.setdefault(key, dict(debit=0.0, credit=0.0, refund=0.0, n=0))
    m["n"] += 1
    if t["kind"] == "debit":
        m["debit"] += t["amount"]
    elif t["kind"] == "credit":
        m["credit"] += t["amount"]
    else:
        m["refund"] += t["amount"]

by_cat = defaultdict(lambda: dict(amount=0.0, n=0))
for t in real_debits:
    by_cat[t["category"]]["amount"] += t["amount"]
    by_cat[t["category"]]["n"] += 1
by_cat = OrderedDict(sorted(by_cat.items(), key=lambda kv: -kv[1]["amount"]))

by_merchant = defaultdict(lambda: dict(amount=0.0, n=0))
for t in real_debits:
    by_merchant[t["name"]]["amount"] += t["amount"]
    by_merchant[t["name"]]["n"] += 1
by_merchant = OrderedDict(sorted(by_merchant.items(), key=lambda kv: -kv[1]["amount"]))

by_sender = defaultdict(lambda: dict(amount=0.0, n=0))
for t in credits:
    by_sender[t["name"]]["amount"] += t["amount"]
    by_sender[t["name"]]["n"] += 1
by_sender = OrderedDict(sorted(by_sender.items(), key=lambda kv: -kv[1]["amount"]))

top20 = sorted(real_debits, key=lambda t: -t["amount"])[:20]
period_days = (last_date - first_date).days + 1

drift = None
if closing_calc is not None and closing_reported is not None:
    drift = closing_calc - closing_reported   # last reported balance in SMS

summary = dict(
    period=f"{first_date.isoformat()} .. {last_date.isoformat()}",
    days=period_days,
    n_lines=len(lines),
    n_otp=otp_lines,
    n_failed=len(failed),
    n_txns=len(txns),
    n_debits=len(debits),
    n_credits=len(credits),
    n_refunds=len(refunds),
    total_debits=round(total_debits, 2),
    total_refunds=round(total_refunds, 2),
    net_out=round(net_out, 2),
    cash_out=round(cash_out, 2),
    card_out=round(card_out, 2),
    total_credits=round(total_credits, 2),
    net_flow=round(total_credits - net_out, 2),
    opening=round(opening, 2) if opening else None,
    closing_calc=round(closing_calc, 2),
    closing_reported=(round(closing_reported, 2)
                      if closing_reported is not None else None),
    drift=round(drift, 2),
    avg_daily_out=round(net_out / period_days, 2),
    avg_monthly_out=round(net_out / (period_days / 30.44), 2),
    monthly={k: {kk: round(vv, 2) for kk, vv in v.items()} for k, v in monthly.items()},
    categories={k: dict(amount=round(v["amount"], 2), n=v["n"]) for k, v in by_cat.items()},
    top_merchants={k: dict(amount=round(v["amount"], 2), n=v["n"])
                   for k, v in list(by_merchant.items())[:25]},
    senders={k: dict(amount=round(v["amount"], 2), n=v["n"]) for k, v in by_sender.items()},
    balance_mismatches=mismatches,
)

# ---------------------------------------------------------------- CSV
csv_path = os.path.join(REP, "transactions_{}.csv".format(YEAR))
with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
    w = csv.writer(fh)
    w.writerow(["date", "kind", "name", "category", "amount_EGP",
                "balance_after", "calc_balance"])
    for t in txns:
        w.writerow([t["date"].isoformat(),
                    {"debit": "debit", "credit": "credit", "refund": "refund"}[t["kind"]],
                    t["name"], t["category"], f"{t['amount']:.2f}",
                    "" if t["balance_after"] is None else f"{t['balance_after']:.2f}",
                    "" if t.get("calc_balance") is None else f"{t['calc_balance']:.2f}"])

# ---------------------------------------------------------------- XLSX report
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

NAVY, GOLD, GREEN, RED = "0E2A47", "C9A227", "15803D", "B91C1C"
hdr_fill = PatternFill("solid", fgColor=NAVY)
hdr_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=11)
title_font = Font(name="Segoe UI", bold=True, size=14, color=NAVY)
money = '#,##0.00 "ج.م"'
thin = Border(*[Side(style="thin", color="BFBFBF")] * 4)


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin


def autofit(ws, widths):
    for i, wd in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = wd


wb = Workbook()
ws = wb.active
ws.title = "المعاملات"
ws.sheet_view.rightToLeft = True
ws["A1"] = f"كشف معاملات البطاقة 3390 ({first_date.strftime('%d/%m/%Y')} - {last_date.strftime('%d/%m/%Y')})"
ws["A1"].font = title_font
head = ["التاريخ", "النوع", "الطرف / التاجر", "التصنيف", "المبلغ", "رصيد بعد", "رصيد محسوب"]
ws.append([])
ws.append(head)
style_header(ws, 3, len(head))
for t in txns:
    kind_ar = {"debit": "مصروف", "credit": "وارد", "refund": "رد مبلغ"}[t["kind"]]
    ws.append([t["date"], kind_ar, t["name"], t["category"], t["amount"],
               t["balance_after"], t.get("calc_balance")])
for r in range(4, 4 + len(txns)):
    ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
    for c in (5, 6, 7):
        ws.cell(row=r, column=c).number_format = money
        ws.cell(row=r, column=c).border = thin
    for c in range(1, 8):
        ws.cell(row=r, column=c).font = Font(name="Segoe UI", size=10)
autofit(ws, [13, 10, 40, 20, 15, 15, 15])
ws.freeze_panes = "A4"

# monthly
ws2 = wb.create_sheet("ملخص_شهري")
ws2.sheet_view.rightToLeft = True
ws2["A1"] = "الملخص الشهري"
ws2["A1"].font = title_font
ws2.append([])
h2 = ["الشهر", "عدد العمليات", "إجمالي الخروج", "رد مبلغ", "صافي الخروج", "الوارد", "الصافي"]
ws2.append(h2)
style_header(ws2, 3, len(h2))
for k, v in monthly.items():
    ws2.append([k, v["n"], v["debit"], v["refund"],
                v["debit"] - v["refund"], v["credit"],
                v["credit"] - (v["debit"] - v["refund"])])
r = 4 + len(monthly)
ws2.cell(row=r, column=1, value="الإجمالي").font = Font(bold=True)
ws2.cell(row=r, column=2, value=len(txns)).font = Font(bold=True)
for col, val in ((3, total_debits), (4, total_refunds), (5, net_out),
                 (6, total_credits), (7, total_credits - net_out)):
    c = ws2.cell(row=r, column=col, value=val)
    c.font = Font(bold=True, color=NAVY)
for row in ws2.iter_rows(min_row=4, max_row=r, min_col=3, max_col=7):
    for c in row:
        c.number_format = money
autofit(ws2, [12, 15, 18, 14, 16, 16, 16])

# categories
ws3 = wb.create_sheet("التصنيفات")
ws3.sheet_view.rightToLeft = True
ws3["A1"] = "توزيع المصروفات حسب التصنيف"
ws3["A1"].font = title_font
ws3.append([])
h3 = ["التصنيف", "عدد العمليات", "الإجمالي", "النسبة %"]
ws3.append(h3)
style_header(ws3, 3, len(h3))
# shares must add to 100%: divide by the category total itself (categories
# are already refund-netted), never by gross total_debits
cat_total = sum(v["amount"] for v in by_cat.values()) or 1.0
for k, v in by_cat.items():
    ws3.append([k, v["n"], v["amount"],
                round(100 * v["amount"] / cat_total, 1)])
for row in ws3.iter_rows(min_row=4, max_row=3 + len(by_cat)):
    row[2].number_format = money
autofit(ws3, [26, 15, 16, 10])

# top merchants
ws4 = wb.create_sheet("أكبر_المصاريف")
ws4.sheet_view.rightToLeft = True
ws4["A1"] = "أكبر 20 عملية صرف"
ws4["A1"].font = title_font
ws4.append([])
h4 = ["#", "التاريخ", "التاجر", "التصنيف", "المبلغ"]
ws4.append(h4)
style_header(ws4, 3, len(h4))
for i, t in enumerate(top20, 1):
    ws4.append([i, t["date"], t["name"], t["category"], t["amount"]])
for row in ws4.iter_rows(min_row=4, max_row=3 + len(top20)):
    row[1].number_format = "yyyy-mm-dd"
    row[4].number_format = money
autofit(ws4, [5, 13, 38, 20, 15])

# credits
ws5 = wb.create_sheet("الواردات")
ws5.sheet_view.rightToLeft = True
ws5["A1"] = "التحويلات الواردة (إيداعات)"
ws5["A1"].font = title_font
ws5.append([])
h5 = ["المرسل", "عدد التحويلات", "الإجمالي"]
ws5.append(h5)
style_header(ws5, 3, len(h5))
for k, v in by_sender.items():
    ws5.append([k, v["n"], v["amount"]])
r5 = 4 + len(by_sender)
ws5.cell(row=r5, column=1, value="الإجمالي").font = Font(bold=True)
c = ws5.cell(row=r5, column=3, value=total_credits)
c.font = Font(bold=True, color=GREEN)
for row in ws5.iter_rows(min_row=4, max_row=r5, min_col=3, max_col=3):
    for cc in row:
        cc.number_format = money
autofit(ws5, [45, 16, 16])

# KPI sheet
ws6 = wb.create_sheet("ملخص_عامة")
ws6.sheet_view.rightToLeft = True
ws6["A1"] = f"مؤشرات عامة - فترة {first_date.strftime('%d/%m/%Y')} حتى {last_date.strftime('%d/%m/%Y')}"
ws6["A1"].font = title_font
rows = [
    ("عدد الأيام", period_days, ""),
    ("عدد العمليات المسجلة", len(txns), ""),
    ("رسائل OTP تم تجاهلها", otp_lines, "أكواد سرية - غير مصاريف"),
    ("عمليات فاشلة (رصيد غير كافٍ)", len(failed), f"{sum(f['amount'] for f in failed):.2f} ج.م مرتين 115"),
    ("إجمالي الخروج (خصومات)", total_debits, ""),
    ("إجمالي الردود", total_refunds, ""),
    ("صافي الخروج من البطاقة", net_out, ""),
    ("منه سحب نقدي كاش (ATM)", cash_out, "مش مصنف كمصاريف فعلية"),
    ("مشتريات ومدفوعات بالبطاقة", card_out, ""),
    ("إجمالي الوارد", total_credits, ""),
    ("الصافي (وارد - خروج)", total_credits - net_out, ""),
    ("متوسط الخروج اليومي", net_out / period_days, ""),
    ("متوسط الخروج الشهري", net_out / (period_days / 30.44), ""),
    ("رصيد أول الفترة (محسوب)", opening, ""),
    ("رصيد آخر الفترة (محسوب)", closing_calc, ""),
    ("رصيد آخر الفترة (من البنك)", closing_reported if closing_reported is not None else "", ""),
    ("فرق التحقق (drift)", drift, "لو أكبر من 0 ففيه رسايل ناقصة"),
]
ws6.append([])
ws6.append(["المؤشر", "القيمة", "ملاحظة"])
style_header(ws6, 3, 3)
for name, val, note in rows:
    ws6.append([name, round(val, 2) if isinstance(val, float) else val, note])
for row in ws6.iter_rows(min_row=4, max_row=3 + len(rows), min_col=2, max_col=2):
    for c in row:
        if isinstance(c.value, (int, float)):
            c.number_format = money
autofit(ws6, [34, 20, 38])

# balance mismatches
if mismatches:
    ws7 = wb.create_sheet("فحص_الأرصدة")
    ws7.sheet_view.rightToLeft = True
    ws7["A1"] = "مواقع اختلف فيها الرصيد المحسوب عن الرصيد المصرح (رسائل ناقصة محتملة)"
    ws7["A1"].font = title_font
    ws7.append([])
    ws7.append(["التاريخ", "التاجر", "رصيد محسوب", "رصيد البنك", "الفرق"])
    style_header(ws7, 3, 5)
    for mm in mismatches:
        ws7.append([mm[0], mm[1], mm[2], mm[3], mm[4]])
    autofit(ws7, [13, 38, 16, 16, 12])

xlsx_path = os.path.join(REP, "Financial_Analysis_{}.xlsx".format(YEAR))
wb.save(xlsx_path)

# ---------------------------------------------------------------- filled workbook
try:
    from openpyxl import load_workbook
    wbk = load_workbook(SRC_XLSX)
    sh = wbk["سجل_المصروفات"]
    # header row is 3, data starts at 4 (see build_excel.py)
    r = 4
    for t in txns:
        if r > 503:
            break
        d = t["date"]
        month = d.strftime("%Y-%m")
        if t["kind"] == "debit":
            kind = "مصروف"
            cat = WB_CAT.get(t["category"], "أخرى")
            note = "سحب نقدي" if t["category"] == "سحب نقدي (ATM)" else "من SMS - 3390"
            method = "بطاقة خصم مباشر"
        elif t["kind"] == "credit":
            kind = "دخل"
            cat = "أخرى"
            note = f"تحويل وارد من {t['name']}"[:120]
            method = "تحويل لحظي"
        else:
            kind = "دخل"
            cat = "أخرى"
            note = f"رد مبلغ من {t['name']}"[:120]
            method = "رد مبلغ"
        sh.cell(row=r, column=1, value=d).number_format = "yyyy-mm-dd"
        sh.cell(row=r, column=2, value=month)
        sh.cell(row=r, column=3, value=cat)
        sh.cell(row=r, column=4, value=kind)
        desc = t["name"] if t["kind"] == "debit" else t["name"]
        sh.cell(row=r, column=5, value=desc[:100])
        sh.cell(row=r, column=6, value=t["amount"])
        sh.cell(row=r, column=7, value=method)
        sh.cell(row=r, column=8, value=note)
        r += 1
    # clear leftover sample rows if any
    for rr in range(r, 504):
        for cc in range(1, 9):
            sh.cell(row=rr, column=cc, value=None)
    # dashboard month = last month of data
    try:
        wbk["لوحة التحكم"]["C3"] = last_date.strftime("%Y-%m")
    except Exception:
        pass
    wbk.save(FILLED_XLSX)
    filled_ok = True
except Exception as e:
    filled_ok = False
    fill_err = repr(e)

# ---------------------------------------------------------------- JSON
json_path = os.path.join(REP, "summary_{}.json".format(YEAR))
with open(json_path, "w", encoding="utf-8") as fh:
    json.dump(summary, fh, ensure_ascii=False, indent=2)

# ---------------------------------------------------------------- HTML dashboard
html_path = None
try:
    import plotly.graph_objects as go
    from plotly.offline import plot
    import plotly.io as pio

    figs = []

    # 1) monthly
    fig = go.Figure()
    fig.add_bar(x=list(monthly.keys()),
                y=[round(v["debit"] - v["refund"], 2) for v in monthly.values()],
                name="صافي الخروج", marker_color="#B91C1C")
    fig.add_bar(x=list(monthly.keys()),
                y=[round(v["credit"], 2) for v in monthly.values()],
                name="الوارد", marker_color="#15803D")
    fig.update_layout(barmode="group", title="الخروج والوارد شهريًا (ج.م)",
                      font=dict(family="Segoe UI, Arial", size=13),
                      plot_bgcolor="#fff", height=420)
    figs.append(fig)

    # 2) categories pie
    fig = go.Figure(go.Pie(labels=list(by_cat.keys()),
                           values=[round(v["amount"], 2) for v in by_cat.values()],
                           hole=0.45))
    fig.update_layout(title="توزيع المصروفات حسب التصنيف",
                      font=dict(family="Segoe UI, Arial", size=13), height=460)
    figs.append(fig)

    # 3) top 12 merchants
    tm = list(by_merchant.items())[:12]
    fig = go.Figure(go.Bar(x=[round(v["amount"], 2) for _, v in tm],
                           y=[k for k, _ in tm], orientation="h",
                           marker_color="#0E2A47"))
    fig.update_layout(yaxis=dict(categoryorder="total ascending"),
                      title="أكبر 12 جهة صرف (ج.م)",
                      font=dict(family="Segoe UI, Arial", size=13), height=460)
    figs.append(fig)

    # 4) running balance
    xs, ys = [], []
    for t in txns:
        if t.get("calc_balance") is not None:
            xs.append(t["date"].isoformat())
            ys.append(round(t["calc_balance"], 2))
    fig = go.Figure(go.Scatter(x=xs, y=ys, mode="lines+markers",
                               line=dict(color="#C9A227", width=2),
                               fill="tozeroy", fillcolor="rgba(201,162,39,0.15)"))
    fig.update_layout(title="حركة الرصيد عبر الفترة (ج.م)",
                      font=dict(family="Segoe UI, Arial", size=13),
                      height=400, plot_bgcolor="#fff")
    figs.append(fig)

    parts = []
    for i, f in enumerate(figs):
        parts.append(pio.to_html(f, full_html=False,
                                 include_plotlyjs=("inline" if i == 0 else False)))
    html_path = os.path.join(REP, "dashboard_{}.html".format(YEAR))
    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(
            '<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="utf-8">'
            '<title>لوحة تحليل المصروفات ' + str(YEAR) + ''
            '<style>body{font-family:"Segoe UI",Arial;background:#F2F5FA;margin:0;padding:24px}'
            '.card{background:#fff;border:1px solid #E3E9F2;border-radius:12px;'
            'box-shadow:0 2px 10px rgba(14,42,71,.07);'
            'padding:16px;margin-bottom:22px}h1{color:#0E2A47}'
            '.kpi{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:22px}'
            '.kpi div{background:#fff;border:1px solid #E3E9F2;border-radius:12px;padding:14px 18px;'
            'box-shadow:0 2px 8px rgba(14,42,71,.07);min-width:180px}'
            '.kpi b{display:block;font-size:20px;color:#0E2A47}'
            '.red{color:#B91C1C}.green{color:#15803D}</style></head><body>'
            "<h1>لوحة تحليل المصروفات - من 28 يناير إلى 19 سبتمبر 2026</h1>"
            '<div class="kpi">'
            f'<div>صافي الخروج من البطاقة<b class="red">{net_out:,.2f} ج.م</b></div>'
            f'<div>مشتريات ومدفوعات<b>{card_out:,.2f} ج.م</b></div>'
            f'<div>سحب نقدي ATM<b>{cash_out:,.2f} ج.م</b></div>'
            f'<div>إجمالي الوارد<b class="green">{total_credits:,.2f} ج.م</b></div>'
            f'<div>الصافي<b>{total_credits - net_out:,.2f} ج.م</b></div>'
            f'<div>متوسط يومي<b>{net_out / period_days:,.2f} ج.م</b></div>'
            f'<div>عدد العمليات<b>{len(txns)}</b></div>'
            f'<div>فرق التحقق<b>{drift:,.2f} ج.م</b></div>'
            "</div>"
            + "".join(f'<div class="card">{p}</div>' for p in parts)
            + "</body></html>"
        )
except Exception as e:
    html_err = repr(e)

# ---------------------------------------------------------------- console (ASCII only)
print(f"=== SMS SPENDING ANALYSIS {first_date.isoformat()} .. {last_date.isoformat()} ===")
print(f"lines={len(lines)} otp_ignored={otp_lines} failed={len(failed)} parsed_txns={len(txns)}")
print(f"gross_debits      = {total_debits:,.2f} EGP")
print(f"refunds           = {total_refunds:,.2f} EGP")
print(f"net_out           = {net_out:,.2f} EGP")
print(f"  cash ATM        = {cash_out:,.2f} EGP")
print(f"  card payments   = {card_out:,.2f} EGP")
print(f"credits_in        = {total_credits:,.2f} EGP")
print(f"net_flow          = {total_credits - net_out:,.2f} EGP")
print(f"days={period_days} avg_daily={net_out / period_days:,.2f} "
      f"avg_monthly={net_out / (period_days / 30.44):,.2f}")
print(f"opening={opening:,.2f} closing_calc={closing_calc:,.2f} "
      f"closing_reported={closing_reported if closing_reported is not None else float('nan'):,.2f} "
      f"drift={(drift if drift is not None else float('nan')):,.2f}")
print(f"balance_mismatches={len(mismatches)}")
for m in mismatches[:20]:
    print("   MISMATCH", m)
print("--- monthly (net_out / credits) ---")
for k, v in monthly.items():
    print(f"  {k}: out={v['debit'] - v['refund']:>10,.2f}  in={v['credit']:>10,.2f}  n={v['n']}")
print("--- categories ---")
for k, v in by_cat.items():
    print(f"  {k:<24} {v['amount']:>10,.2f}  n={v['n']}")
print("--- top 12 merchants ---")
for k, v in list(by_merchant.items())[:12]:
    print(f"  {k[:44]:<44} {v['amount']:>10,.2f}  n={v['n']}")
print("--- senders ---")
for k, v in by_sender.items():
    print(f"  {k[:44]:<44} {v['amount']:>10,.2f}  n={v['n']}")
print(f"filled_workbook={FILLED_XLSX if filled_ok else 'FAILED: ' + fill_err}")
print(f"xlsx_report={xlsx_path}")
print(f"csv={csv_path}")
print(f"json={json_path}")
print(f"dashboard={html_path if html_path else 'FAILED: ' + html_err}")

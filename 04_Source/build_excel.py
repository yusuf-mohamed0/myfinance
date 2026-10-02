# -*- coding: utf-8 -*-
"""Personal Finance System - Excel Workbook Builder"""
import datetime as dt
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.worksheet.datavalidation import DataValidation

_BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(_BASE, "03_System", "My_Financial_System.xlsx")

# ---------- Style palette ----------
NAVY = "0E2A47"
GOLD = "C9A227"
GREEN = "15803D"
RED = "B91C1C"
LIGHT = "EEF2F9"
BAND = "F5F8FC"
WHITE = "FFFFFF"

H1 = Font(name="Segoe UI", size=16, bold=True, color=WHITE)
H2 = Font(name="Segoe UI", size=12, bold=True, color=NAVY)
BOLD = Font(name="Segoe UI", size=11, bold=True)
BODY = Font(name="Segoe UI", size=11)
SMALL = Font(name="Segoe UI", size=10, italic=True, color="555555")
GOLD_F = Font(name="Segoe UI", size=11, bold=True, color="7A5C00")

fill_nav = PatternFill("solid", fgColor=NAVY)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_light = PatternFill("solid", fgColor=LIGHT)
fill_band = PatternFill("solid", fgColor=BAND)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_green = PatternFill("solid", fgColor="E3F2E6")
fill_red = PatternFill("solid", fgColor="FDEAEA")
fill_yellow = PatternFill("solid", fgColor="FFF6DD")

thin = Side(style="thin", color="B8C4D9")
med = Side(style="medium", color=NAVY)
box = Border(left=thin, right=thin, top=thin, bottom=thin)

money_fmt = '#,##0.00 "ج.م"'
pct_fmt = "0.0%"
date_fmt = "yyyy-mm-dd"

def title_bar(ws, text, ncols, row=1):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.font = H1
    c.fill = fill_nav
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 30

def header_row(ws, row, headers, fill=None):
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=row, column=i, value=h)
        c.font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
        c.fill = fill if fill else PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = box
    ws.row_dimensions[row].height = 26

def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

wb = Workbook()

# ==========================================================
# SHEET 1: Dashboard (لوحة التحكم)
# ==========================================================
ws = wb.active
ws.title = "لوحة التحكم"
set_widths(ws, [3, 34, 22, 22, 22, 22, 3])
title_bar(ws, "لوحة التحكم المالية | Your Financial Command Center", 7, 1)

ws["B3"] = "الشهر الحالي (غيّره كل شهر بالنص):"
ws["B3"].font = BOLD
ws["C3"] = "2026-09"
ws["C3"].font = GOLD_F
ws["C3"].fill = fill_yellow
ws["C3"].border = box
ws["C3"].alignment = Alignment(horizontal="center")
ws["D3"] = "مثال: 2026-10 للشهر الجاي"
ws["D3"].font = SMALL

# KPI cards
kpis = [
    ("إجمالي الدخل", '=SUMIFS(سجل_المصروفات!F:F,سجل_المصروفات!D:D,"دخل",سجل_المصروفات!B:B,$C$3)'),
    ("إجمالي المصروفات", '=SUMIFS(سجل_المصروفات!F:F,سجل_المصروفات!D:D,"مصروف",سجل_المصروفات!B:B,$C$3)'),
    ("صافي التدفق (وفر)", '=B6-B7'),
    ("نسبة الادخار %", '=IF(B6=0,0,B8/B6)'),
]
r = 5
for i, (label, formula) in enumerate(kpis):
    col = 2 + i
    c = ws.cell(row=r, column=col, value=label)
    c.font = Font(name="Segoe UI", size=11, bold=True, color=WHITE)
    c.fill = fill_nav
    c.alignment = Alignment(horizontal="center")
    c.border = box
    v = ws.cell(row=r + 1, column=col, value=formula)
    v.font = Font(name="Segoe UI", size=14, bold=True, color=NAVY)
    v.fill = fill_light
    v.alignment = Alignment(horizontal="center", vertical="center")
    v.border = box
    v.number_format = pct_fmt if i == 3 else money_fmt
    ws.row_dimensions[r + 1].height = 34
# fix references (kpis use B6/B7/B8 => col B is index 2)
ws["B8"] = "=B6-C6"      # net = income - expenses
ws["B9"] = "=IF(B6=0,0,B8/B6)"
ws["B8"].number_format = money_fmt
ws["B9"].number_format = pct_fmt
ws["B8"].font = Font(name="Segoe UI", size=14, bold=True, color=GREEN)
ws["B9"].font = Font(name="Segoe UI", size=14, bold=True, color=GREEN)

# Budget vs actual table
ws["B12"] = "مقارنة الميزانية بالفعل (من شيت الميزانية)"
ws["B12"].font = H2
ws.merge_cells("B12:E12")

header_row(ws, 13, ["", "البند", "المخطط", "الفعلي", "الفرق"], None)
ws.cell(row=13, column=1).fill = fill_white

cats = ["سكن وإيجار", "فواتير ومرافق", "مواصلات ووقود", "بقالة وسوبر ماركت",
        "مطاعم وكافيهات", "صحة", "تعليم", "ملابس وعناية",
        "ترفيه واشتراكات", "هدايا وتبرعات", "أخرى"]
for i, cat in enumerate(cats):
    rr = 14 + i
    ws.cell(row=rr, column=2, value=cat).font = BODY
    ws.cell(row=rr, column=3, value=f"=IFERROR(VLOOKUP(B{rr},الميزانية!B:D,3,FALSE),0)").number_format = money_fmt
    ws.cell(row=rr, column=4, value=f'=SUMIFS(سجل_المصروفات!F:F,سجل_المصروفات!D:D,"مصروف",سجل_المصروفات!C:C,B{rr},سجل_المصروفات!B:B,$C$3)').number_format = money_fmt
    ws.cell(row=rr, column=5, value=f"=C{rr}-D{rr}").number_format = money_fmt
    for cc in range(2, 6):
        ws.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws.cell(row=rr, column=cc).fill = fill_band

tr = 14 + len(cats)
ws.cell(row=tr, column=2, value="الإجمالي").font = BOLD
ws.cell(row=tr, column=3, value=f"=SUM(C14:C{tr-1})").number_format = money_fmt
ws.cell(row=tr, column=4, value=f"=SUM(D14:D{tr-1})").number_format = money_fmt
ws.cell(row=tr, column=5, value=f"=SUM(E14:E{tr-1})").number_format = money_fmt
for cc in range(2, 6):
    c = ws.cell(row=tr, column=cc)
    c.font = BOLD
    c.fill = fill_gold
    c.border = box

ws.conditional_formatting.add(f"E14:E{tr-1}",
    CellIsRule(operator="lessThan", formula=["0"], fill=fill_red))
ws.conditional_formatting.add(f"E14:E{tr-1}",
    CellIsRule(operator="greaterThanOrEqual", formula=["0"], fill=fill_green))

# Emergency fund + goals status
gr = tr + 2
ws.cell(row=gr, column=2, value="حالة الأهداف السريعة").font = H2
goals_status = [
    ("صندوق الطوارئ (مستهدف: 6 مصاريف شهرية)", '=IFERROR(VLOOKUP("صندوق الطوارئ",الأهداف!B:G,4,FALSE),0)', money_fmt),
    ("إجمالي الادخار الحالي", '=SUM(الأهداف!E:E)', money_fmt),
    ("إجمالي الديون المتبقية", '=IFERROR(الديون!C30,0)', money_fmt),
]
for i, (lbl, f, fmt) in enumerate(goals_status):
    rr = gr + 1 + i
    ws.cell(row=rr, column=2, value=lbl).font = BODY
    c = ws.cell(row=rr, column=3, value=f)
    c.number_format = fmt
    c.font = BOLD
    c.border = box
    c.fill = fill_light
    ws.cell(row=rr, column=2).border = box

ws.sheet_view.showGridLines = False
ws.freeze_panes = "A4"

# ==========================================================
# SHEET 2: سجل المصروفات (Transactions Log)
# ==========================================================
ws2 = wb.create_sheet("سجل_المصروفات")
set_widths(ws2, [14, 12, 26, 12, 34, 16, 16, 30])
title_bar(ws2, "سجل الدخل والمصروفات اليومي - اكتب كل جنيه هنا", 8, 1)
ws2["A2"] = "طريقة الاستخدام: التاريخ | الشهر (2026-09) | الفئة | دخل/مصروف | الوصف | المبلغ | طريقة الدفع | ملاحظات"
ws2["A2"].font = SMALL
ws2.merge_cells("A2:H2")

header_row(ws2, 3, ["التاريخ", "الشهر", "الفئة", "النوع", "الوصف", "المبلغ", "طريقة الدفع", "ملاحظات"])

sample = [
    (dt.date(2026, 9, 1), "2026-09", "مرتب", "دخل", "مرتب سبتمبر", 15000, "تحويل بنكي", ""),
    (dt.date(2026, 9, 1), "2026-09", "سكن وإيجار", "مصروف", "إيجار الشقة", 5000, "كاش", ""),
    (dt.date(2026, 9, 2), "2026-09", "بقالة وسوبر ماركت", "مصروف", "سوبر ماركت الشهر", 1800, "فيزا", ""),
    (dt.date(2026, 9, 3), "2026-09", "مواصلات ووقود", "مصروف", "بنزين + مواصلات", 700, "كاش", ""),
    (dt.date(2026, 9, 5), "2026-09", "فواتير ومرافق", "مصروف", "كهرباء + جاز + إنترنت", 900, "أونلاين", ""),
]
for i, row_data in enumerate(sample):
    rr = 4 + i
    for j, v in enumerate(row_data, start=1):
        c = ws2.cell(row=rr, column=j, value=v)
        c.border = box
        c.font = BODY
        if i % 2 == 1:
            c.fill = fill_band
        if j == 1:
            c.number_format = date_fmt
            c.alignment = Alignment(horizontal="center")
        if j == 6:
            c.number_format = money_fmt
ws2.cell(row=4, column=6).font = Font(name="Segoe UI", size=11, bold=True, color=GREEN)

# blank rows for data entry with borders + validation
FIRST, LAST = 4, 503
for rr in range(4 + len(sample), LAST + 1):
    for j in range(1, 9):
        c = ws2.cell(row=rr, column=j)
        c.border = box
        c.font = BODY
        if rr % 2 == 1:
            c.fill = fill_band
    ws2.cell(row=rr, column=1).number_format = date_fmt
    ws2.cell(row=rr, column=6).number_format = money_fmt

dv_type = DataValidation(type="list", formula1='"دخل,مصروف"', allow_blank=True)
dv_type.error = "اختر دخل أو مصروف"
ws2.add_data_validation(dv_type)
dv_type.add(f"D{FIRST}:D{LAST}")

dv_cat = DataValidation(type="list", formula1='"مرتب,عمل إضافي,استثمار,مرتب,سكن وإيجار,فواتير ومرافق,مواصلات ووقود,بقالة وسوبر ماركت,مطاعم وكافيهات,صحة,تعليم,ملابس وعناية,ترفيه واشتراكات,هدايا وتبرعات,أخرى"', allow_blank=True)
ws2.add_data_validation(dv_cat)
dv_cat.add(f"C{FIRST}:C{LAST}")

dv_pay = DataValidation(type="list", formula1='"كاش,فيزا,محفظة إلكترونية,تحويل بنكي,أونلاين,أخرى"', allow_blank=True)
ws2.add_data_validation(dv_pay)
dv_pay.add(f"G{FIRST}:G{LAST}")

ws2.freeze_panes = "A4"
ws2.auto_filter.ref = f"A3:H{LAST}"

# ==========================================================
# SHEET 3: الميزانية (Monthly Budget)
# ==========================================================
ws3 = wb.create_sheet("الميزانية")
set_widths(ws3, [3, 30, 18, 18, 18, 40])
title_bar(ws3, "ميزانية الشهر بنظام 50/30/20 المعدل", 6, 1)

ws3["B3"] = "الراتب الصافي الشهري:"
ws3["B3"].font = BOLD
ws3["C3"] = 0
ws3["C3"].number_format = money_fmt
ws3["C3"].fill = fill_yellow
ws3["C3"].border = box
ws3["C3"].font = GOLD_F
ws3["D3"] = "← ضع الراتب هنا أولاً"
ws3["D3"].font = SMALL

header_row(ws3, 5, ["", "الفئة", "الحد الشهري", "المصروف فعلياً", "المتبقي", "ملاحظات الاستخدام"])

budget_rows = [
    ("الأساسيات 50%", [
        ("سكن وإيجار", 0.30, "الأعلى ثباتاً - لا يزيد عن 30% من الدخل"),
        ("فواتير ومرافق", 0.07, "كهربا/غاز/ماء/إنترنت/موبايل"),
        ("مواصلات ووقود", 0.06, "بنزين + مواصلات + صيانة عربية"),
        ("بقالة وسوبر ماركت", 0.07, "اشتري بقائمة.. بلا تسوق انفعالي"),
    ]),
    ("احتياجات 25%", [
        ("صحة", 0.05, "تأمين + أدوية + كشف"),
        ("تعليم", 0.05, "كتب + كورسات + تطوير ذات"),
        ("ملابس وعناية", 0.05, "ملابس أساسية وعناية شخصية"),
        ("أخرى", 0.05, "طوارئ مصروفات غير متوقعة"),
    ]),
    ("رغبات 15%", [
        ("مطاعم وكافيهات", 0.08, "خليه محدود.. أكبر قاتل للميزانية"),
        ("ترفيه واشتراكات", 0.04, "نت/نتفليكس/أخرجات"),
        ("هدايا وتبرعات", 0.03, "صدقة + هدايا (أجر معاك)"),
    ]),
    ("ادخار واستثمار 20%", []),
]

r = 6
first_data_row = None
for group, items in budget_rows:
    if not items:
        continue
    ws3.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    c = ws3.cell(row=r, column=2, value=group)
    c.font = Font(name="Segoe UI", size=12, bold=True, color=WHITE)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = Alignment(horizontal="right", vertical="center")
    r += 1
    if first_data_row is None:
        first_data_row = r
    for name, pct, note in items:
        ws3.cell(row=r, column=2, value=name).font = BODY
        ws3.cell(row=r, column=3, value=f"=ROUND($C$3*{pct},0)").number_format = money_fmt
        ws3.cell(row=r, column=4, value=f'=SUMIFS(سجل_المصروفات!F:F,سجل_المصروفات!D:D,"مصروف",سجل_المصروفات!C:C,B{r},سجل_المصروفات!B:B,TEXT($C$3,""))').number_format = money_fmt
        ws3.cell(row=r, column=5, value=f"=C{r}-D{r}").number_format = money_fmt
        ws3.cell(row=r, column=6, value=note).font = SMALL
        for cc in range(2, 7):
            ws3.cell(row=r, column=cc).border = box
        r += 1
last_data_row = r - 1

# savings block
ws3.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
c = ws3.cell(row=r, column=2, value="ادخار واستثمار 20% (الأولوية القصوى - ادفعه لنفسك أول الشهر)")
c.font = Font(name="Segoe UI", size=12, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=GREEN)
c.alignment = Alignment(horizontal="right")
r += 1
save_start = r
savings_items = [
    ("صندوق الطوارئ", "ادخار آمن 6 أشهر مصاريف - أولوية قصوى"),
    ("استثمار طويل الأجل", "70% من الادخار: أسهم/صناديق/ذهب"),
    ("أهداف خاصة", "سفر / عربية / مشروع / تأمين"),
]
for name, note in savings_items:
    ws3.cell(row=r, column=2, value=name).font = BODY
    ws3.cell(row=r, column=3, value=f"=ROUND($C$3*0.2/3,0)").number_format = money_fmt
    ws3.cell(row=r, column=5, value=f"=C{r}").number_format = money_fmt
    ws3.cell(row=r, column=6, value=note).font = SMALL
    for cc in range(2, 7):
        ws3.cell(row=r, column=cc).border = box
        ws3.cell(row=r, column=cc).fill = fill_green
    r += 1

r += 1
ws3.cell(row=r, column=2, value="الإجمالي المخطط:").font = BOLD
ws3.cell(row=r, column=3, value=f"=SUM(C{first_data_row}:C{save_start+2})-SUMIF(B{first_data_row}:B{last_data_row},\"\",0)").number_format = money_fmt
ws3.cell(row=r, column=3).font = BOLD
ws3.cell(row=r, column=3).fill = fill_gold
ws3.cell(row=r, column=3).border = box
ws3.cell(row=r, column=4, value="يجب = الراتب").font = SMALL

ws3.conditional_formatting.add(f"E{first_data_row}:E{last_data_row}",
    CellIsRule(operator="lessThan", formula=["0"], fill=fill_red))
ws3.sheet_view.showGridLines = False

# ==========================================================
# SHEET 4: الأهداف (Savings Goals)
# ==========================================================
ws4 = wb.create_sheet("الأهداف")
set_widths(ws4, [3, 32, 18, 18, 18, 14, 26])
title_bar(ws4, "أهداف الادخار والاستثمار", 7, 1)
header_row(ws4, 3, ["", "الهدف", "المستهدف", "المدخر شهرياً", "الحالي", "النسبة", "الموعد المتوقع"])

goals = [
    ("صندوق الطوارئ", 60000, 3000, 15000, "2027-06"),
    ("شقة/مقدم عقد", 300000, 5000, 20000, "2030-12"),
    ("استثمار طويل الأجل", 500000, 5000, 0, "2035-12"),
    ("سفر/ترفيه", 30000, 1500, 5000, "2027-09"),
]
for i, (name, target, monthly, current, deadline) in enumerate(goals):
    rr = 4 + i
    ws4.cell(row=rr, column=2, value=name).font = BODY
    ws4.cell(row=rr, column=3, value=target).number_format = money_fmt
    ws4.cell(row=rr, column=4, value=monthly).number_format = money_fmt
    ws4.cell(row=rr, column=5, value=current).number_format = money_fmt
    ws4.cell(row=rr, column=6, value=f"=IF(C{rr}=0,0,E{rr}/C{rr})").number_format = pct_fmt
    ws4.cell(row=rr, column=7, value=deadline).alignment = Alignment(horizontal="center")
    for cc in range(2, 8):
        ws4.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws4.cell(row=rr, column=cc).fill = fill_band

ws4.conditional_formatting.add("F4:F20",
    DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=GREEN))

ws4["B10"] = "قاعدة: صندوق الطوارئ أولاً (6 مصاريف) → سداد الديون → الاستثمار طويل الأجل"
ws4["B10"].font = GOLD_F
ws4.merge_cells("B10:G10")

# ==========================================================
# SHEET 5: الديون (Debt Tracker)
# ==========================================================
ws5 = wb.create_sheet("الديون")
set_widths(ws5, [3, 28, 16, 16, 16, 14, 16, 22])
title_bar(ws5, "خطة سداد الديون ( Avalanche = الأعلى فائدة أولاً )", 8, 1)
header_row(ws5, 3, ["", "الدين", "المتبقي", "القسط الشهري", "فائدة شهرية %", "شهر متوقع السداد", "أولوية", "ملاحظة"])

debts = [
    ("بطاقة ائتمان", 10000, 1000, 0.025, "احرق الفائدة أولاً"),
    ("قسط عربية", 80000, 3000, 0.009, "ثابت لحد نهاية المدة"),
    ("دين شخصي", 20000, 2000, 0.015, "تفاوض على خفض الفائدة"),
]
for i, (name, bal, pmt, rate, note) in enumerate(debts):
    rr = 4 + i
    ws5.cell(row=rr, column=2, value=name).font = BODY
    ws5.cell(row=rr, column=3, value=bal).number_format = money_fmt
    ws5.cell(row=rr, column=4, value=pmt).number_format = money_fmt
    ws5.cell(row=rr, column=5, value=rate).number_format = "0.0%"
    ws5.cell(row=rr, column=6, value=f"=IFERROR(ROUNDUP(-LOG(1-C{rr}/(D{rr}/E{rr}+C{rr}))/LOG(1+E{rr}),0),\"-\")").alignment = Alignment(horizontal="center")
    ws5.cell(row=rr, column=7, value=f"=RANK(E{rr},$E$4:$E$6,0)").alignment = Alignment(horizontal="center")
    ws5.cell(row=rr, column=8, value=note).font = SMALL
    for cc in range(2, 9):
        ws5.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws5.cell(row=rr, column=cc).fill = fill_band

ws5.cell(row=9, column=2, value="إجمالي الديون:").font = BOLD
ws5.cell(row=9, column=3, value="=SUM(C4:C6)").number_format = money_fmt
ws5.cell(row=9, column=3).font = Font(name="Segoe UI", bold=True, color=RED)
ws5.cell(row=9, column=3).fill = fill_red
ws5.cell(row=9, column=3).border = box
ws5.cell(row=9, column=4, value="=SUM(D4:D6)").number_format = money_fmt
ws5.cell(row=9, column=4).font = BOLD
ws5.cell(row=9, column=4).border = box

ws5["B12"] = "استراتيجية الأفالانش: سدّد أكتر فائدة (بطاقة ائتمان 2.5% شهرياً) بأقصى مبلغ، وثبّت الأقساط الباقية."
ws5["B12"].font = GOLD_F
ws5.merge_cells("B12:H12")
ws5["B13"] = "بمجرد ما تقفل دين → فوّت قسطه كله للدين التالي (سلسلة الثلج). ه توفّر ألاف."
ws5["B13"].font = SMALL
ws5.merge_cells("B13:H13")
ws5.sheet_view.showGridLines = False

# ==========================================================
# SHEET 6: صافي الثروة (Net Worth)
# ==========================================================
ws6 = wb.create_sheet("صافي_الثروة")
set_widths(ws6, [3, 30, 18, 18, 18, 18, 18])
title_bar(ws6, "تتبع صافي الثروة شهرياً (الهدف: رقم يطلع لفوق دائماً)", 7, 1)
header_row(ws6, 3, ["", "الشهر", "نقد وأرصدة", "استثمارات", "عقارات/سيارات", "إجمالي الأصول", "صافي الثروة"])

months = ["2026-09", "2026-10", "2026-11", "2026-12", "2027-01", "2027-02"]
for i, m in enumerate(months):
    rr = 4 + i
    ws6.cell(row=rr, column=2, value=m).alignment = Alignment(horizontal="center")
    for cc, val in [(3, 0), (4, 0), (5, 0)]:
        ws6.cell(row=rr, column=cc, value=val).number_format = money_fmt
    ws6.cell(row=rr, column=6, value=f"=SUM(C{rr}:E{rr})").number_format = money_fmt
    ws6.cell(row=rr, column=7, value=f"=F{rr}-IFERROR(الديون!C9,0)").number_format = money_fmt
    for cc in range(2, 8):
        ws6.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws6.cell(row=rr, column=cc).fill = fill_band
ws6.conditional_formatting.add("G4:G40",
    CellIsRule(operator="lessThan", formula=["0"], fill=fill_red))
ws6.sheet_view.showGridLines = False

# ==========================================================
# SHEET 7: قواعد ذهبية (Golden Rules) - reference
# ==========================================================
ws7 = wb.create_sheet("قواعد_ذهبية")
set_widths(ws7, [4, 6, 90])
title_bar(ws7, "القواعد الذهبية لبناء الثروة", 3, 1)
rules = [
    "ادفع لنفسك الأول: أول ما يدخل الراتب، حوّل 20% فوراً للادخار/الاستثمار قبل أي مصروف.",
    "قاعدة 50/30/20: 50% أساسيات، 30% رغبات، 20% ادخار — وعدّل النسب لحالتك.",
    "صندوق الطوارئ أولاً: 6 أشهر مصاريف في حساب م Aktuel عالي الفائدة قبل أي استثمار.",
    "لا تدخل ديناً استهلاكياً: أي فائدة فوق 1% شهرياً = ثروة بتولّع. بطاقة الائتمان عدوّك.",
    "الأفالانش للديون: سدّد الأعلى فائدة أولاً بأقصى مبلغ.",
    "ال Inflate Lifestyle ممنوع: لما الراتب يزيد، زوّد الادخار 70% منها والرغبات 30% فقط.",
    "استثمر بانتظام (DCA): مبلغ ثابت كل شهر مهما كان السوق — الانضباط يهزم التوقّع.",
    "نوّع: 70% طويل الأجل (أسهم/صناديق) + 20% آمن (خزينة/سندات) + 10% سيولة/ذهب.",
    "راجع أرقامك معايا كل شهر: أول 3 أيام من كل شهر نبعتلك التقرير ونعدّل.",
    "الدخل هو مفتاح الثروة: 50% من طاقتك لزيادة الدخل (مهارات/عمل حر/مشروع) مش فقط لتوفير.",
    "اشتري أصول: كل جنيه يشتري أصل (سهم، ذهب، أرض، مهارة) = يشتري مستقبلك. المشتريات تهبط.",
    "التأمين والصحة: تأمين طبي + تأمين حياة يحمي ثروتك من كارثة واحدة.",
]
for i, rule in enumerate(rules):
    rr = 3 + i
    ws7.cell(row=rr, column=1, value=f"{i+1}.").font = Font(name="Segoe UI", size=12, bold=True, color=GOLD)
    ws7.cell(row=rr, column=1).alignment = Alignment(horizontal="center")
    ws7.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    c = ws7.cell(row=rr, column=2, value=rule)
    c.font = BODY
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws7.row_dimensions[rr].height = 32
    for cc in [1, 2, 3]:
        ws7.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws7.cell(row=rr, column=cc).fill = fill_band
ws7.sheet_view.showGridLines = False

# ==========================================================
# SHEET 8: تعليمات (How to use)
# ==========================================================
ws8 = wb.create_sheet("كيف_تستخدم")
set_widths(ws8, [4, 100])
title_bar(ws8, "كيف تستخدم النظام ده + إيه اللي تبعته للمساعد", 2, 1)
steps = [
    "1) كل يوم: افتح شيت 'سجل_المصروفات' واكتب كل عملية — دخل أو مصروف — في أقل من دقيقة.",
    "2) في أول الشهر: حدّد الراتب في خلية C3 بشيت 'الميزانية'، وهتتحسب الحدود تلقائياً بقاعدة 50/30/20.",
    "3) كل أسبوع: افتح 'لوحة التحكم' وشوف أنت زوّدت عن أنهي فئة (الأحمر = تجاوزت).",
    "4) أول 3 أيام من كل شهر: بعتلي رقمين — إجمالي الدخل وإجمالي المصروفات الشهر اللي فات — وأنا أعملك التقرير والخطة.",
    "5) 'الأهداف': حدّث رقم 'الحالي' كل شهر عشان نشوف نسبة الإنجاز.",
    "6) 'الديون': حدّث المتبقي بعد كل قسط، والأولوية بتتحسب تلقائياً.",
    "7) 'صافي_الثروة': سجل أصولك آخر كل شهر — الرقم ده هو مقياس نجاحك الحقيقي.",
    "8) 'قواعد_ذهبية': اقرأها مرة كل أسبوع لحد ما تبقى غريزة.",
    "9) اللي تبعتهولي: صورة كشف الحساب أو قائمة مصروفات الشهر كاملة (أي شكل: نص، صورة، إكسل) — وأنا أرتّبها وأحللها.",
    "10) كل شهر أديك: تقرير + ميزانية الشهر الجديد + نصائح مخصصة + تعديل على الخطة.",
]
for i, s in enumerate(steps):
    rr = 3 + i
    c = ws8.cell(row=rr, column=2, value=s)
    c.font = BODY
    c.alignment = Alignment(wrap_text=True, vertical="center")
    ws8.row_dimensions[rr].height = 34
    for cc in [1, 2]:
        ws8.cell(row=rr, column=cc).border = box
        if i % 2 == 1:
            ws8.cell(row=rr, column=cc).fill = fill_band
ws8.sheet_view.showGridLines = False

wb.save(OUT)
print("SAVED:", OUT)
print("Sheets:", wb.sheetnames)

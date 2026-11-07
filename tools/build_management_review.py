#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ساخت فایل اکسل بازنگری مدیریت (IMS) برای شرکت‌های تولیدی قطعات خودرو و مواد کامپوزیتی.

مبنای الزامات (استخراج‌شده از متن/خلاصه بند 9.3 استانداردها):
  ISO 9001:2026      بندهای 9.3.1 تا 9.3.3
  ISO 14001:2026     بندهای 9.3.1 تا 9.3.3
  ISO 45001:2018     بند 9.3
  IATF 16949:2016    بندهای 9.3.1.1 ، 9.3.2.1 (a تا k) و 9.3.3.1

خروجی: IMS_Management_Review_FA_v1.xlsx
اجرا:  python3 tools/build_management_review.py
"""

import datetime
import math
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.workbook.defined_name import DefinedName

OUT = "IMS_Management_Review_FA_v1.xlsx"

S1 = "1-جلسه و تنظیمات"
S2 = "2-ورودی‌ها"
S3 = "3-خروجی و مصوبات"
S4 = "4-پایش و اثربخشی"
S5 = "5-صورتجلسه"

FONT_NAME = "Tahoma"

NAVY = "1F3864"
HEAD = "1F4E79"
SUBHEAD = "2E75B6"
SECTION = "D9E2F3"
YELLOW = "FFF2CC"
BLUE_CALC = "DEEAF6"
GREY = "F2F2F2"
GREEN = "E2EFDA"
RED = "FCE4E4"
AMBER = "FFF2CC"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

RTL = 2  # readingOrder: راست‌به‌چپ


def fill(color):
    return PatternFill("solid", start_color=color, end_color=color)


def align(h="right", v="center", wrap=True, size=None):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, readingOrder=RTL)


def title_cell(ws, rng, text, size=15, color=NAVY, fg="FFFFFF", height=30):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    c.font = Font(name=FONT_NAME, size=size, bold=True, color=fg)
    c.fill = fill(color)
    c.alignment = align("center", "center")
    ws.row_dimensions[c.row].height = height


def section(ws, rng, text):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    c.font = Font(name=FONT_NAME, size=11, bold=True, color=NAVY)
    c.fill = fill(SECTION)
    c.alignment = align("right", "center")
    ws.row_dimensions[c.row].height = 22


def merge_text(ws, rng, text, style="normal", height=None):
    first = rng.split(":")[0]
    ws.merge_cells(rng)
    c = ws[first]
    c.value = text
    if style == "normal":
        c.font = Font(name=FONT_NAME, size=9)
        c.alignment = align()
    elif style == "note":
        c.font = Font(name=FONT_NAME, size=8, italic=True, color="595959")
        c.alignment = align()
    if height:
        ws.row_dimensions[c.row].height = height
    return c




def style_range(ws, rng, border=True, fill_color=None):
    """اعمال کادر و رنگ پس‌زمینه روی همه سلول‌های یک محدوده (برای سلول‌های ادغام‌شده)."""
    for row in ws[rng]:
        for c in row:
            if border:
                c.border = BORDER
            if fill_color:
                c.fill = fill(fill_color)



def autofit_rows(ws, rows, col_specs, pad=4, min_h=18, max_h=230):
    """ارتفاع سطر را بر پایه طول متن و عرض ستون‌ها (با در نظر گرفتن ادغام) تخمین می‌زند."""
    widths = {}
    for letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        w = ws.column_dimensions[letter].width
        widths[letter] = w if w else 8.43
    merged_width = {}
    for m in ws.merged_cells.ranges:
        if m.min_row == m.max_row:
            merged_width[(m.min_row, m.min_col)] = sum(
                widths.get(get_column_letter(c), 8.43) for c in range(m.min_col, m.max_col + 1))
    line_h = {9: 12.2, 8: 11.2, 7: 10.2, 6: 9.6}
    factor = {9: 1.12, 8: 1.28, 7: 1.5, 6: 1.7}
    for r in rows:
        need = min_h
        for col_letter, size in col_specs.items():
            v = ws[f"{col_letter}{r}"].value
            if not isinstance(v, str) or v.startswith("="):
                continue
            w = merged_width.get((r, ws[f"{col_letter}{r}"].column), widths.get(col_letter, 8.43))
            cpl = max(6, w * factor.get(size, 1.12))
            lines = sum(max(1, math.ceil(len(seg) / cpl)) for seg in v.split("\n"))
            need = max(need, lines * line_h.get(size, 12.2) + pad)
        ws.row_dimensions[r].height = min(max_h, round(need, 1))


wb = Workbook()
wb.remove(wb.active)

ws1 = wb.create_sheet(S1)
ws2 = wb.create_sheet(S2)
ws3 = wb.create_sheet(S3)
ws4 = wb.create_sheet(S4)
ws5 = wb.create_sheet(S5)

for ws in wb.worksheets:
    ws.sheet_view.rightToLeft = True
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)

# ----------------------------------------------------------------------------
# برگه ۱ – جلسه و تنظیمات
# ----------------------------------------------------------------------------
ws = ws1
for col, w in zip("ABCDEF", [5, 30, 34, 26, 18, 16]):
    ws.column_dimensions[col].width = w

title_cell(ws, "A1:F1", "بازنگری مدیریت سیستم‌های مدیریت یکپارچه (IMS)", size=16, height=34)
title_cell(
    ws, "A2:F2",
    "ISO 9001:2026  ·  ISO 14001:2026  ·  ISO 45001:2018  ·  IATF 16949:2016   |   صنعت قطعات خودرو و مواد کامپوزیتی",
    size=9, color=SUBHEAD, height=22,
)
title_cell(
    ws, "A3:F3",
    "کد فرم: MR-IMS-F01  ·  نسخه 1.0  ·  تاریخ تدوین: 2026-11-07  ·  تهیه: تضمین کیفیت / IMS",
    size=8, color=GREY, fg="404040", height=18,
)

section(ws, "A5:F5", "۱) اطلاعات جلسه و دوره بازنگری")

FIELDS = [
    ("نام سازمان / شرکت", None, "نام حقوقی مطابق گواهینامه"),
    ("واحد / محل استقرار", None, "کارخانه یا سایت مورد بازنگری"),
    ("دامنه سیستم (محصول و فرآیند)", None,
     "مطابق دامنه گواهینامه؛ مثال: قطعات خودرو و مواد کامپوزیتی – قالب‌گیری، ماشین‌کاری، مونتاژ"),
    ("دوره بازنگری (کد دوره)", None, "مثال: 1404-H2 یا 2026-H2 — برای هر دوره یک کد ثابت"),
    ("شماره صورت‌جلسه", None, "مثال: MR-1404-H2-01"),
    ("نوع بازنگری", "list:دوره‌ای برنامه‌ریزی‌شده,فوق‌العاده,پس از تغییر عمده,پیگیری مصوبات", "مبنای بند 9.3.1"),
    ("تاریخ جلسه", "date", "میلادی"),
    ("ساعت شروع و پایان", None, "مثال: 09:00 تا 12:30"),
    ("محل برگزاری", None, None),
    ("رئیس جلسه (نماینده ارشد مدیریت)", None, "الزام مشارکت مدیریت ارشد – بند 5.1"),
    ("دبیر جلسه", None, "مسئول ثبت صورت‌جلسه و مصوبات"),
    ("بازه داده‌های مورد بررسی", None, "مثال: از 1404/01/01 تا 1404/06/31 — همه ورودی‌ها باید همین بازه را پوشش دهند"),
    ("تاریخ بازنگری قبلی", "date", "مبنای ورودی «وضعیت مصوبات قبلی»"),
    ("تاریخ بازنگری بعدی", "date", "برنامه‌ریزی‌شده در همین جلسه"),
    ("تواتر بازنگری", None,
     "IATF 16949 بند 9.3.1.1: حداقل سالانه و افزایش تواتر بر پایه ریسک عدم انطباق با الزامات مشتری"),
]

first_field_row = 6
for i, (label, dv, note) in enumerate(FIELDS):
    r = first_field_row + i
    c = ws.cell(row=r, column=2, value=label)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align()
    c.fill = fill(GREY)
    c.border = BORDER

    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    v = ws.cell(row=r, column=3)
    v.fill = fill(YELLOW)
    v.border = BORDER
    v.font = Font(name=FONT_NAME, size=9)
    v.alignment = align("center" if dv else "right")
    if dv == "date":
        v.number_format = "yyyy-mm-dd"
    if dv and dv.startswith("list:"):
        d = DataValidation(type="list", formula1='"' + dv[5:] + '"', allow_blank=True)
        ws.add_data_validation(d)
        d.add(f"C{r}:D{r}")
    elif dv == "date":
        d = DataValidation(type="date", operator="between",
                           formula1="DATE(2024,1,1)", formula2="DATE(2045,12,31)", allow_blank=True)
        ws.add_data_validation(d)
        d.add(f"C{r}:D{r}")
    ws.cell(row=r, column=4).border = BORDER

    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    n = ws.cell(row=r, column=5, value=note)
    n.font = Font(name=FONT_NAME, size=8, italic=True, color="595959")
    n.alignment = align()
    n.border = BORDER
    ws.cell(row=r, column=6).border = BORDER
    ws.row_dimensions[r].height = 24 if not note or len(note) < 60 else 32

# حاضرین
r = 22
section(ws, f"A{r}:F{r}", "۲) حاضرین جلسه و نقش‌ها")
r += 1
headers = ["ردیف", "نام و نام خانوادگی", "سمت / واحد سازمانی", "نقش در جلسه", "امضا", ""]
for j, h in enumerate(headers[:5]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    c.border = BORDER
ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
ws.cell(row=r, column=5).value = "امضا"
ws.cell(row=r, column=6).border = BORDER
ws.row_dimensions[r].height = 22

attendee_dv = DataValidation(
    type="list", formula1='"رئیس جلسه,دبیر جلسه,عضو,مسئول داده,مدعو"', allow_blank=True)
ws.add_data_validation(attendee_dv)
for i in range(10):
    rr = r + 1 + i
    ws.cell(row=rr, column=1, value=i + 1).font = Font(name=FONT_NAME, size=9)
    ws.cell(row=rr, column=1).alignment = align("center")
    ws.cell(row=rr, column=1).border = BORDER
    for j in (2, 3, 4):
        c = ws.cell(row=rr, column=j)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=9)
        c.alignment = align()
    ws.merge_cells(start_row=rr, start_column=5, end_row=rr, end_column=6)
    ws.cell(row=rr, column=6).border = BORDER
    attendee_dv.add(f"D{rr}")
    ws.row_dimensions[rr].height = 20

# استانداردها
r = 35
section(ws, f"A{r}:F{r}", "۳) استانداردهای مرجع، نسخه معتبر و دامنه کاربرد")
r += 1
std_headers = ["استاندارد", "نسخه معتبر (تاریخ انتشار)", "وضعیت گواهینامه سازمان", "مهلت / برنامه انتقال", "بند بازنگری مدیریت"]
for j, h in enumerate(std_headers):
    c = ws.cell(row=r, column=2 + j, value=h)
    c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    c.border = BORDER
ws.cell(row=r, column=1).border = BORDER
ws.row_dimensions[r].height = 30

STDS = [
    ("ISO 9001:2026", "ویرایش ششم – انتشار 2026-09-16",
     "بند 9.3.1 تا 9.3.3"),
    ("ISO 14001:2026", "ویرایش چهارم – انتشار 2026-04-15",
     "بند 9.3.1 تا 9.3.3"),
    ("ISO 45001:2018 (+Amd 1:2024)", "جاری؛ بازنگری در مرحله DIS – انتشار احتمالی 2027",
     "بند 9.3"),
    ("IATF 16949:2016", "جاری؛ ویرایش دوم برنامه‌ریزی‌شده میانه 2027 – آخرین SI انتشار یافته 2026",
     "بندهای 9.3.1.1 ، 9.3.2.1 ، 9.3.3.1"),
]
TRANS = [
    "انتقال تا 2029-09-30",
    "انتقال حدود 2029-04 (بر پایه سند انتقال IAF؛ تأیید با گواهی‌دهنده)",
    "تا انتشار ویرایش جدید، 2018 مبنای گواهینامه است",
    "پایان انتقال هم‌زمان با انتقال ISO 9001:2026",
]
for i, (std, ver, clause) in enumerate(STDS):
    rr = r + 1 + i
    ws.cell(row=rr, column=1).border = BORDER
    c = ws.cell(row=rr, column=2, value=std)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align()
    c.border = BORDER
    c.fill = fill(GREY)
    c = ws.cell(row=rr, column=3, value=ver)
    c.font = Font(name=FONT_NAME, size=9)
    c.alignment = align()
    c.border = BORDER
    c = ws.cell(row=rr, column=4)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.alignment = align()
    c.font = Font(name=FONT_NAME, size=9)
    c = ws.cell(row=rr, column=5, value=TRANS[i])
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align()
    c.border = BORDER
    c = ws.cell(row=rr, column=6, value=clause)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align()
    c.border = BORDER
    ws.row_dimensions[rr].height = 30

merge_text(
    ws, "B41:F41",
    "وضعیت نسخه‌ها بر پایه اطلاعات عمومی منتشرشده تا 2026-11-07 است. «انتشار ویرایش جدید» با «نسخه معتبر گواهینامه سازمان»"
    " یکسان نیست؛ پیش از استناد، متن رسمی خریداری‌شده و دامنه گواهینامه خود را ملاک قرار دهید.",
    style="note", height=26,
)

# کنترل تکمیل
r = 43
section(ws, f"A{r}:F{r}", "۴) کنترل تکمیل و وضعیت جلسه (خودکار)")
r += 1
for j, h in enumerate(["شاخص", "مقدار", "منبع / توضیح"]):
    c = ws.cell(row=r, column=2 + j, value=h)
    c.font = Font(name=FONT_NAME, size=9, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    c.border = BORDER
ws.cell(row=r, column=1).border = BORDER
ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
ws.cell(row=r, column=4).border = BORDER

METRICS = [
    ("تعداد ورودی‌های الزامی تعریف‌شده", f"=COUNTA('{S2}'!$B$5:$B$33)", "0", "برگه ۲ – همیشه 29 ورودی"),
    ("ورودی‌های دارای نتیجه بررسی", f'=COUNTIF(\'{S2}\'!$H$5:$H$33,"?*")', "0", "ستون «نتیجه بررسی» تکمیل شده"),
    ("درصد پوشش ورودی‌ها", None, "0%", "باید پیش از جلسه 100% شود"),
    ("ورودی‌های ناقص یا ارائه‌نشده", f'=COUNTIF(\'{S2}\'!$G$5:$G$33,"ناقص")+COUNTIF(\'{S2}\'!$G$5:$G$33,"ارائه نشد")', "0", "اقدام پیش از جلسه"),
    ("مصوبات ثبت‌شده", f'=COUNTIF(\'{S3}\'!$H$5:$H$44,"?*")', "0", "برگه ۳ – ستون تصمیم یا مصوبه"),
    ("مصوبات انجام‌شده", f'=COUNTIF(\'{S3}\'!$L$5:$L$44,"انجام‌شده")', "0", None),
    ("مصوبات باز یا در حال اجرا", f'=COUNTIF(\'{S3}\'!$L$5:$L$44,"باز")+COUNTIF(\'{S3}\'!$L$5:$L$44,"در حال اجرا")', "0", None),
    ("مصوبات معوق", f'=COUNTIF(\'{S3}\'!$L$5:$L$44,"معوق")', "0", "نیازمند تصمیم در جلسه بعد"),
    ("درصد اجرای مصوبات", None, "0%", "انجام‌شده ÷ کل مصوبات"),
    ("مصوبات با اثربخشی تأییدشده", f'=COUNTIF(\'{S3}\'!$N$5:$N$44,"اثربخش")+COUNTIF(\'{S3}\'!$N$5:$N$44,"نسبتاً اثربخش")', "0", None),
    ("تصمیم «نیاز به اقدام نیست» با توجیه", f'=COUNTIF(\'{S3}\'!$L$5:$L$44,"نیاز به اقدام نیست")', "0", "ستون توجیه باید پر باشد"),
    ("تعداد پایش‌های ثبت‌شده", f'=COUNTIF(\'{S4}\'!$B$5:$B$44,"?*")', "0", "برگه ۴"),
]

metric_rows = {}
for i, (label, formula, numfmt, note) in enumerate(METRICS):
    rr = r + 1 + i
    metric_rows[label] = rr
    ws.cell(row=rr, column=1).border = BORDER
    c = ws.cell(row=rr, column=2, value=label)
    c.font = Font(name=FONT_NAME, size=9)
    c.alignment = align()
    c.border = BORDER
    c = ws.cell(row=rr, column=3)
    c.border = BORDER
    c.fill = fill(BLUE_CALC)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align("center")
    c.number_format = numfmt
    if formula:
        c.value = formula
    ws.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=6)
    c = ws.cell(row=rr, column=4, value=note)
    c.font = Font(name=FONT_NAME, size=8, italic=True, color="595959")
    c.alignment = align()
    c.border = BORDER
    ws.cell(row=rr, column=6).border = BORDER
    ws.row_dimensions[rr].height = 18

row_inputs = metric_rows["تعداد ورودی‌های الزامی تعریف‌شده"]
row_checked = metric_rows["ورودی‌های دارای نتیجه بررسی"]
row_decisions = metric_rows["مصوبات ثبت‌شده"]
row_done = metric_rows["مصوبات انجام‌شده"]
ws1.cell(row=metric_rows["درصد پوشش ورودی‌ها"], column=3).value = (
    f"=IF(C{row_inputs}=0,0,C{row_checked}/C{row_inputs})")
ws1.cell(row=metric_rows["درصد اجرای مصوبات"], column=3).value = (
    f"=IF(C{row_decisions}=0,0,C{row_done}/C{row_decisions})")

# رنگ‌ها و قواعد
r = 58
section(ws, f"A{r}:F{r}", "۵) علامت‌گذاری رنگ‌ها و قواعد تکمیل")
legend = [
    ("سلول زرد", "ورود دستی اطلاعات؛ توسط تیم تهیه تکمیل شود.", YELLOW),
    ("سلول آبی روشن", "فرمول خودکار؛ تغییر داده نشود.", BLUE_CALC),
    ("سلول خاکستری", "عنوان یا برچسب ثابت فرم.", GREY),
]
rr = r + 1
for text, desc, color in legend:
    c = ws.cell(row=rr, column=2, value=text)
    c.fill = fill(color)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align()
    c.border = BORDER
    ws.cell(row=rr, column=1).border = BORDER
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
    c = ws.cell(row=rr, column=3, value=desc)
    c.font = Font(name=FONT_NAME, size=9)
    c.alignment = align()
    c.border = BORDER
    ws.cell(row=rr, column=6).border = BORDER
    rr += 1

rules = [
    "سبز: انجام‌شده یا اثربخش  ·  قرمز: معوق یا غیراثربخش  ·  زرد: باز یا در انتظار  ·  خاکستری: نیاز به اقدام نیست.",
    "برای افزودن ردیف جدید، فرمول‌ها و اعتبارسنجی ردیف آخر را کپی کنید؛ محدوده‌های کنترل تا ردیف 44 (برگه ۳ و ۴) و 33 (برگه ۲) است.",
    "تاریخ‌ها میلادی و با قالب yyyy-mm-dd ثبت شوند؛ محاسبه «روز باقی‌مانده» در برگه ۴ بر پایه تاریخ امروز است.",
    "این فایل هیچ داده واقعی یا ساختگی در جداول ندارد؛ تنها نمونه‌های آموزشی بخش ۶ در همین برگه است و باید بازنویسی شود.",
]
for text in rules:
    merge_text(ws, f"B{rr}:F{rr}", "• " + text, height=26)
    rr += 1

# نمونه آموزشی
r = 66
section(ws, f"A{r}:F{r}", "۶) نمونه نگارش (صرفاً آموزشی و غیرواقعی – در فایل عملیاتی باید بازنویسی یا حذف شود)")
samples = [
    ("نمونه ورودی", "IN-12 عدم انطباق‌ها و اقدامات اصلاحی",
     "نتیجه بررسی: 23 عدم انطباق در دوره ثبت شده؛ 31٪ تکراری در فرآیند ماشین‌کاری (انحراف ابعادی)؛ 4 اقدام اصلاحی معوق با تأخیر بیش از 30 روز."),
    ("نمونه خروجی", "OUT-05 اقدامات لازم برای اهداف محقق‌نشده",
     "تصمیم: اجرای اقدام اصلاحی ریشه‌ای برای تلرانس ابعادی قطعه X تا پایان دوره بعد | مسئول: مدیر مهندسی | منابع: 1 نفر-ماه + 15 میلیون تومان | مهلت: 1404/09/15."),
    ("نمونه خروجی (عدم اقدام)", "OUT-04 نیازهای منابع",
     "تصمیم: نیاز به اقدام نیست | توجیه: ظرفیت آزمایشگاه پاسخگوی حجم نمونه‌های دوره بود و شاخص زمان پاسخ زیر هدف باقی ماند."),
]
rr = r + 1
for kind, subject, text in samples:
    c = ws.cell(row=rr, column=2, value=f"{kind}: {subject}")
    c.font = Font(name=FONT_NAME, size=8, bold=True, italic=True, color="7F6000")
    c.alignment = align()
    c.fill = fill(AMBER)
    c.border = BORDER
    ws.cell(row=rr, column=1).border = BORDER
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=6)
    c = ws.cell(row=rr, column=3, value=text)
    c.font = Font(name=FONT_NAME, size=8, italic=True)
    c.alignment = align()
    c.border = BORDER
    ws.cell(row=rr, column=6).border = BORDER
    ws.row_dimensions[rr].height = 34
    rr += 1

ws1.freeze_panes = "A6"
ws1.print_area = "A1:F70"
ws1.page_setup.orientation = "portrait"
ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
ws1.page_setup.fitToWidth = 1
ws1.page_setup.fitToHeight = 0
ws1.print_options.horizontalCentered = True
ws1.oddFooter.center.text = "بازنگری مدیریت IMS – برگه ۱ از ۵ – صفحه &P از &N"
ws1.oddFooter.center.size = 8

# ----------------------------------------------------------------------------
# برگه ۲ – ورودی‌های بازنگری مدیریت
# ----------------------------------------------------------------------------
ws = ws2
widths2 = {"A": 4, "B": 9, "C": 12, "D": 9, "E": 30, "F": 58, "G": 13, "H": 42, "I": 10, "J": 26, "K": 30, "L": 9}
for col, w in widths2.items():
    ws.column_dimensions[col].width = w

title_cell(ws, "A1:L1", "ورودی‌های بازنگری مدیریت – جدول یکپارچه الزامات", size=13, height=28)
title_cell(
    ws, "A2:L2",
    "29 ورودی الزامی؛ الزامات مشترک ادغام و الزامات اختصاصی محیط‌زیست، ایمنی و IATF جداگانه حفظ شده است. "
    "«9001» = ISO 9001:2026 ، «14001» = ISO 14001:2026 ، «45001» = ISO 45001:2018 ، «IATF» = IATF 16949:2016",
    size=8, color=SUBHEAD, height=26,
)
title_cell(
    ws, "A3:L3",
    "ستون‌های زرد = ورود دستی   |   ستون آبی = فرمول   |   برای هر ورودی حداقل یک نتیجه بررسی ثبت شود",
    size=8, color=GREY, fg="404040", height=18,
)

H2 = [
    ("ردیف", "شماره سطر"),
    ("کد ورودی", "شناسه ثابت برای ردیابی"),
    ("دوره", "کد دوره بازنگری"),
    ("سیستم", "QMS / EMS / OHS / IMS"),
    ("موضوع ورودی", "عنوان الزام"),
    ("اطلاعات موردنیاز و شاخص‌های کلیدی", "داده، سند و تحلیلی که باید ارائه شود"),
    ("وضعیت ارائه", "ارائه شد / ناقص / ارائه نشد / مصداق ندارد"),
    ("نتیجه بررسی و تحلیل مدیریت", "روند، انحراف از هدف، علت و جمع‌بندی"),
    ("نیاز به مصوبه؟", "بله / خیر"),
    ("مستندات و مرجع داده", "نام سند، شماره، مسئول داده"),
    ("بند استاندارد مرتبط", "9001 / 14001 / 45001 / IATF"),
    ("تعداد مصوبات مرتبط", "فرمول – برگه ۳"),
]
HROW2 = 4
for j, (h, sub) in enumerate(H2, start=1):
    c = ws.cell(row=HROW2, column=j, value=h + "\n" + f"({sub})")
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center", "center")
    c.border = BORDER
ws.row_dimensions[HROW2].height = 40

INPUTS = [
    ("IN-01", "IMS", "وضعیت مصوبات بازنگری قبلی",
     "فهرست مصوبات دوره قبل با کد، مسئول، مهلت، درصد پیشرفت، وضعیت و شاهد اجرا؛ مصوبات معوق و علت تأخیر؛ "
     "مصوبات منتقل‌شده به دوره جاری؛ نتیجه اثربخشی مصوبات بسته‌شده.",
     "9001: 9.3.2 a)  ·  14001: 9.3.2 a)  ·  45001: 9.3.2 a)"),
    ("IN-02", "IMS", "تغییرات مسائل داخلی و خارجی",
     "تغییر ساختار و مالکیت، چیدمان یا ظرفیت خطوط (قالب‌گیری، ماشین‌کاری، مونتاژ)، محصول یا قالب جدید، "
     "تغییر تأمین‌کننده رزین/الیاف/فولاد، نوسان انرژی و نرخ ارز، محدودیت زنجیره تأمین، رویدادهای اقلیمی مؤثر بر تولید و آب، "
     "قوانین جدید محیط‌زیستی و کار، تحریم و لجستیک.",
     "9001: 9.3.2 b)  ·  14001: 9.3.2 b) 1  ·  45001: 9.3.2 b)"),
    ("IN-03", "IMS", "تغییرات نیازها و انتظارات ذی‌نفعان (شامل الزامات مشتریان خودروساز)",
     "به‌روزرسانی الزامات اختصاصی مشتریان (CSR) و پورتال‌ها، تغییر مشخصات و تلرانس‌ها، الزامات ممیزی مشتری، "
     "انتظارات کارکنان و پیمانکاران، الزامات سازمان‌های نظارتی و محلی، مطالبات سهامداران و بیمه؛ تغییر فهرست ذی‌نفعان.",
     "9001: 9.3.2 c)  ·  14001: 9.3.2 b) 2  ·  45001: 9.3.2 b) 1"),
    ("IN-04", "IMS", "تغییرات الزامات قانونی و نتیجه ارزیابی انطباق",
     "قوانین جدید یا اصلاح‌شده محیط‌زیست (پساب، هوا، پسماند، مواد شیمیایی و VOC)، ایمنی و بهداشت کار، آتش‌نشانی و بیمه؛ "
     "نتیجه ارزیابی انطباق و موارد عدم انطباق؛ مجوزها و مهلت تمدید؛ الزامات قانونی مشتریان.",
     "14001: 9.3.2 b) 2 و d) 3  ·  45001: 9.3.2 b) 2 و d) 3  ·  9001: 9.3.2 b) و c)"),
    ("IN-05", "IMS", "تغییرات الزامات استانداردها و برنامه انتقال گواهینامه",
     "وضعیت گواهینامه‌های جاری و تاریخ انقضا، الزامات ویرایش‌های جدید (ISO 9001:2026 و ISO 14001:2026)، زمان‌بندی انتقال، "
     "شکاف‌های شناسایی‌شده و برنامه رفع، برنامه ویرایش دوم IATF 16949 و تغییرات CSR، آموزش‌های لازم.",
     "9001: 9.3.2 b)  ·  14001: 9.3.2 b) 1  ·  45001: 9.3.2 b)"),
    ("IN-06", "QMS", "تحقق اهداف کیفیت",
     "هر هدف کیفیت با هدف کمی، مقدار واقعی، روند سه دوره، درصد تحقق، علت انحراف و اقدام انجام‌شده؛ "
     "اهداف محقق‌نشده و تصمیم موردنیاز.",
     "9001: 9.3.2 d) 5"),
    ("IN-07", "QMS", "عملکرد، اثربخشی و کارایی فرآیندها",
     "شاخص هر فرآیند (OEE، ضایعات و دوباره‌کاری، بهره‌وری، زمان چرخه، تحویل به‌موقع، توقف خط)، مقایسه با هدف و روند؛ "
     "اثربخشی (تحقق خروجی مورد انتظار) و کارایی (نسبت خروجی به منابع) فرآیندهای تولید؛ ظرفیت گلوگاه.",
     "9001: 9.3.2 d) 6  ·  IATF: 9.3.2.1 b) و c)"),
    ("IN-08", "QMS", "رضایت مشتری، شکایات و بازخورد",
     "نتیجه نظرسنجی یا ارزیابی رضایت، تعداد و نوع شکایات، 8Dهای باز و بسته، زمان پاسخ، پذیرش مشتری، "
     "بازخورد مثبت و منفی، رضایت از تحویل و پشتیبانی فنی.",
     "9001: 9.3.2 d) 4  ·  IATF: 9.3.2.1 f)"),
    ("IN-09", "QMS", "کارت امتیاز مشتری و عملکرد تحویل",
     "امتیاز مشتری در پورتال، وضعیت ویژه یا مسدود، PPM تحویل، عملکرد تحویل به‌موقع (OTD)، هزینه حمل اضطراری، "
     "اطلاع‌رسانی مشتری درباره کیفیت و تحویل، روند و علت افت امتیاز.",
     "IATF: 9.3.2.1 i)  ·  9001: 9.1.2"),
    ("IN-10", "QMS", "انطباق محصول و کیفیت قطعه تحویلی",
     "نرخ قطعات معیوب داخلی و برگشتی، PPAP و تأیید نمونه، انطباق با مشخصات محصول و استانداردهای مشتری، "
     "نتایج آزمایشگاه و آزمون‌های مخرب/غیرمخرب، انحرافات مجاز (Concession)، برچسب‌گذاری و بسته‌بندی.",
     "9001: 9.3.2 d) 6  ·  IATF: 9.3.2.1 d)"),
    ("IN-11", "IMS", "نتایج پایش و اندازه‌گیری",
     "نتایج بازرسی و آزمون، کالیبراسیون تجهیزات و MSA، پایش فرآیند (SPC/CPK)، پایش محیط کار (صدا، روشنایی، گردوغبار، "
     "بخارات رزین و استایرن، دما)، پایش محیط‌زیستی (پساب، هوا، پسماند، انرژی، آب) و پایش ایمنی (بازدید، رفتار، تجهیزات حفاظتی).",
     "9001: 9.3.2 d) 2  ·  14001: 9.3.2 d) 2  ·  45001: 9.3.2 d) 2"),
    ("IN-12", "IMS", "عدم انطباق‌ها و اقدامات اصلاحی",
     "تعداد و روند عدم انطباق‌ها به تفکیک منشأ (داخلی، مشتری، ممیزی، تأمین‌کننده)، تحلیل ریشه‌ای، اقدامات باز/معوق/بسته، "
     "نرخ تکرار، اثربخشی اقدامات بسته‌شده، عدم انطباق‌های ایمنی و محیط‌زیستی.",
     "9001: 9.3.2 d) 1  ·  14001: 9.3.2 d) 1  ·  45001: 9.3.2 d) 1"),
    ("IN-13", "IMS", "نتایج ممیزی‌ها",
     "ممیزی داخلی (برنامه، پوشش، یافته‌ها، درجه‌بندی)، ممیزی مشتری (VDA6.3 و CSR)، ممیزی گواهی‌دهنده و ممیزی‌های "
     "محیط‌زیستی و ایمنی؛ یافته‌های باز و مهلت رفع؛ تحلیل تکرار یافته‌ها.",
     "9001: 9.3.2 d) 3  ·  14001: 9.3.2 d) 4  ·  45001: 9.3.2 d) 4"),
    ("IN-14", "IMS", "عملکرد تأمین‌کنندگان و پیمانکاران",
     "ارزیابی و رتبه‌بندی تأمین‌کننده (کیفیت، تحویل، قیمت، پاسخ‌گویی)، عدم انطباق ورودی و 8D تأمین‌کننده، "
     "تأمین‌کننده انحصاری و ریسک قطع، عملکرد محیط‌زیستی و ایمنی پیمانکاران (مجوز، آموزش، حادثه)، برنامه توسعه و جایگزینی.",
     "9001: 9.3.2 d) 7  ·  14001 و 45001: 9.3.2 d) (عملکرد ارائه‌دهندگان بیرونی – 8.1)"),
    ("IN-15", "QMS", "هزینه کیفیت ضعیف (COPQ)",
     "هزینه عدم انطباق داخلی (ضایعات، دوباره‌کاری، توقف) و خارجی (برگشت، گارانتی، جریمه، حمل اضطراری)، "
     "نسبت به فروش، روند و هدف، تفکیک محصول و فرآیند.",
     "IATF: 9.3.2.1 a)"),
    ("IN-16", "QMS", "عملکرد گارانتی و برگشتی میدانی",
     "نرخ گارانتی (PPM یا درصد)، هزینه گارانتی، موارد NTF، برگشتی‌های میدانی، روند و مقایسه با هدف و مشتریان.",
     "IATF: 9.3.2.1 h)"),
    ("IN-17", "QMS", "خرابی‌های بالقوه میدانی (خروجی تحلیل ریسک)",
     "فهرست حالت‌های خرابی با شدت بالا و AP بالا در DFMEA/PFMEA، اقدامات پیشگیرانه و وضعیت اجرا، "
     "موارد مرتبط با ایمنی محصول و ویژگی‌های ویژه.",
     "IATF: 9.3.2.1 j)"),
    ("IN-18", "QMS", "خرابی‌های واقعی میدانی و اثر بر ایمنی و محیط‌زیست",
     "خرابی‌ها و حوادث واقعی، فراخوان و اطلاع‌رسانی ایمنی، اثر بر ایمنی سرنشین و محیط‌زیست، اقدامات مهار و ریشه‌ای، "
     "بازخورد مشتری و مراجع قانونی.",
     "IATF: 9.3.2.1 k)"),
    ("IN-19", "QMS", "امکان‌سنجی تولید برای تغییرات و محصول جدید",
     "ارزیابی امکان‌سنجی محصول، قالب یا فرآیند جدید و تغییر فرآیند موجود (ظرفیت، توان تجهیزات، مهارت، زمان‌بندی)، "
     "وضعیت APQP و PPAP، ریسک راه‌اندازی، منابع موردنیاز.",
     "IATF: 9.3.2.1 e) (و 7.1.3.1)"),
    ("IN-20", "QMS", "عملکرد تعمیر و نگهداری نسبت به اهداف",
     "MTBF، MTTR، توقف برنامه‌ریزی‌شده و نشده، نسبت نگهداری پیشگیرانه انجام‌شده، وضعیت تجهیزات بحرانی، "
     "اثر توقف بر تحویل و کیفیت.",
     "IATF: 9.3.2.1 g)"),
    ("IN-21", "IMS", "کفایت منابع",
     "نیروی انسانی (تعداد، مهارت، نرخ ترک، شیفت)، آموزش و صلاحیت، بودجه و نقدینگی، تجهیزات و ابزار، "
     "زیرساخت (ساختمان، تأسیسات، انبار)، منابع پایش و آزمایش، منابع اضطراری و تجهیزات ایمنی، فناوری اطلاعات.",
     "9001: 9.3.2 f)  ·  14001: 9.3.2 e)  ·  45001: 9.3.2 e)"),
    ("IN-22", "IMS", "ریسک‌ها و اثربخشی اقدامات ریسک",
     "فهرست ریسک‌های کلیدی کیفیت، محیط‌زیست، ایمنی و کسب‌وکار؛ سطح ریسک قبل و بعد از اقدام؛ اثربخشی اقدامات با شاهد؛ "
     "ریسک‌های جدید و ریسک باقی‌مانده؛ ریسک اختلال در تأمین و تولید.",
     "9001: 9.3.2 g)  ·  14001: 9.3.2 b) 4  ·  45001: 9.3.2 b) 3 و d) 6"),
    ("IN-23", "IMS", "فرصت‌ها و اثربخشی اقدامات فرصت",
     "فرصت‌های شناسایی‌شده (محصول جدید، بازار جدید، فناوری، بهبود فرآیند، صرفه‌جویی انرژی و مواد)، وضعیت اجرا، "
     "نتیجه و شاخص سنجش، فرصت‌های رهاشده و دلیل.",
     "9001: 9.3.2 h)  ·  14001: 9.3.2 b) 4  ·  45001: 9.3.2 d) 6"),
    ("IN-24", "IMS", "فرصت‌های بهبود",
     "پیشنهادهای کارکنان، نتایج پروژه‌های بهبود و Lean/Kaizen، ایده‌های بهبود فرآیند و محصول، بهبود بهره‌وری انرژی و "
     "کاهش پسماند، بهبود ایمنی؛ فهرست اولویت‌بندی‌شده برای تصمیم‌گیری.",
     "9001: 9.3.2 e)  ·  14001: 9.3.2 g)  ·  45001: 9.3.2 g)"),
    ("IN-25", "EMS", "جنبه‌های محیط‌زیستی مهم و عملکرد محیط‌زیست",
     "فهرست جنبه‌های مهم و تغییرات آن‌ها (انتشار VOC رزین و ژل‌کوت، گردوغبار الیاف، پساب شست‌وشو، پسماند رزین و الیاف و برش، "
     "مصرف انرژی و آب، پسماند خطرناک)، نتایج پایش و حدود مجاز، روند شدت انرژی، شدت پسماند و کربن، اقدامات کاهشی.",
     "14001: 9.3.2 b) 3 و d) 2"),
    ("IN-26", "IMS", "میزان تحقق خط‌مشی و اهداف محیط‌زیست و ایمنی",
     "اهداف محیط‌زیستی و ایمنی با هدف کمی، مقدار واقعی، روند، درصد تحقق، علت انحراف و اقدام؛ "
     "میزان تحقق تعهدات خط‌مشی.",
     "14001: 9.3.2 c)  ·  45001: 9.3.2 c)"),
    ("IN-27", "OHS", "عملکرد ایمنی و بهداشت شغلی",
     "حوادث و شبه‌حوادث، شاخص فراوانی و شدت، روزهای از کارافتادگی، بیماری شغلی و مواجهه (صدا، گردوغبار، بخارات)، "
     "بازدید ایمنی، آموزش، تجهیزات حفاظت فردی، مجوز کار گرم و ارتفاع، مانور و آمادگی اضطراری، شاخص پیشرو و پسرو.",
     "45001: 9.3.2 d) 1"),
    ("IN-28", "OHS", "مشارکت و مشاوره کارکنان",
     "صورت‌جلسات کمیته ایمنی و بهداشت، تعداد و موضوع مشارکت کارکنان، گزارش‌های نزدیک به حادثه ارسالی کارکنان، "
     "موانع مشارکت و اقدامات رفع آن‌ها، بازخورد نمایندگان کارگران.",
     "45001: 9.3.2 d) 5 (و 5.4)"),
    ("IN-29", "IMS", "ارتباطات ذی‌نفعان و شکایات محیط‌زیستی و ایمنی",
     "شکایات محلی (بو، صدا، گردوغبار، پساب)، مکاتبات با اداره محیط‌زیست و بازرسی کار، درخواست کارکنان و پیمانکاران، "
     "اطلاع‌رسانی‌های انجام‌شده، ارتباط با آتش‌نشانی و اورژانس.",
     "14001: 9.3.2 f)  ·  45001: 9.3.2 f)"),
]

FIRST2 = 5
for i, (code, system, subject, info, clause) in enumerate(INPUTS):
    r = FIRST2 + i
    c = ws.cell(row=r, column=1, value=i + 1)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(GREY)

    c = ws.cell(row=r, column=2, value=code)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(GREY)

    c = ws.cell(row=r, column=3)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")

    c = ws.cell(row=r, column=4, value=system)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(GREY)

    c = ws.cell(row=r, column=5, value=subject)
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align("right", "top")
    c.border = BORDER
    c.fill = fill(GREY)

    c = ws.cell(row=r, column=6, value=info)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("right", "top")
    c.border = BORDER

    c = ws.cell(row=r, column=7)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")

    c = ws.cell(row=r, column=8)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.font = Font(name=FONT_NAME, size=9)
    c.alignment = align("right", "top")

    c = ws.cell(row=r, column=9)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")

    c = ws.cell(row=r, column=10)
    c.fill = fill(YELLOW)
    c.border = BORDER
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("right", "top")

    c = ws.cell(row=r, column=11, value=clause)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("right", "top")
    c.border = BORDER
    c.fill = fill(GREY)

    c = ws.cell(row=r, column=12,
                value=f'=IF($B{r}="","",COUNTIF(\'{S3}\'!$F$5:$F$44,"*"&$B{r}&"*"))')
    c.font = Font(name=FONT_NAME, size=9, bold=True)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(BLUE_CALC)

LAST2 = FIRST2 + len(INPUTS) - 1

dv_system = DataValidation(type="list", formula1='"QMS,EMS,OHS,IMS"', allow_blank=True)
dv_status = DataValidation(type="list", formula1='"ارائه شد,ناقص,ارائه نشد,مصداق ندارد"', allow_blank=True)
dv_yesno = DataValidation(type="list", formula1='"بله,خیر"', allow_blank=True)
for d in (dv_system, dv_status, dv_yesno):
    ws.add_data_validation(d)
dv_system.add(f"D{FIRST2}:D{LAST2}")
dv_status.add(f"G{FIRST2}:G{LAST2}")
dv_yesno.add(f"I{FIRST2}:I{LAST2}")

ws.conditional_formatting.add(
    f"G{FIRST2}:G{LAST2}",
    FormulaRule(formula=[f'ISNUMBER(SEARCH("ارائه شد",$G{FIRST2}))'], fill=fill(GREEN)))
ws.conditional_formatting.add(
    f"G{FIRST2}:G{LAST2}",
    FormulaRule(formula=[f'OR(ISNUMBER(SEARCH("ناقص",$G{FIRST2})),ISNUMBER(SEARCH("ارائه نشد",$G{FIRST2})))'],
                fill=fill(RED)))
ws.conditional_formatting.add(
    f"I{FIRST2}:I{LAST2}",
    FormulaRule(formula=[f'$I{FIRST2}="بله"'], fill=fill(AMBER)))

ws.auto_filter.ref = f"A{HROW2}:L{LAST2}"
ws.freeze_panes = f"F{FIRST2}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.print_title_rows = f"{HROW2}:{HROW2}"
ws.print_area = f"A1:L{LAST2}"
ws.oddFooter.center.text = "ورودی‌های بازنگری مدیریت – صفحه &P از &N"
ws.oddFooter.center.size = 8

# ----------------------------------------------------------------------------
# برگه ۳ – خروجی‌ها و مصوبات
# ----------------------------------------------------------------------------
ws = ws3
widths3 = {"A": 4, "B": 9, "C": 11, "D": 27, "E": 8, "F": 10, "G": 30, "H": 36, "I": 15,
           "J": 22, "K": 11, "L": 13, "M": 22, "N": 13, "O": 20, "P": 26}
for col, w in widths3.items():
    ws.column_dimensions[col].width = w

title_cell(ws, "A1:P1", "خروجی‌ها و مصوبات بازنگری مدیریت", size=13, height=28)
title_cell(
    ws, "A2:P2",
    "9 خروجی الزامی استانداردها از پیش تعریف شده است (OUT-01 تا OUT-09). برای هر موضوع می‌توانید چند مصوبه/اقدام با کد جدید "
    "اضافه کنید. ستون «ورودی مرتبط» پیوند ردیابی با برگه ۲ را برقرار می‌کند.",
    size=8, color=SUBHEAD, height=26,
)
title_cell(
    ws, "A3:P3",
    "اگر پس از بررسی، اقدامی لازم نیست: وضعیت «نیاز به اقدام نیست» را انتخاب و دلیل را در ستون «توجیه» ثبت کنید. "
    "برای تعویق نیز دلیل و تصمیم را در همان ستون بنویسید.",
    size=8, color=GREY, fg="404040", height=20,
)

H3 = [
    ("ردیف", "شماره سطر"),
    ("کد مصوبه", "شناسه ثابت ردیابی"),
    ("دوره", "کد دوره بازنگری"),
    ("موضوع خروجی (الزام استاندارد)", "9 موضوع الزامی + موارد افزوده"),
    ("سیستم", "QMS / EMS / OHS / IMS"),
    ("ورودی مرتبط", "کد IN از برگه ۲"),
    ("نتیجه بررسی مدیریت", "جمع‌بندی وضعیت موضوع"),
    ("تصمیم یا مصوبه", "تصمیم مشخص و سنجش‌پذیر"),
    ("مسئول اجرا", "نام و سمت"),
    ("منابع موردنیاز", "نوع و مقدار: نفر، بودجه، تجهیز"),
    ("مهلت اجرا", "تاریخ میلادی"),
    ("وضعیت اقدام", "باز/در حال اجرا/انجام‌شده/معوق/نیاز به اقدام نیست"),
    ("نتیجه اجرا", "شاهد و خروجی واقعی"),
    ("ارزیابی اثربخشی", "اثربخش / نسبتاً / غیراثربخش"),
    ("توجیه عدم اقدام یا تعویق", "الزامی برای عدم اقدام/تعویق"),
    ("بند استاندارد", "9001 / 14001 / 45001 / IATF"),
]
HROW3 = 4
for j, (h, sub) in enumerate(H3, start=1):
    c = ws.cell(row=HROW3, column=j, value=h + "\n" + f"({sub})")
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center", "center")
    c.border = BORDER
ws.row_dimensions[HROW3].height = 42

OUTPUTS = [
    ("OUT-01", "IMS", "نتیجه‌گیری درباره تناسب، کفایت و اثربخشی سیستم‌های مدیریت",
     "14001: 9.3.3 a)  ·  45001: 9.3.3  ·  9001: 9.3.1 و 9.3.3"),
    ("OUT-02", "IMS", "تصمیمات مربوط به فرصت‌های بهبود (فرآیند، محصول، سیستم)",
     "9001: 9.3.3  ·  14001: 9.3.3 b)  ·  45001: 9.3.3"),
    ("OUT-03", "IMS", "تغییرات موردنیاز در سیستم مدیریت (خط‌مشی، اهداف، فرآیندها، مستندات)",
     "9001: 9.3.3  ·  14001: 9.3.3 c)  ·  45001: 9.3.3  ·  IATF: 9.3.3"),
    ("OUT-04", "IMS", "نیازهای منابع (انسانی، مالی، تجهیزات، زیرساخت)",
     "9001: 9.3.3  ·  14001: 9.3.3 c)  ·  45001: 9.3.3"),
    ("OUT-05", "IMS", "اقدامات لازم برای اهداف محقق‌نشده و مشکلات شناسایی‌شده",
     "14001: 9.3.3 d)  ·  45001: 9.3.3 (اقدامات در صورت نیاز)  ·  9001: 9.3.3"),
    ("OUT-06", "IMS", "فرصت‌های بهبود یکپارچگی سیستم با سایر فرآیندهای کسب‌وکار",
     "14001: 9.3.3 e)  ·  45001: 9.3.3"),
    ("OUT-07", "IMS", "پیامدهای مصوبات بر جهت‌گیری استراتژیک سازمان",
     "14001: 9.3.3 f)  ·  45001: 9.3.3"),
    ("OUT-08", "QMS", "برنامه اقدام مستند برای عدم تحقق اهداف عملکرد مشتری",
     "IATF: 9.3.3.1"),
    ("OUT-09", "OHS", "ابلاغ مصوبات مرتبط به کارکنان و نمایندگان آن‌ها",
     "45001: 9.3.3 (ابلاغ به کارکنان)"),
]

FIRST3 = 5
LAST3 = 44
dv_inputs = DataValidation(type="list", formula1="LIST_INPUTS", allow_blank=True)

for i in range(FIRST3, LAST3 + 1):
    idx = i - FIRST3
    ws.cell(row=i, column=1, value=idx + 1).font = Font(name=FONT_NAME, size=8)
    ws.cell(row=i, column=1).alignment = align("center", "top")
    ws.cell(row=i, column=1).border = BORDER
    ws.cell(row=i, column=1).fill = fill(GREY)

    if idx < len(OUTPUTS):
        code, system, subject, clause = OUTPUTS[idx]
        c = ws.cell(row=i, column=2, value=code)
        c.font = Font(name=FONT_NAME, size=9, bold=True)
        c.fill = fill(GREY)
        c.alignment = align("center", "top")
        c.border = BORDER
        c = ws.cell(row=i, column=4, value=subject)
        c.font = Font(name=FONT_NAME, size=9, bold=True)
        c.fill = fill(GREY)
        c.alignment = align("right", "top")
        c.border = BORDER
        c = ws.cell(row=i, column=5, value=system)
        c.font = Font(name=FONT_NAME, size=8)
        c.fill = fill(GREY)
        c.alignment = align("center", "top")
        c.border = BORDER
        c = ws.cell(row=i, column=16, value=clause)
        c.font = Font(name=FONT_NAME, size=8)
        c.fill = fill(GREY)
        c.alignment = align("right", "top")
        c.border = BORDER
    else:
        c = ws.cell(row=i, column=2)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=9)
        c.alignment = align("center", "top")
        for col in (4, 16):
            c = ws.cell(row=i, column=col)
            c.fill = fill(YELLOW)
            c.border = BORDER
            c.font = Font(name=FONT_NAME, size=8)
            c.alignment = align("right", "top")
        c = ws.cell(row=i, column=5)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=8)
        c.alignment = align("center", "top")

    for col in (3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15):
        c = ws.cell(row=i, column=col)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=8 if col in (3, 6, 11, 12, 14) else 9)
        c.alignment = align("center", "top") if col in (3, 6, 11, 12, 14) else align("right", "top")
        if col == 11:
            c.number_format = "yyyy-mm-dd"

dv_out_system = DataValidation(type="list", formula1='"QMS,EMS,OHS,IMS"', allow_blank=True)
dv_out_status = DataValidation(
    type="list", formula1='"باز,در حال اجرا,انجام‌شده,معوق,نیاز به اقدام نیست"', allow_blank=True)
dv_effect = DataValidation(
    type="list", formula1='"اثربخش,نسبتاً اثربخش,غیراثربخش,در انتظار ارزیابی,ارزیابی نشده"', allow_blank=True)
for d in (dv_inputs, dv_out_system, dv_out_status, dv_effect):
    ws.add_data_validation(d)
dv_inputs.add(f"F{FIRST3}:F{LAST3}")
dv_out_system.add(f"E{FIRST3}:E{LAST3}")
dv_out_status.add(f"L{FIRST3}:L{LAST3}")
dv_effect.add(f"N{FIRST3}:N{LAST3}")

# رنگ وضعیت‌ها
ws.conditional_formatting.add(
    f"L{FIRST3}:L{LAST3}", FormulaRule(formula=[f'$L{FIRST3}="انجام‌شده"'], fill=fill(GREEN)))
ws.conditional_formatting.add(
    f"L{FIRST3}:L{LAST3}", FormulaRule(formula=[f'$L{FIRST3}="معوق"'], fill=fill(RED)))
ws.conditional_formatting.add(
    f"L{FIRST3}:L{LAST3}", FormulaRule(formula=[f'$L{FIRST3}="باز"'], fill=fill(AMBER)))
ws.conditional_formatting.add(
    f"N{FIRST3}:N{LAST3}", FormulaRule(formula=[f'ISNUMBER(SEARCH("اثربخش",$N{FIRST3}))'], fill=fill(GREEN)))
ws.conditional_formatting.add(
    f"N{FIRST3}:N{LAST3}", FormulaRule(formula=[f'$N{FIRST3}="غیراثربخش"'], fill=fill(RED)))
ws.conditional_formatting.add(
    f"N{FIRST3}:N{LAST3}", FormulaRule(formula=[f'$N{FIRST3}="در انتظار ارزیابی"'], fill=fill(AMBER)))

ws.auto_filter.ref = f"A{HROW3}:P{LAST3}"
ws.freeze_panes = f"E{FIRST3}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.print_title_rows = f"{HROW3}:{HROW3}"
ws.print_area = f"A1:P{LAST3}"
ws.oddFooter.center.text = "خروجی‌ها و مصوبات بازنگری مدیریت – صفحه &P از &N"
ws.oddFooter.center.size = 8

# ----------------------------------------------------------------------------
# برگه ۴ – پایش اجرای مصوبات و اثربخشی
# ----------------------------------------------------------------------------
ws = ws4
widths4 = {"A": 4, "B": 10, "C": 27, "D": 15, "E": 11, "F": 11, "G": 8, "H": 14,
           "I": 28, "J": 26, "K": 11, "L": 13, "M": 14, "N": 11, "O": 24}
for col, w in widths4.items():
    ws.column_dimensions[col].width = w

title_cell(ws, "A1:O1", "پایش اجرای مصوبات و ارزیابی اثربخشی", size=13, height=28)
title_cell(
    ws, "A2:O2",
    "هر ردیف یک نوبت پایش از یک مصوبه است. برای پایش دوره‌ای، ردیف جدید با همان کد مصوبه ثبت کنید. "
    "ستون‌های موضوع، مسئول و مهلت به‌صورت خودکار از برگه ۳ خوانده می‌شوند؛ نتیجه نهایی اثربخشی باید با شاهد و معیار ثبت شود.",
    size=8, color=SUBHEAD, height=26,
)
title_cell(
    ws, "A3:O3",
    "«اثربخش» بدون معیار سنجش، تاریخ ارزیابی و نام ارزیاب معتبر نیست. به‌روزرسانی ستون «وضعیت اقدام» برگه ۳ بر پایه همین برگه انجام می‌شود.",
    size=8, color=GREY, fg="404040", height=20,
)

H4 = [
    ("ردیف", "شماره سطر"),
    ("کد مصوبه", "انتخاب از برگه ۳"),
    ("موضوع مصوبه", "فرمول"),
    ("مسئول اجرا", "فرمول"),
    ("مهلت اجرا", "فرمول"),
    ("تاریخ پایش", "میلادی"),
    ("درصد پیشرفت", "0 تا 100"),
    ("وضعیت در تاریخ پایش", "باز/در حال اجرا/انجام‌شده/معوق/متوقف"),
    ("شاهد پیشرفت یا اجرا", "مدرک عینی"),
    ("معیار ارزیابی اثربخشی", "کمی یا کیفی"),
    ("تاریخ ارزیابی اثربخشی", "میلادی"),
    ("ارزیاب", "نام و سمت"),
    ("نتیجه ارزیابی اثربخشی", "اثربخش/نسبتاً/غیراثربخش/در انتظار"),
    ("روز تا مهلت", "منفی = تأخیر"),
    ("اقدام بعدی و مسئول", "تصمیم پیگیری"),
]
HROW4 = 4
for j, (h, sub) in enumerate(H4, start=1):
    c = ws.cell(row=HROW4, column=j, value=h + "\n" + f"({sub})")
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center", "center")
    c.border = BORDER
ws.row_dimensions[HROW4].height = 40

FIRST4 = 5
LAST4 = 44
dv_out_codes = DataValidation(type="list", formula1="LIST_OUTPUTS", allow_blank=True)
dv_progress = DataValidation(type="whole", operator="between", formula1="0", formula2="100", allow_blank=True)
dv_mon_status = DataValidation(type="list", formula1='"باز,در حال اجرا,انجام‌شده,معوق,متوقف"', allow_blank=True)
dv_mon_effect = DataValidation(type="list", formula1='"اثربخش,نسبتاً اثربخش,غیراثربخش,در انتظار"', allow_blank=True)
for d in (dv_out_codes, dv_progress, dv_mon_status, dv_mon_effect):
    ws.add_data_validation(d)
dv_out_codes.add(f"B{FIRST4}:B{LAST4}")
dv_progress.add(f"G{FIRST4}:G{LAST4}")
dv_mon_status.add(f"H{FIRST4}:H{LAST4}")
dv_mon_effect.add(f"M{FIRST4}:M{LAST4}")

for r in range(FIRST4, LAST4 + 1):
    ws.cell(row=r, column=1, value=r - FIRST4 + 1).font = Font(name=FONT_NAME, size=8)
    ws.cell(row=r, column=1).alignment = align("center", "top")
    ws.cell(row=r, column=1).border = BORDER
    ws.cell(row=r, column=1).fill = fill(GREY)

    for col in (2, 6, 7, 8, 11, 12, 13):
        c = ws.cell(row=r, column=col)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=8)
        c.alignment = align("center", "top")
    for col in (9, 10, 15):
        c = ws.cell(row=r, column=col)
        c.fill = fill(YELLOW)
        c.border = BORDER
        c.font = Font(name=FONT_NAME, size=8)
        c.alignment = align("right", "top")

    ws.cell(row=r, column=6).number_format = "yyyy-mm-dd"
    ws.cell(row=r, column=11).number_format = "yyyy-mm-dd"

    c = ws.cell(row=r, column=3,
                value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},\'{S3}\'!$B${FIRST3}:$P${LAST3},3,FALSE),""))')
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("right", "top")
    c.border = BORDER
    c.fill = fill(BLUE_CALC)

    c = ws.cell(row=r, column=4,
                value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},\'{S3}\'!$B${FIRST3}:$P${LAST3},8,FALSE),""))')
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(BLUE_CALC)

    c = ws.cell(row=r, column=5,
                value=f'=IF($B{r}="","",IFERROR(VLOOKUP($B{r},\'{S3}\'!$B${FIRST3}:$P${LAST3},10,FALSE),""))')
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(BLUE_CALC)
    c.number_format = "yyyy-mm-dd"

    c = ws.cell(row=r, column=14,
                value=f'=IF(OR($B{r}="",$E{r}=""),"",IF($H{r}="انجام‌شده","",$E{r}-TODAY()))')
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.alignment = align("center", "top")
    c.border = BORDER
    c.fill = fill(BLUE_CALC)
    c.number_format = "0"

ws.conditional_formatting.add(
    f"H{FIRST4}:H{LAST4}", FormulaRule(formula=[f'$H{FIRST4}="انجام‌شده"'], fill=fill(GREEN)))
ws.conditional_formatting.add(
    f"H{FIRST4}:H{LAST4}", FormulaRule(formula=[f'$H{FIRST4}="معوق"'], fill=fill(RED)))
ws.conditional_formatting.add(
    f"H{FIRST4}:H{LAST4}", FormulaRule(formula=[f'$H{FIRST4}="در حال اجرا"'], fill=fill(AMBER)))
ws.conditional_formatting.add(
    f"M{FIRST4}:M{LAST4}", FormulaRule(formula=[f'ISNUMBER(SEARCH("اثربخش",$M{FIRST4}))'], fill=fill(GREEN)))
ws.conditional_formatting.add(
    f"M{FIRST4}:M{LAST4}", FormulaRule(formula=[f'$M{FIRST4}="غیراثربخش"'], fill=fill(RED)))
ws.conditional_formatting.add(
    f"N{FIRST4}:N{LAST4}",
    FormulaRule(formula=[f'AND(ISNUMBER($N{FIRST4}),$N{FIRST4}<0)'], fill=fill(RED),
                font=Font(name=FONT_NAME, size=8, bold=True, color="9C0006")))

ws.auto_filter.ref = f"A{HROW4}:O{LAST4}"
ws.freeze_panes = f"C{FIRST4}"
ws.page_setup.orientation = "landscape"
ws.page_setup.paperSize = ws.PAPERSIZE_A3
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.print_title_rows = f"{HROW4}:{HROW4}"
ws.print_area = f"A1:O{LAST4}"
ws.oddFooter.center.text = "پایش اجرا و اثربخشی مصوبات – صفحه &P از &N"
ws.oddFooter.center.size = 8

# ----------------------------------------------------------------------------
# برگه ۵ – صورت‌جلسه قابل چاپ
# ----------------------------------------------------------------------------
ws = ws5
widths5 = {"A": 12, "B": 22, "C": 28, "D": 13, "E": 15, "F": 11, "G": 12}
for col, w in widths5.items():
    ws.column_dimensions[col].width = w

title_cell(ws, "A1:G1", "صورت‌جلسه بازنگری مدیریت – سیستم‌های مدیریت یکپارچه", size=14, height=30)
title_cell(
    ws, "A2:G2",
    "مطابق بند 9.3 استانداردهای ISO 9001:2026 ، ISO 14001:2026 ، ISO 45001:2018 و الزامات تکمیلی IATF 16949:2016",
    size=8, color=SUBHEAD, height=20,
)
title_cell(ws, "A3:G3", "کد فرم: MR-IMS-F01  ·  نسخه 1.0  ·  این برگه از داده‌های برگه‌های ۱ تا ۳ تولید می‌شود",
           size=8, color=GREY, fg="404040", height=16)


def ref(row, col):
    return f"'{S1}'!{col}{row}"


# اطلاعات سربرگ
info_rows = [
    (5, "سازمان / واحد", f'=IF({ref(6,"C")}="","",{ref(6,"C")})&IF({ref(7,"C")}="",""," – "&{ref(7,"C")})',
     "شماره صورت‌جلسه", f'=IF({ref(10,"C")}="","",{ref(10,"C")})'),
    (6, "تاریخ جلسه", f'=IF({ref(12,"C")}="","",TEXT({ref(12,"C")},"yyyy-mm-dd"))',
     "دوره بازنگری", f'=IF({ref(9,"C")}="","",{ref(9,"C")})'),
    (7, "رئیس جلسه", f'=IF({ref(15,"C")}="","",{ref(15,"C")})',
     "دبیر جلسه", f'=IF({ref(16,"C")}="","",{ref(16,"C")})'),
    (8, "بازنگری قبلی / بعدی",
     f'=IF({ref(18,"C")}="","",TEXT({ref(18,"C")},"yyyy-mm-dd"))&" / "&IF({ref(19,"C")}="","",TEXT({ref(19,"C")},"yyyy-mm-dd"))',
     "بازه داده‌ها", f'=IF({ref(17,"C")}="","",{ref(17,"C")})'),
]
for r, l1, v1, l2, v2 in info_rows:
    c = ws.cell(row=r, column=1, value=l1)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.fill = fill(GREY)
    c.alignment = align()
    c.border = BORDER
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    c = ws.cell(row=r, column=2, value=v1)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.alignment = align()
    c.border = BORDER
    ws.cell(row=r, column=3).border = BORDER
    c = ws.cell(row=r, column=4, value=l2)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.fill = fill(GREY)
    c.alignment = align()
    c.border = BORDER
    c = ws.cell(row=r, column=5, value=v2)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align()
    c.border = BORDER
    ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=7)
    ws.cell(row=r, column=6).border = BORDER
    ws.cell(row=r, column=7).border = BORDER
    ws.row_dimensions[r].height = 20

# نوار خلاصه وضعیت مصوبات
m1, m2, m3, m4 = (metric_rows["مصوبات ثبت‌شده"], metric_rows["مصوبات انجام‌شده"],
                  metric_rows["مصوبات باز یا در حال اجرا"], metric_rows["مصوبات معوق"])
ws.merge_cells("A9:G9")
c = ws["A9"]
pref = "'" + S1 + "'!C"
c.value = ('="خلاصه وضعیت مصوبات — ثبت‌شده: "&' + pref + str(m1)
           + '&"  ·  انجام‌شده: "&' + pref + str(m2)
           + '&"  ·  باز یا در حال اجرا: "&' + pref + str(m3)
           + '&"  ·  معوق: "&' + pref + str(m4))
c.font = Font(name=FONT_NAME, size=8, bold=True, color=NAVY)
c.fill = fill(SECTION)
c.alignment = align("center")
style_range(ws, "A9:G9", border=True, fill_color=SECTION)
ws.row_dimensions[9].height = 20

# بخش ۱ – خلاصه ورودی‌ها
r = 10
section(ws, f"A{r}:G{r}", "۱) خلاصه ورودی‌های بررسی‌شده (شرح کامل در برگه ۲)")
r += 1
for j, (h) in enumerate(["کد", "موضوع ورودی", "نتیجه بررسی و تحلیل مدیریت", "", "وضعیت ارائه", "نیاز به مصوبه", "بند"]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    c.border = BORDER
ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
ws.row_dimensions[r].height = 22

MIRROR2 = 12  # تعداد ردیف ورودی در صورت‌جلسه
for i in range(MIRROR2):
    src = FIRST2 + i
    rr = r + 1 + i
    c = ws.cell(row=rr, column=1, value=f"=IF('{S2}'!B{src}=\"\",\"\",'{S2}'!B{src})")
    c.font = Font(name=FONT_NAME, size=7, bold=True)
    c.alignment = align("center", "top")
    c.border = BORDER
    c = ws.cell(row=rr, column=2, value=f"=IF('{S2}'!E{src}=\"\",\"\",'{S2}'!E{src})")
    c.font = Font(name=FONT_NAME, size=7)
    c.alignment = align("right", "top")
    c.border = BORDER
    ws.merge_cells(start_row=rr, start_column=3, end_row=rr, end_column=4)
    c = ws.cell(row=rr, column=3, value=f"=IF('{S2}'!H{src}=\"\",\"\",'{S2}'!H{src})")
    c.font = Font(name=FONT_NAME, size=7)
    c.alignment = align("right", "top")
    c.border = BORDER
    ws.cell(row=rr, column=4).border = BORDER
    c = ws.cell(row=rr, column=5, value=f"=IF('{S2}'!G{src}=\"\",\"\",'{S2}'!G{src})")
    c.font = Font(name=FONT_NAME, size=7)
    c.alignment = align("center", "top")
    c.border = BORDER
    c = ws.cell(row=rr, column=6, value=f"=IF('{S2}'!I{src}=\"\",\"\",'{S2}'!I{src})")
    c.font = Font(name=FONT_NAME, size=7)
    c.alignment = align("center", "top")
    c.border = BORDER
    c = ws.cell(row=rr, column=7, value=f"=IF('{S2}'!K{src}=\"\",\"\",'{S2}'!K{src})")
    c.font = Font(name=FONT_NAME, size=6)
    c.alignment = align("right", "top")
    c.border = BORDER
    ws.row_dimensions[rr].height = 30

r = r + 1 + MIRROR2
merge_text(ws, f"A{r}:G{r}",
           "توجه: این جدول 12 ورودی اول را نشان می‌دهد؛ جدول کامل 29 ورودی در برگه ۲ است و بخشی از سند بازنگری محسوب می‌شود.",
           style="note", height=16)

# بخش ۲ – نتیجه‌گیری
r += 1
section(ws, f"A{r}:G{r}", "۲) نتیجه‌گیری مدیریت (از برگه ۳)")
conclusions = [
    ("تناسب، کفایت و اثربخشی", 5, "H"),
    ("پیامد بر جهت‌گیری استراتژیک", 11, "H"),
    ("ابلاغ به کارکنان و نمایندگان", 13, "H"),
]
for label, row3, col3 in conclusions:
    r += 1
    c = ws.cell(row=r, column=1, value=label)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.fill = fill(GREY)
    c.alignment = align()
    c.border = BORDER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    ws.cell(row=r, column=2).border = BORDER
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    c = ws.cell(row=r, column=3,
                value=f"=IF('{S3}'!{col3}{row3}=\"\",\"(در برگه ۳ ثبت نشده)\",'{S3}'!{col3}{row3})")
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("right", "top")
    c.border = BORDER
    for cc in (4, 5, 6, 7):
        ws.cell(row=r, column=cc).border = BORDER
    ws.row_dimensions[r].height = 30

# بخش ۳ – مصوبات
r += 2
section(ws, f"A{r}:G{r}", "۳) مصوبات و تصمیمات (از برگه ۳ – حداکثر 15 مصوبه اول)")
r += 1
for j, h in enumerate(["کد مصوبه", "موضوع خروجی", "تصمیم یا مصوبه", "مسئول اجرا", "منابع موردنیاز", "مهلت", "وضعیت"]):
    c = ws.cell(row=r, column=1 + j, value=h)
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    c.border = BORDER
ws.row_dimensions[r].height = 22

MIRROR3 = 15
for i in range(MIRROR3):
    src = FIRST3 + i
    rr = r + 1 + i
    for j, (src_col, fname) in enumerate(
        [(2, "B"), (4, "D"), (8, "H"), (9, "I"), (10, "J"), (11, "K"), (12, "L")]):
        c = ws.cell(row=rr, column=1 + j,
                    value=f"=IF('{S3}'!{fname}{src}=\"\",\"\",'{S3}'!{fname}{src})")
        c.font = Font(name=FONT_NAME, size=7, bold=(j == 0))
        c.alignment = align("center" if j in (0, 5, 6) else "right", "top")
        c.border = BORDER
        if fname == "K":
            c.number_format = "yyyy-mm-dd"
    ws.row_dimensions[rr].height = 32
merge_text(ws, f"A{r + 1 + MIRROR3}:G{r + 1 + MIRROR3}",
           "مصوبات بیشتر: ادامه جدول در برگه ۳ (تا ردیف 44). کل برگه ۳ سند رسمی مصوبات است.",
           style="note", height=16)

# بخش ۴ – امضاها
r = r + MIRROR3 + 3
section(ws, f"A{r}:G{r}", "۴) تأیید و امضا")
r += 1
sig_headers = [("نقش و مسئولیت", 1, 2), ("نام و سمت", 3, 3), ("امضا", 4, 5), ("تاریخ", 6, 7)]
for h, c1, c2 in sig_headers:
    if c1 != c2:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    c = ws.cell(row=r, column=c1, value=h)
    c.font = Font(name=FONT_NAME, size=8, bold=True, color="FFFFFF")
    c.fill = fill(HEAD)
    c.alignment = align("center")
    style_range(ws, f"{get_column_letter(c1)}{r}:{get_column_letter(c2)}{r}", border=True, fill_color=HEAD)
ws.row_dimensions[r].height = 20

sign_rows = [
    ("تهیه‌کننده / دبیر جلسه", f"=IF({ref(16,'C')}=\"\",\"\",{ref(16,'C')})"),
    ("بررسی: مدیر تضمین کیفیت / IMS", ""),
    ("تأیید: مدیریت ارشد (مدیرعامل)", f"=IF({ref(15,'C')}=\"\",\"\",{ref(15,'C')})"),
]
for label, formula in sign_rows:
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    c = ws.cell(row=r, column=1, value=label)
    c.font = Font(name=FONT_NAME, size=8, bold=True)
    c.alignment = align()
    style_range(ws, f"A{r}:B{r}", border=True)

    c = ws.cell(row=r, column=3, value=formula)
    c.font = Font(name=FONT_NAME, size=8)
    c.alignment = align("center")
    c.border = BORDER

    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
    style_range(ws, f"D{r}:E{r}", border=True, fill_color=YELLOW)
    style_range(ws, f"F{r}:G{r}", border=True, fill_color=YELLOW)
    ws.cell(row=r, column=6).number_format = "yyyy-mm-dd"
    ws.row_dimensions[r].height = 32

# بخش ۵ – توزیع
r += 2
section(ws, f"A{r}:G{r}", "۵) توزیع و پیوست‌ها")
r += 1
dist = [
    "توزیع: مدیران فرآیندها و مسئولان اجرا · کارکنان و نمایندگان آن‌ها (مصوبات مرتبط ایمنی – 45001 بند 9.3.3) · سیستم کنترل مستندات.",
    "پیوست‌های الزامی: گزارش ورودی‌های برگه ۲ · جدول مصوبات برگه ۳ · سوابق پایش و اثربخشی برگه ۴ · فهرست حاضرین.",
    "نگهداری: نسخه مصوب امضاشده در سیستم کنترل مستندات ثبت و در بازنگری بعدی به‌عنوان ورودی IN-01 بررسی شود.",
]
for text in dist:
    merge_text(ws, f"A{r}:G{r}", text, height=24)
    r += 1


ws5.page_setup.orientation = "landscape"
ws5.page_setup.paperSize = ws5.PAPERSIZE_A4
ws5.page_setup.fitToWidth = 1
ws5.page_setup.fitToHeight = 0
ws5.print_options.horizontalCentered = True
ws5.print_area = f"A1:G{r}"
ws5.oddFooter.center.text = "صورت‌جلسه بازنگری مدیریت – صفحه &P از &N"
ws5.oddFooter.center.size = 8
ws5.page_margins.left = 0.3
ws5.page_margins.right = 0.3
ws5.page_margins.top = 0.4
ws5.page_margins.bottom = 0.4

# فهرست‌های کشویی از طریق نام تعریف‌شده (سازگارتر با نسخه‌های مختلف Excel)
wb.defined_names["LIST_INPUTS"] = DefinedName("LIST_INPUTS", attr_text=f"'{S2}'!$B${FIRST2}:$B${LAST2}")
wb.defined_names["LIST_OUTPUTS"] = DefinedName("LIST_OUTPUTS", attr_text=f"'{S3}'!$B${FIRST3}:$B${LAST3}")

wb.active = 0

wb.properties.title = "بازنگری مدیریت IMS – ISO 9001:2026 / ISO 14001:2026 / ISO 45001:2018 / IATF 16949:2016"
wb.properties.subject = "فرم و صورت‌جلسه بازنگری مدیریت"
wb.properties.creator = "IMS / تضمین کیفیت"
wb.properties.description = "ورودی‌ها، خروجی‌ها، مصوبات، پایش اثربخشی و صورت‌جلسه قابل چاپ"


# --- تنظیم ارتفاع سطرها بر پایه طول متن (پیش از ذخیره) ---
autofit_rows(ws1, range(6, 21), {"B": 9, "E": 8}, min_h=24)
autofit_rows(ws1, range(37, 41), {"B": 9, "C": 9, "D": 9, "E": 8, "F": 8}, min_h=28)
autofit_rows(ws1, range(59, 70), {"B": 9, "C": 9}, min_h=18)
autofit_rows(ws2, range(5, LAST2 + 1), {"E": 9, "F": 8, "H": 9, "J": 8, "K": 8}, min_h=28)
autofit_rows(ws3, range(5, LAST3 + 1), {"D": 9, "G": 9, "H": 9, "I": 8, "J": 8, "M": 8, "O": 8, "P": 8}, min_h=28)
autofit_rows(ws4, range(5, LAST4 + 1), {"I": 8, "J": 8, "O": 8}, min_h=26)
autofit_rows(ws5, range(5, 9), {"A": 8, "B": 8, "D": 8, "E": 8}, min_h=22)
autofit_rows(ws5, range(12, 24), {"A": 7, "B": 7, "C": 7, "E": 7, "F": 7, "G": 6}, min_h=26)
autofit_rows(ws5, range(32, 47), {"A": 7, "B": 7, "C": 7, "D": 7, "E": 7, "G": 7}, min_h=26)
autofit_rows(ws5, range(26, 29), {"A": 8, "C": 8}, min_h=26)

wb.save(OUT)
print("saved:", OUT)

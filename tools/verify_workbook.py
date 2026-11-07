#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
بررسی سلامت ساختاری و محاسباتی فایل بازنگری مدیریت.

اجرا:
    python3 tools/verify_workbook.py IMS_Management_Review_FA_v1.xlsx
    python3 tools/verify_workbook.py IMS_Management_Review_FA_v1_Demo.xlsx --calc
"""

import sys
import zipfile

import openpyxl
from openpyxl.utils import range_boundaries


def check_structure(path):
    errors, warnings = [], []
    z = zipfile.ZipFile(path)
    if z.testzip() is not None:
        errors.append("فایل xlsx آسیب‌دیده است")

    wb = openpyxl.load_workbook(path)
    names = wb.sheetnames
    for ws in wb.worksheets:
        if not ws.sheet_view.rightToLeft:
            errors.append(f"{ws.title}: راست‌به‌چپ نیست")

        boxes = [range_boundaries(str(m)) for m in ws.merged_cells.ranges]
        for i in range(len(boxes)):
            for j in range(i + 1, len(boxes)):
                a, b = boxes[i], boxes[j]
                if not (a[2] < b[0] or b[2] < a[0] or a[3] < b[1] or b[3] < a[1]):
                    errors.append(f"{ws.title}: ادغام هم‌پوشان {a} و {b}")

        for m in ws.merged_cells.ranges:
            if m.min_row == m.max_row and m.min_col == m.max_col:
                warnings.append(f"{ws.title}: ادغام تک‌سلولی {m.coord}")

        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1 and not dv.formula1.startswith('"'):
                ref = dv.formula1.strip("'")
                sheet = ref.split("!")[0]
                if sheet not in names and sheet not in wb.defined_names:
                    errors.append(f"{ws.title}: فهرست کشویی به کاربرگ ناموجود {sheet} اشاره دارد")

        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    for token in ("#REF", "#NAME", "#VALUE"):
                        if token in c.value:
                            errors.append(f"{ws.title}!{c.coordinate}: خطای {token} در فرمول")

        if not ws.print_area:
            warnings.append(f"{ws.title}: ناحیه چاپ تعریف نشده")

    return names, errors, warnings


def check_calc(path):
    try:
        from pycel import ExcelCompiler
    except ImportError:
        return ["pycel نصب نیست؛ بررسی محاسبات انجام نشد"], []

    errors = []
    xl = ExcelCompiler(filename=path)
    wb = openpyxl.load_workbook(path)
    count = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    count += 1
                    try:
                        xl.evaluate(f"'{ws.title}'!{c.coordinate}")
                    except Exception as exc:  # noqa: BLE001
                        errors.append(f"{ws.title}!{c.coordinate}: {type(exc).__name__} – {str(exc)[:120]}")
    return errors, [f"{count} فرمول ارزیابی شد"]


def main():
    paths = [a for a in sys.argv[1:] if not a.startswith("--")]
    calc = "--calc" in sys.argv
    if not paths:
        paths = ["IMS_Management_Review_FA_v1.xlsx", "IMS_Management_Review_FA_v1_Demo.xlsx"]

    failed = False
    for path in paths:
        print("=" * 78)
        print("بررسی:", path)
        names, errors, warnings = check_structure(path)
        print("  برگه‌ها:", " · ".join(names))
        if calc:
            cerr, info = check_calc(path)
            errors += cerr
            for i in info:
                print("  محاسبه:", i)
        for w in warnings:
            print("  هشدار:", w)
        for e in errors:
            print("  خطا:", e)
        print("  نتیجه:", "سالم" if not errors else f"{len(errors)} خطا")
        failed = failed or bool(errors)

    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()

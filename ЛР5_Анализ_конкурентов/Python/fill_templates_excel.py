# -*- coding: utf-8 -*-
"""
Проверка без VBA: заполняет КОПИИ трёх шаблонов теми же TSV-файлами (../VBA/data), что и VBA-модули,
через Excel (COM, pywin32), пересчитывает и сверяет результаты формул шаблона с lr5_calc.py.
Логика LoadBlock повторяет VBA 1:1. Сохраняет ../Excel/ЛР5_*_NotaCode.xlsx.
Запуск (Windows + Excel):  python -B fill_templates_excel.py
"""
import os
import shutil
import sys

import win32com.client as w32

from lr5_calc import calc_a, calc_b, calc_c

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.normpath(os.path.join(HERE, ".."))
DATA = os.path.join(ROOT, "VBA", "data")
OUT = os.path.join(ROOT, "Excel")
TPL = os.path.join(ROOT, "Материалы")


def is_num(t):
    return 0 < len(t) <= 15 and t.count(".") <= 1 and t.replace(".", "").isdigit() and not t.startswith(".") and not t.endswith(".")


def sheet(wb, p):
    if p.startswith("+"):
        p = p[1:]
        for ws in wb.Worksheets:
            if ws.Name == p:
                return ws
        ws = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = p
        return ws
    for ws in wb.Worksheets:
        if ws.Name.startswith(p):
            return ws
    raise KeyError(p)


def load(wb, fn):
    lines = open(os.path.join(DATA, fn), encoding="utf-8").read().split("\n")
    _, pref, anchor = lines[0].split("\t")
    ws = sheet(wb, pref)
    r0, c0 = ws.Range(anchor).Row, ws.Range(anchor).Column
    for i, line in enumerate(lines[1:], 1):
        if not line:
            continue
        for j, t in enumerate(line.split("\t")):
            if not t:
                continue
            t = t.replace("\\n", "\n")
            c = ws.Cells(r0 + i - 1, c0 + j)
            if t.startswith("="):
                c.Formula = t
            elif is_num(t):
                c.Value = float(t)
            else:
                c.NumberFormat = "@"
                c.Value = t


def run(xl, tpl, out, files, checks):
    dst = os.path.join(OUT, out)
    shutil.copy(os.path.join(TPL, tpl), dst)
    wb = xl.Workbooks.Open(dst)
    xl.Calculation = -4135
    if out.startswith("ЛР5_1"):   # как FixXlookup в VBA
        import re
        rx = re.compile(r"(_xlfn\.)?XLOOKUP\(([^,()]+),([^,()]+),([^,()]+)\)")
        for ws in wb.Worksheets:
            rng = ws.UsedRange
            fs = rng.Formula
            if not isinstance(fs, tuple) or not any("XLOOKUP" in str(v).upper() for row in fs for v in row):
                continue
            rng.Formula = tuple(tuple(rx.sub(lambda m: f"INDEX({m.group(4)},MATCH({m.group(2)},{m.group(3)},0))", v)
                                      if isinstance(v, str) and "XLOOKUP" in v.upper() else v for v in row) for row in fs)
    if out.startswith("ЛР5_3"):
        sheet(wb, "09_").Range("A2:Q2").UnMerge()   # как в VBA: строка весов была объединена
        sheet(wb, "09_").Range("A2").Value = "Вес"
    for f in files:
        load(wb, f)
    xl.Calculation = -4105
    xl.CalculateFull()
    bad = checks(wb)
    wb.Save()
    wb.Close()
    print(out, "— расхождений:", bad)
    return bad


def chk_a(wb):
    ws = sheet(wb, "04_")
    bad = 0
    for i, r in enumerate(calc_a()):
        n, fin, zone = ws.Cells(4 + i, 14).Value, ws.Cells(4 + i, 18).Value, sheet(wb, "07_").Cells(4 + i, 7).Value
        if abs(n - r["N"]) > 1e-9 or fin != r["final"] or zone != r["zone"]:
            bad += 1
            print("  A", r["id"], n, fin, zone)
    s6 = sheet(wb, "06_")
    top = [s6.Cells(19 + k, 2).Value for k in range(10)]
    print("  топ-10:", top)
    bad += len(top) - len(set(top))
    return bad


def chk_b(wb):
    B = calc_b()
    ws = sheet(wb, "06_")
    bad = 0
    for i, c in enumerate(B["comp"]):
        v = ws.Cells(90, 5 + i).Value
        if abs(v - c["total"]) > 1e-9 or ws.Cells(91, 5 + i).Value != c["interp"]:
            bad += 1
            print("  B06", c["id"], v)
    k = sheet(wb, "04_")
    for i, kk in enumerate(B["kano"]):
        if k.Cells(7 + i, 10).Value != kk["interp"]:
            bad += 1
            print("  Кано", kk["id"], k.Cells(7 + i, 10).Value)
    d = sheet(wb, "09_")
    print("  дашборд B5..B11:", [d.Cells(r, 2).Value for r in range(5, 12)])
    return bad


def chk_c(wb):
    Cc = calc_c()
    ws = sheet(wb, "09_")
    bad = 0
    for i, m in enumerate(Cc["mat"]):
        if abs(ws.Cells(4 + i, 17).Value - m["Q"]) > 1e-9:
            bad += 1
            print("  C09", m["id"], ws.Cells(4 + i, 17).Value)
    st = sheet(wb, "10_")
    for i, s in enumerate(Cc["std"]):
        if st.Cells(4 + i, 6).Value != s["F"]:
            bad += 1
            print("  C10", s["pr"], st.Cells(4 + i, 5).Value, st.Cells(4 + i, 6).Value)
    sv = sheet(wb, "12_")
    print("  12_Сводка B5..B11:", [sv.Cells(r, 2).Value for r in range(5, 12)], "топ:", [sv.Cells(r, 6).Value for r in range(6, 10)])
    return bad


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT, exist_ok=True)
    xl = w32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        run(xl, "Шаблон_классификации_компаний_по_уровням_конкуренции.xlsx", "ЛР5_1_Классификация_NotaCode.xlsx",
            ["k_01_kompanii.tsv", "k_02_dokazatelstva.tsv", "k_04_ocenka.tsv", "k_04_fix.tsv", "k_07_karta.tsv", "k_06_top10_fix.tsv",
             "k_06_vyvod.tsv", "k_vvod.tsv"], chk_a)
        run(xl, "Шаблон_анализа_конкурентов_Левитт_Кано.xlsx", "ЛР5_2_Левитт_Кано_NotaCode.xlsx",
            [f"l_{x}.tsv" for x in ["01_konkurenty", "02_karta", "03_levitt", "04_kano", "05_matrica", "06_sravnenie",
                                    "06_fix", "07_standart", "08_preim", "09_dash", "vvod"]], chk_b)
        run(xl, "Шаблон_анализа_бизнес_модели_конкурентов.xlsx", "ЛР5_3_Бизнес_модель_NotaCode.xlsx",
            [f"b_{x}.tsv" for x in ["02_konkurenty", "03_fakty", "04_profil", "05_tovar", "06_monet", "07_kanaly", "08_ops",
                                    "09_ves", "09_matrica", "10_standart", "10_fix", "11_preim", "12_vyvod", "00_pasport", "vvod"]], chk_c)
    finally:
        xl.Quit()

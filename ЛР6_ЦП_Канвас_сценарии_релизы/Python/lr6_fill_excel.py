# -*- coding: utf-8 -*-
"""ЛР6. Заполнение копий 5 Excel-шаблонов данными NotaCode через Excel COM (pywin32)
тем же форматом VBA/data/lr6_N.tsv, что и VBA-модули, + проверка рассчитанных Excel значений.

Нужны установленный Microsoft Excel и pywin32. Оригиналы в Материалы не изменяются.
Запуск: python -B lr6_fill_excel.py
Результат: ../Excel_заполненные/ЛР6_N_*.xlsx и out/excel_check.json
"""
import json
import os
import re
import shutil

import win32com.client as win32

from lr6_common import LAB, MATERIALS, OUTDIR, TEMPLATES, VBA_DATA, col_index

DEST = os.path.join(LAB, "Excel_заполненные")
NUM = re.compile(r"^-?\d+(\.\d+)?$")
CHECKS = {  # что прочитать после пересчёта
    "lr6_1": [("Проблемы клиентов", "K5:L12"), ("Проверка связности", "G5:H7"), ("Дашборд", "B5:B10")],
    "lr6_2": [("02_Канвас", "J3:J11"), ("11_Дашборд", "B4:B9")],
    "lr6_3": [("Пользовательские сценарии", "R4:S16"), ("Функциональные требования", "M4:O17"), ("Связность", "F4:H13"), ("Дашборд", "B5:B14")],
    "lr6_4": [("Проверка_связности", "D5:E11"), ("Приоритизация", "I5:J28"), ("Дашборд", "B5:B12")],
    "lr6_5": [("Реестр_единиц", "Q2:S39"), ("Оценка_релизов", "B4:J6"), ("Дашборд", "B4:B9")],
}


def load(wb, path):
    ws = None
    n = 0
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.rstrip("\n")
            if not line or line.startswith("# "):
                continue
            f = line.split("\t")
            if f[0] == "#SHEET":
                ws = wb.Worksheets(f[1])
            elif f[0] == "R":
                r, c = int(f[1]), col_index(f[2])
                for j, v in enumerate(f[3:]):
                    cell = ws.Cells(r, c + j)
                    if v == "~":
                        continue
                    if v == "":
                        cell.ClearContents()
                    elif NUM.match(v):
                        cell.Value = float(v)
                    else:
                        cell.Value = v
                n += 1
    return n


def fixes(key, wb):
    if key == "lr6_5":
        ws = wb.Worksheets("Реестр_единиц")
        for r in range(2, 102):
            ws.Cells(r, 17).Formula = (f'=IF(A{r}="","",ROUND((G{r}*0.18)+(H{r}*0.15)+(I{r}*0.12)+(J{r}*0.08)+(K{r}*0.13)+(L{r}*0.10)+(M{r}*0.12)+(N{r}*0.12)-(O{r}*0.07)+(P{r}*0.10),2))')
            ws.Cells(r, 18).Formula = (f'=IF(A{r}="","",ROUND((I{r}*0.16)+(J{r}*0.20)+(K{r}*0.12)+(L{r}*0.10)+((6-M{r})*0.08)+((6-N{r})*0.08)+(O{r}*0.06)+(P{r}*0.14)+(G{r}*0.06),2))')
    if key == "lr6_3":
        ws = wb.Worksheets("Пользовательские сценарии")
        ws.Range("N4:Q203").Validation.Delete()
        ws.Range("N4:P203").Validation.Add(3, 1, 1, "1,2,3,4,5")
        ws.Range("Q4:Q203").Validation.Add(3, 1, 1, "1,2,3,4")


def main():
    os.makedirs(DEST, exist_ok=True)
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    result = {}
    try:
        for key, tmpl in TEMPLATES.items():
            dst = os.path.join(DEST, key.replace("lr6_", "ЛР6_") + "_" + tmpl.replace("Шаблон_", "NotaCode_"))
            shutil.copyfile(os.path.join(MATERIALS, tmpl), dst)
            wb = xl.Workbooks.Open(dst)
            n = load(wb, os.path.join(VBA_DATA, key + ".tsv"))
            fixes(key, wb)
            xl.CalculateFull()
            result[key] = {"rows": n}
            for sh, rng in CHECKS[key]:
                vals = wb.Worksheets(sh).Range(rng).Value
                result[key][sh + "!" + rng] = [[v for v in row] for row in vals]
            wb.Save()
            wb.Close(False)
            print("OK", os.path.basename(dst), n, "строк")
    finally:
        xl.Quit()
    with open(os.path.join(OUTDIR, "excel_check.json"), "w", encoding="utf-8") as fh:
        json.dump(result, fh, ensure_ascii=False, indent=1, default=str)


if __name__ == "__main__":
    main()

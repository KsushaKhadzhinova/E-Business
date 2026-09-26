# -*- coding: utf-8 -*-
"""ЛР6. Общие константы и функции для скриптов расчёта и визуализации.

Все исходные данные NotaCode для ЛР6 лежат в модулях lr6_d1_cp.py … lr6_d5_releases.py.
Строки листов Excel записаны в порядке столбцов шаблона; значение None означает
«ячейка с формулой шаблона — не трогать» (в TSV для VBA кодируется символом ~).
"""
import os

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
LAB = os.path.normpath(os.path.join(HERE, ".."))
FIG = os.path.join(LAB, "Рисунки")
OUTDIR = os.path.join(HERE, "out")
VBA_DATA = os.path.join(LAB, "VBA", "data")
MATERIALS = os.path.join(LAB, "Материалы")

SKIP = None

TEMPLATES = {
    "lr6_1": "Шаблон_формулирования_ценностного_предложения_и_цифрового_товара.xlsx",
    "lr6_2": "Шаблон_сборки_бизнес_модели_по_Канвас.xlsx",
    "lr6_3": "Шаблон_бизнес_логики_и_пользовательских_сценариев.xlsx",
    "lr6_4": "Шаблон_перехода_к_страницам_и_экранам.xlsx",
    "lr6_5": "Шаблон_деления_разработки_сайта_на_релизы.xlsx",
}

# Цвета релизов (одинаковые на всех рисунках)
REL_COLORS = {
    "R0": "#9e9e9e",
    "R1": "#2e7d32",
    "R2": "#1565c0",
    "R3": "#ef6c00",
    "R4": "#6a1b9a",
    "Искл.": "#c62828",
}


def setup_fonts():
    plt.rcParams["font.family"] = ["DejaVu Sans", "Arial", "sans-serif"]
    plt.rcParams["axes.unicode_minus"] = False


def ensure_dirs():
    for d in (FIG, OUTDIR, VBA_DATA):
        os.makedirs(d, exist_ok=True)


def col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letter):
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - 64)
    return n


def md_table(header, rows):
    """Markdown-таблица из списка строк."""
    def cell(v):
        if v is None:
            return ""
        if isinstance(v, float):
            s = f"{v:.2f}".rstrip("0").rstrip(".") if abs(v - round(v)) > 1e-9 else str(int(round(v)))
            return s.replace(".", ",")
        return str(v).replace("|", "/").replace("\n", " ")
    out = ["| " + " | ".join(header) + " |", "|" + "---|" * len(header)]
    for r in rows:
        out.append("| " + " | ".join(cell(v) for v in r) + " |")
    return "\n".join(out)


def fmt(v, nd=2):
    if isinstance(v, (int,)) and not isinstance(v, bool):
        return str(v)
    return f"{v:.{nd}f}".replace(".", ",")

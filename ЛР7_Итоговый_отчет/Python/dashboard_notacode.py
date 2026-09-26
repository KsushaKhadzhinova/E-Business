# -*- coding: utf-8 -*-
"""
ЛР7. Сводный дашборд бизнес-анализа NotaCode (одна картинка 2x3).

Панели:
  1) Тренд спроса Google Trends (PlantUML и text to diagram, 40 мес.) — ЛР1
  2) TAM / SAM / SOM (диапазоны, BYN) и целевой KPI выручки — ЛР2-3
  3) Радар пяти сил Портера + цифровые усилители — ЛР4
  4) Конкуренты: взвешенный балл N по 10 критериям и уровни — ЛР5
  5) ИРП единиц разработки по релизам R1–R4 (методика ЛР6) — ЛР6
  6) Пороги первого релиза (ЛР6) и KPI первого года (03-smart-goals.md)

Все числа собраны в блоке DATA ниже. Если в ОТЧЕТ.md ЛР1–ЛР6 цифры
изменятся — поправьте только этот блок и перезапустите скрипт:
    python dashboard_notacode.py
Результат: ../Рисунки/Рис_7_1_Дашборд_NotaCode.png (+ отдельные панели).
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

plt.rcParams["font.family"] = ["DejaVu Sans", "Arial"]
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False

OUT = Path(__file__).resolve().parent.parent / "Рисунки"
OUT.mkdir(exist_ok=True)

INK = "#0b0b0b"
INK2 = "#52514e"
GRID = "#e4e3df"
S1 = "#2a78d6"
S2 = "#eb6834"
S3 = "#1baf7a"

# ---------------------------------------------------------------- DATA
# ЛР1: Google Trends, Worldwide, снято 17.09.2026 (архив ЛР1, лист 01_Данные_и_расчёт)
MONTHS = [f"{y}-{m:02d}" for y in range(2021, 2025) for m in range(1, 13)][8:]
PLANTUML = [16, 20, 21, 20, 22, 28, 25, 25, 27, 22, 26, 22, 24, 25, 28, 26,
            25, 33, 31, 30, 31, 30, 33, 26, 30, 34, 38, 34, 37, 35, 46, 48,
            47, 49, 40, 37, 39, 50, 51, 53]
TEXT2DIAG = [9, 7, 10, 8, 11, 15, 12, 9, 9, 8, 4, 3, 9, 7, 7, 7,
             7, 8, 7, 9, 8, 7, 5, 5, 8, 8, 8, 9, 7, 9, 9, 11,
             10, 6, 5, 6, 8, 10, 11, 12]

# ЛР2-3: итоговая оценка рынка, BYN в год (ОТЧЕТ.md ЛР2-3, таблица 13)
MARKET = [  # уровень, нижняя граница, верхняя граница
    ("TAM", 2_650_000, 14_380_000),
    ("SAM", 23_000, 66_000),
    ("SOM\n(1-й год)", 1_400, 2_900),
]
KPI_REVENUE = 6_480  # целевой KPI: 45 платящих (ЛР2-3)

# ЛР4: пять сил Портера, шкала 1/3/5 (ОТЧЕТ.md ЛР4, раздел 10)
PORTER = {
    "Конкуренция\nигроков": 5,
    "Новые\nучастники": 3,
    "Сила\nпокупателей": 5,
    "Сила\nпоставщиков": 3,
    "Заменители": 5,
    "Цифровые\nусилители": 3,
}

# ЛР5: взвешенный балл N (0–5) и итоговый уровень (ОТЧЕТ.md ЛР5, таблица 3)
COMPETITORS = [
    ("PlantUML", 3.75, "Прямой"), ("Eraser", 3.49, "Близкий"),
    ("draw.io", 3.27, "Косвенный/заменитель"), ("Mermaid", 3.25, "Косвенный/заменитель"),
    ("Gleek", 3.23, "Косвенный/заменитель"), ("Ramus", 3.22, "Косвенный/заменитель"),
    ("D2", 3.09, "Косвенный/заменитель"), ("VP Online", 2.94, "Косвенный/заменитель"),
    ("Mermaid Chart", 2.93, "Косвенный/заменитель"), ("Lucidchart", 2.90, "Косвенный/заменитель"),
    ("dbdiagram.io", 2.82, "Косвенный/заменитель"), ("AI-ассистенты", 2.47, "Смежный/потенциальный"),
    ("Structurizr", 2.46, "Смежный/потенциальный"), ("Visio", 2.35, "Смежный/потенциальный"),
    ("Sparx EA", 2.34, "Смежный/потенциальный"), ("Camunda", 2.32, "Смежный/потенциальный"),
    ("Graphviz", 2.05, "Смежный/потенциальный"), ("Kroki", 1.79, "Наблюдение"),
    ("Miro", 1.68, "Смежный/потенциальный"), ("Figma/FigJam", 1.52, "Наблюдение"),
]
LEVEL_COL = {"Прямой": "#0d3b73", "Близкий": "#2a78d6", "Косвенный/заменитель": "#6fa6e8",
             "Смежный/потенциальный": "#a9c9f0", "Наблюдение": "#d3e3f7"}
THRESHOLDS = [(4.2, "прямой"), (3.4, "близкий"), (2.5, "косвенный"), (1.6, "смежный")]

# ЛР6: ИРП единиц разработки по методике (ОТЧЕТ.md ЛР6, таблица А.19), сгруппировано по релизам
IRP = {
    1: [3.65, 3.25, 3.25, 3.25, 3.45, 3.55, 3.05, 3.0, 3.25, 3.0, 2.75, 3.65, 2.65, 3.55, 3.15, 2.9, 1.65],
    2: [2.85, 2.55, 1.85, 2.8, 2.45, 2.2, 1.4, 1.95, 1.8, 1.65],
    3: [1.75, 2.25, 2.05, 2.3, 2.45, 1.65],
    4: [2.4, 1.3, 1.35, 1.65],
}

# Пороги R1 (ЛР6, протокол утверждения) и KPI первого года (03-smart-goals.md)
KPI_PCT = [("Активация D0 (KPI-02)", 60), ("D7 – порог R1→R2 (ЛР6)", 25),
           ("D30 (KPI-04)", 15), ("Заявки Pro / активные Free (ЛР6)", 3),
           ("Конверсия Free→Pro (KPI-05)", 3), ("Churn Pro / мес, не более (KPI-07)", 8)]
KPI_ABS = "Пересмотр ЦП: активация < 20 % или D7 < 10 % · KPI-01: 1 500 регистраций"
# ---------------------------------------------------------------- /DATA


def style(ax, title):
    ax.set_title(title, loc="left", fontsize=12, color=INK, fontweight="bold", pad=10)
    ax.tick_params(colors=INK2, labelsize=9)
    for s in ax.spines.values():
        s.set_color(GRID)


def panel_trend(ax):
    x = np.arange(len(MONTHS))
    ax.plot(x, PLANTUML, color=S1, lw=2, label="PlantUML")
    ax.plot(x, TEXT2DIAG, color=S2, lw=2, label="text to diagram")
    k, b = np.polyfit(x, PLANTUML, 1)
    ax.plot(x, k * x + b, color=S1, lw=1, ls="--", alpha=0.7,
            label=f"Линейный тренд PlantUML (+{k:.3f} п./мес)".replace(".", ",", 1))
    m1, m2 = np.mean(PLANTUML[:12]), np.mean(PLANTUML[-12:])
    ax.hlines(m1, 0, 11, color=INK2, lw=1.5)
    ax.hlines(m2, len(x) - 12, len(x) - 1, color=INK2, lw=1.5)
    ax.text(0, m1 - 4, "22,83", fontsize=8, color=INK2)
    ax.text(len(x) - 12, m2 + 9, "44,33 (+94,2 %)", fontsize=9, color=INK)
    ticks = [i for i, m in enumerate(MONTHS) if m.endswith("-01")]
    ax.set_xticks(ticks, [MONTHS[i][:7] for i in ticks])
    ax.set_ylabel("Индекс Google Trends (0–100)", color=INK2, fontsize=9)
    ax.set_ylim(0, 65)
    ax.grid(axis="y", color=GRID, lw=0.8)
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    style(ax, "1. Спрос: Google Trends, 09.2021–12.2024 (ЛР1)")


def fmt_byn(v):
    if v >= 1e6:
        return f"{v / 1e6:.2f} млн".replace(".", ",")
    return f"{v / 1e3:.1f}".replace(".", ",").replace(",0", "") + " тыс."


def panel_market(ax):
    names = [m[0] for m in MARKET]
    lo = [m[1] for m in MARKET]
    hi = [m[2] for m in MARKET]
    cols = ["#a9c9f0", "#6fa6e8", S1]
    x = np.arange(len(MARKET))
    ax.bar(x, lo, color=cols, width=0.55, edgecolor="white", linewidth=2)
    ax.bar(x, [h - l for h, l in zip(hi, lo)], bottom=lo, color=cols, alpha=0.35,
           width=0.55, edgecolor="white", linewidth=2)
    ax.set_yscale("log")
    ax.set_ylim(500, 80_000_000)
    for xi, l, h in zip(x, lo, hi):
        ax.text(xi, h * 1.3, f"{fmt_byn(l)} – {fmt_byn(h)}", ha="center", fontsize=9.5, color=INK, fontweight="bold")
    ax.axhline(KPI_REVENUE, color=S2, lw=1.5, ls="--")
    ax.text(0.6, KPI_REVENUE * 1.15, "Целевой KPI: 45 платящих = 6 480 BYN", fontsize=8, color=INK2)
    ax.set_xticks(x, names)
    ax.set_ylabel("BYN в год (лог. шкала); светлая часть – диапазон", color=INK2, fontsize=9)
    ax.grid(axis="y", color=GRID, lw=0.8, which="major")
    style(ax, "2. Рынок: TAM / SAM / SOM (ЛР2–3)")


def panel_radar(ax):
    labels = list(PORTER)
    vals = list(PORTER.values())
    ang = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    ax.plot(ang + ang[:1], vals + vals[:1], color=S1, lw=2)
    ax.fill(ang + ang[:1], vals + vals[:1], color=S1, alpha=0.18)
    ax.scatter(ang, vals, color=S1, s=30, zorder=3, edgecolor="white", linewidth=1.5)
    ax.set_xticks(ang, labels, fontsize=8, color=INK)
    ax.tick_params(axis="x", pad=12)
    ax.set_ylim(0, 5)
    ax.set_yticks([1, 3, 5], ["1", "3", "5"], fontsize=8, color=INK2)
    ax.grid(color=GRID)
    ax.spines["polar"].set_color(GRID)
    avg = f"{sum(vals) / len(vals):.2f}".replace(".", ",")
    ax.set_title(f"3. Пять сил Портера: среднее {avg} (ЛР4)", loc="left",
                 fontsize=12, color=INK, fontweight="bold", pad=18)


def panel_competitors(ax):
    data = sorted(COMPETITORS, key=lambda t: t[1])
    y = np.arange(len(data))
    ax.barh(y, [d[1] for d in data], color=[LEVEL_COL[d[2]] for d in data], height=0.66,
            edgecolor="white", linewidth=1.5)
    for yi, d in zip(y, data):
        ax.text(d[1] + 0.05, yi, f"{d[1]:.2f}".replace(".", ","), va="center", fontsize=7, color=INK)
    ax.set_yticks(y, [d[0] for d in data], fontsize=7.5, color=INK)
    for thr, name in THRESHOLDS:
        ax.axvline(thr, color=INK2, lw=0.8, ls=":")
        ax.text(thr, len(data) - 0.1, name, fontsize=7, color=INK2, ha="center")
    ax.set_xlim(0, 4.8)
    ax.set_xlabel("Взвешенный балл N (0–5), пороги шаблона ЛР5", color=INK2, fontsize=9)
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in LEVEL_COL.values()]
    ax.legend(handles, list(LEVEL_COL), frameon=False, fontsize=7, loc="lower right",
              title="Уровень (итог ЛР5)", title_fontsize=7.5)
    style(ax, "4. Конкуренты: 20 компаний (ЛР5)")


def panel_releases(ax):
    rel_col = {1: S1, 2: S3, 3: "#eda100", 4: "#b9b8b2"}
    for r, vals in IRP.items():
        xs = np.full(len(vals), r) + np.linspace(-0.18, 0.18, len(vals))
        ax.scatter(xs, vals, color=rel_col[r], s=36, edgecolor="white", linewidth=1, zorder=3)
        m = np.mean(vals)
        ax.hlines(m, r - 0.3, r + 0.3, color=INK, lw=2)
        ax.text(r + 0.32, m, f"ср. {m:.2f}".replace(".", ","), va="center", fontsize=8, color=INK)
    ax.axhline(3.5, color=S2, lw=1.2, ls="--")
    ax.text(0.55, 3.56, "ИРП ≥ 3,5 – первый релиз по методике", fontsize=7.5, color=INK2)
    ax.set_xticks(list(IRP), [f"R{r}\n{len(v)} единиц" for r, v in IRP.items()])
    ax.set_xlim(0.5, 4.8)
    ax.set_ylim(1, 4)
    ax.set_ylabel("ИРП по методике ЛР6", color=INK2, fontsize=9)
    ax.grid(axis="y", color=GRID, lw=0.8)
    style(ax, "5. ИРП единиц разработки по релизам (ЛР6)")


def panel_kpi(ax):
    names = [k[0] for k in KPI_PCT][::-1]
    vals = [k[1] for k in KPI_PCT][::-1]
    y = np.arange(len(names))
    ax.barh(y, vals, color=S1, height=0.55, edgecolor="white", linewidth=2)
    for yi, v in zip(y, vals):
        ax.text(v + 1, yi, f"{v} %", va="center", fontsize=9, color=INK, fontweight="bold")
    ax.set_yticks(y, names, fontsize=8.5, color=INK)
    ax.set_xlim(0, 75)
    ax.set_xlabel("Пороговое / целевое значение, %", color=INK2, fontsize=9)
    ax.text(0, -1.3, KPI_ABS, fontsize=8, color=INK2, transform=ax.transData)
    style(ax, "6. Метрики первого релиза и KPI")



def build(fig_axes=None):
    fig = plt.figure(figsize=(20, 11.5), facecolor="white")
    gs = fig.add_gridspec(2, 3, hspace=0.42, wspace=0.5, left=0.075, right=0.955, top=0.9, bottom=0.07)
    panel_trend(fig.add_subplot(gs[0, 0]))
    panel_market(fig.add_subplot(gs[0, 1]))
    panel_radar(fig.add_subplot(gs[0, 2], projection="polar"))
    panel_competitors(fig.add_subplot(gs[1, 0]))
    panel_releases(fig.add_subplot(gs[1, 1]))
    panel_kpi(fig.add_subplot(gs[1, 2]))
    fig.suptitle("NotaCode — сводный дашборд бизнес-анализа (ЛР1–ЛР6)", x=0.05, ha="left",
                 fontsize=17, fontweight="bold", color=INK)
    fig.text(0.05, 0.925, "Источники: Google Trends 17.09.2026; ВВП РБ (Macrotrends), ПВТ (opencompanyinbelarus.com); "
             "SimilarWeb 17/23.09.2026; отчёты ЛР1–ЛР6. KPI и коэффициенты — допущения.",
             fontsize=9, color=INK2)
    path = OUT / "Рис_7_1_Дашборд_NotaCode.png"
    fig.savefig(path, dpi=150)
    plt.close(fig)
    print("OK:", path)


def single(fn, name, polar=False, size=(8, 5.5)):
    fig = plt.figure(figsize=size, facecolor="white")
    ax = fig.add_subplot(111, projection="polar" if polar else None)
    fn(ax)
    fig.tight_layout()
    path = OUT / name
    fig.savefig(path, dpi=150)
    plt.close(fig)
    print("OK:", path)


if __name__ == "__main__":
    build()
    single(panel_trend, "Рис_7_2_Тренд_спроса.png")
    single(panel_market, "Рис_7_3_TAM_SAM_SOM.png")
    single(panel_radar, "Рис_7_5_Радар_Портера.png", polar=True, size=(7, 6.5))
    single(panel_competitors, "Рис_7_4_Карта_конкурентов.png")
    single(panel_releases, "Рис_7_6_ИРП_релизов.png", size=(9, 5.5))
    single(panel_kpi, "Рис_7_7_KPI.png")

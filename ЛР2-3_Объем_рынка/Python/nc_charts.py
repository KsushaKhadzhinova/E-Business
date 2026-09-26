# -*- coding: utf-8 -*-
"""Диаграммы ЛР2-3 (NotaCode). Все PNG сохраняются в ../Рисунки/.
Запуск всех рисунков:  python nc_charts.py
Отдельные группы: 01_tam_sam_som.py, 02_voronka.py, 03_sravnenie_metodov.py
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

import nc_model as M

OUT = Path(__file__).resolve().parent.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.grid": True, "grid.color": "#e4e3df",
                     "grid.linewidth": 0.8, "axes.axisbelow": True, "axes.edgecolor": "#898781",
                     "text.color": "#0b0b0b", "axes.labelcolor": "#52514e", "xtick.color": "#52514e",
                     "ytick.color": "#52514e", "figure.facecolor": "#fcfcfb", "axes.facecolor": "#fcfcfb"})
C3 = ["#2a78d6", "#eb6834", "#1baf7a"]            # осторожный / базовый / оптимистичный
CAT = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
R = M.compute_all()


def num(x, pos=None):
    if x >= 1e6:
        return f"{x / 1e6:.1f} млн".replace(".", ",")
    if x >= 1e3:
        return f"{x / 1e3:.0f} тыс."
    return f"{x:.0f}"


def lab(x):
    return M.fmt(x)


def save(fig, name):
    p = OUT / name
    fig.tight_layout()
    fig.savefig(p, dpi=150)
    plt.close(fig)
    print("saved", p)


def fig_tam_sam_som():
    pk = R["potential"]["scen"]
    fig, ax = plt.subplots(figsize=(9, 5))
    w = 0.26
    for j, key in enumerate(("TAM", "SAM", "SOM")):
        for i in range(3):
            x = j + (i - 1) * w
            v = pk[i][key]
            ax.bar(x, v, w - 0.02, color=C3[i], label=M.SCEN[i] if j == 0 else None)
            ax.text(x, v * 1.12, lab(v), ha="center", fontsize=8, color="#52514e")
    ax.set_yscale("log")
    ax.yaxis.set_major_formatter(FuncFormatter(num))
    ax.set_xticks(range(3), ["TAM\n(все потенциальные\nпользователи × Pro)", "SAM\n(потребность × онлайн ×\nготовность платить)",
                             "SOM\n(достижимая доля\n1 / 3 / 7 %)"])
    ax.set_ylabel("BYN в год (логарифмическая шкала)")
    ax.set_title("TAM / SAM / SOM NotaCode по сценариям (метод потенциальных клиентов, РБ)")
    ax.legend(frameon=False)
    save(fig, "fig01_tam_sam_som_scenarii.png")

    # вложенная оценка базового сценария из трёх методов
    td = R["topdown"]["scen"][1]
    rows = [("TAM сверху вниз (ВВП → инструменты моделирования)", td["TAM"]),
            ("TAM снизу вверх (33 315 польз. × цена Pro)", R["potential"]["TAM"]),
            ("SAM сверху вниз (diagram-as-code × нотации)", td["SAM"]),
            ("SAM снизу вверх (адресуемые 590 польз.)", R["potential"]["SAM"]),
            ("SAM по Similarweb (методика ЛР3)", R["lr3"]["scen"][1]["rev_y"]),
            ("SOM потенциальные клиенты (3 %)", R["potential"]["SOM"]),
            ("SOM цифровая воронка NotaCode", R["funnel"]["scen"][1]["rev_y_rep"])]
    fig, ax = plt.subplots(figsize=(9.5, 4.8))
    cols = ["#184f95", "#2a78d6", "#1c5cab", "#5598e7", "#86b6ef", "#eb6834", "#d95926"]
    for i, (n, v) in enumerate(rows[::-1]):
        ax.barh(i, v, color=cols[::-1][i], height=0.62)
        ax.text(v * 1.1, i, lab(v) + " BYN", va="center", fontsize=8.5)
    ax.set_yticks(range(len(rows)), [r[0] for r in rows[::-1]])
    ax.set_xscale("log")
    ax.xaxis.set_major_formatter(FuncFormatter(num))
    ax.set_xlim(500, 6e7)
    ax.set_xlabel("BYN в год, базовый сценарий (логарифмическая шкала)")
    ax.set_title("Воронка TAM → SAM → SOM (базовый сценарий, разные методы)")
    ax.grid(axis="y", visible=False)
    save(fig, "fig02_tam_sam_som_voronka.png")


def fig_voronka():
    f = R["funnel"]
    s = f["scen"][1]
    steps = [("Релевантные посещения", f["rel_m"]), ("Реализованный охват (80 %)", s["reach_visits"]),
             ("Регистрации Free (4 %)", s["leads"]), ("Подписки Pro (3 %)", s["sales"])]
    fig, ax = plt.subplots(figsize=(9, 4.2))
    mx = steps[0][1]
    for i, (n, v) in enumerate(steps):
        wdt = max(v / mx, 0.02)
        ax.barh(len(steps) - 1 - i, wdt, left=(1 - wdt) / 2, color=C3[0] if i < 2 else C3[1], height=0.7)
        ax.text(0.5, len(steps) - 1 - i, f"{n}: {M.fmt(v, 1 if v < 100 else 0)} в мес.", ha="center", va="center",
                color="#0b0b0b", fontsize=10,
                bbox=dict(boxstyle="round,pad=0.2", fc="#fcfcfb", ec="none", alpha=0.85))
    ax.set_xlim(0, 1)
    ax.axis("off")
    ax.set_title(f"Цифровая воронка NotaCode, базовый сценарий: {M.fmt(s['pro_year'], 1)} Pro/год × 120 BYN = "
                 f"{M.fmt(s['rev_y_rep'])} BYN/год")
    save(fig, "fig03_voronka_bazovyj.png")

    fig, ax = plt.subplots(figsize=(8, 4.2))
    vals = [x["rev_y_rep"] for x in f["scen"]]
    ax.bar(M.SCEN, vals, color=C3, width=0.55)
    for i, v in enumerate(vals):
        ax.text(i, v + max(vals) * 0.02, lab(v) + " BYN", ha="center")
    ax.axhline(R["kpi"]["rev_byn"], color="#52514e", ls="--", lw=1.2)
    ax.text(2.3, R["kpi"]["rev_byn"] * 1.03, "KPI: 1500 рег. × 3 % × 144 BYN = 6 480", ha="right", fontsize=8.5,
            color="#52514e")
    ax.set_ylabel("BYN в год")
    ax.set_title("Цифровая воронка NotaCode: годовая выручка по сценариям")
    save(fig, "fig04_voronka_scenarii.png")


def _methods(d, title, name, xlim):
    items = list(d.items())
    fig, ax = plt.subplots(figsize=(10, 0.55 * len(items) + 1.6))
    h = 0.26
    for k, (n, vals) in enumerate(items):
        y = len(items) - 1 - k
        for i in range(3):
            ax.barh(y + (1 - i) * h, vals[i], h - 0.03, color=C3[i], label=M.SCEN[i] if k == 0 else None)
        ax.text(vals[2] * 1.15, y + (1 - 2) * h, lab(vals[2]), va="center", fontsize=7.5, color="#52514e")
        ax.text(vals[1] * 1.15, y, lab(vals[1]), va="center", fontsize=7.5, color="#52514e")
    ax.set_yticks(range(len(items)), [n for n, _ in items][::-1])
    ax.set_xscale("log")
    ax.set_xlim(*xlim)
    ax.xaxis.set_major_formatter(FuncFormatter(num))
    ax.set_xlabel("BYN в год (логарифмическая шкала)")
    ax.set_title(title)
    ax.grid(axis="y", visible=False)
    ax.legend(frameon=False, loc="upper center", bbox_to_anchor=(0.5, -0.14), ncol=3, fontsize=8.5)
    save(fig, name)


def fig_sravnenie():
    _methods(R["som"], "Сверка методов: достижимая выручка NotaCode (SOM, 1-й год)", "fig05_sravnenie_som.png",
             (20, 8e5))
    _methods(R["sam"], "Сверка методов: объём доступного рынка (SAM / рынок в год)", "fig06_sravnenie_sam.png",
             (2e3, 1e7))


def fig_sw():
    rows = M.sw_rows()
    l3 = R["lr3"]
    tot = sum(r["rel_visits"] for r in rows)
    fig, axs = plt.subplots(1, 2, figsize=(11, 4.3))
    names = [r["domain"] for r in rows]
    raw_tot = sum(r["visits"] for r in rows)
    for ax, vals, ttl, h in ((axs[0], [r["visits"] / raw_tot * 100 for r in rows], "Все визиты (мировой трафик)",
                              l3["conc_raw"]["HHI"]),
                             (axs[1], [r["rel_visits"] / tot * 100 for r in rows],
                              "Релевантный трафик РФ/РБ × категория", l3["conc_rel"]["HHI"])):
        ax.bar(names, vals, color=CAT[:4], width=0.6)
        for i, v in enumerate(vals):
            ax.text(i, v + 1.5, f"{v:.1f} %".replace(".", ","), ha="center", fontsize=9)
        ax.set_ylim(0, 100)
        ax.set_ylabel("доля, %")
        ax.set_title(f"{ttl}\nHHI = {h:.0f} ({M.hhi_level(h)} концентрация)")
        ax.tick_params(axis="x", labelrotation=15)
    fig.suptitle(f"Доли цифрового внимания (Similarweb, 17–23.09.2026); CR3 = {l3['conc_rel']['CR3']:.1f} %".replace(".", ","))
    save(fig, "fig07_similarweb_doli_hhi.png")


def fig_pay():
    e = R["payment"]["econ"]
    fig, axs = plt.subplots(1, 2, figsize=(11, 4.2))
    n = [x["name"] for x in e]
    ax = axs[0]
    ax.bar(n, [x["index"] for x in e], color=CAT[:4], width=0.6)
    ax.axhline(1, color="#d03b3b", ls="--", lw=1)
    ax.axhline(1.5, color="#0ca30c", ls="--", lw=1)
    ax.text(3.4, 1.03, "1,0 достаточная", ha="right", fontsize=8)
    ax.text(3.4, 1.53, "1,5 высокая", ha="right", fontsize=8)
    for i, x in enumerate(e):
        ax.text(i, x["index"] + 0.08, f'{x["index"]:.2f}'.replace(".", ","), ha="center")
    ax.set_title("Индекс платёжеспособности = бюджет / 12 BYN")
    ax.tick_params(axis="x", labelrotation=12)
    ax = axs[1]
    ax.bar(n, [x["K"] for x in e], color=CAT[:4], width=0.6)
    for y, t in ((1, "1"), (2, "2"), (5, "5")):
        ax.axhline(y, color="#898781", ls=":", lw=1)
    for i, x in enumerate(e):
        ax.text(i, x["K"] + 0.15, f'{x["K"]:.2f}'.replace(".", ","), ha="center")
    ax.set_title("Коэффициент оправданности = эффект / цена")
    ax.tick_params(axis="x", labelrotation=12)
    save(fig, "fig08_platezhesposobnost.png")


def fig_gt():
    try:
        import openpyxl
        wb = openpyxl.load_workbook(Path(__file__).resolve().parents[2] /
                                    "_архив_DiagramCode/labs/lab1-demand/ЛР1_Тренд_Сезонность_DiagramCode.xlsx")
        ws = wb.worksheets[0]
        rows = [(ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value) for r in range(3, 43)]
    except Exception as exc:  # noqa: BLE001
        print("GT: архив недоступен, рисую последние 12 мес.", exc)
        rows = list(zip(M.GT_MONTHS, M.GT_PLANTUML, M.GT_TEXT2DIAGRAM))
    fig, ax = plt.subplots(figsize=(10, 4))
    x = range(len(rows))
    ax.plot(x, [r[1] for r in rows], color=C3[0], lw=2, label="PlantUML")
    ax.plot(x, [r[2] for r in rows], color=C3[1], lw=2, label="text to diagram")
    ax.set_xticks(list(x)[::3], [r[0] for r in rows][::3], rotation=45)
    ax.set_ylabel("индекс Google Trends (0–100)")
    ax.set_title("Google Trends, Worldwide, 2021-09…2024-12 (снято 17.09.2026)")
    ax.legend(frameon=False)
    save(fig, "fig09_google_trends.png")


GROUPS = {"tam": [fig_tam_sam_som], "voronka": [fig_voronka], "sravnenie": [fig_sravnenie, fig_sw, fig_pay, fig_gt]}

if __name__ == "__main__":
    for fs in GROUPS.values():
        for fn in fs:
            fn()

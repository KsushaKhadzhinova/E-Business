# -*- coding: utf-8 -*-
"""
ЛР5 NotaCode — рисунки. Запуск: python -B lr5_charts.py  →  ../Рисунки/*.png
  fig_01_karta_konkurentnogo_polya.png  карта (X = ср(C1,C3), Y = ср(C2,C7)) с зонами 07_Карта
  fig_02_vzveshennye_bally.png          взвешенный балл N (04_Оценка) с порогами уровней
  fig_03_levitt_heatmap.png             средняя оценка признака по Левитту: конкурент × уровень
  fig_04_kano_heatmap.png               присутствие требований Кано у 12 конкурентов
  fig_05_kano_raspredelenie.png         распределение требований по категориям Кано и доли конкурентов
  fig_06_sravnenie_P01_P25.png          итоговый взвешенный балл 06_Сравнение
  fig_07_biznes_model.png               итог бизнес-модели с учётом надёжности (09_Матрица)
"""
import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Rectangle, Patch

import lr5_data as D
from lr5_calc import calc_a, calc_b, calc_c

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["font.size"] = 9
OUT = os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "Рисунки"))
os.makedirs(OUT, exist_ok=True)

LVL_COL = {
    "Прямой конкурент": "#c0392b",
    "Близкий альтернативный конкурент": "#e67e22",
    "Косвенный конкурент / заменитель": "#2e86c1",
    "Смежная или потенциальная конкуренция": "#7d8c8d",
    "Не конкурент / объект наблюдения": "#bdc3c7",
}


def short(n):
    return n.split(" (")[0].replace(" / diagrams.net", "").replace("AI-ассистенты общего назначения", "AI-ассистенты")


def save(fig, name):
    fig.savefig(os.path.join(OUT, name), dpi=150, bbox_inches="tight")
    plt.close(fig)
    print("saved", name)


OFFS = {"Eraser": (10, -14), "Lucidchart": (-20, 12), "Visual Paradigm Online": (8, -13), "draw.io": (-50, 10),
        "D2": (-22, 6), "Mermaid": (10, -4), "Microsoft Visio": (-20, -16), "Camunda Modeler": (-40, -16),
        "Sparx Enterprise Architect": (-60, -16), "Structurizr": (-30, -16), "dbdiagram.io": (-30, 10)}


def fig_map(A):
    fig, ax = plt.subplots(figsize=(9, 7.5))
    zones = [((0, 0), 5, 5, "#f4f6f7", "Периферия"),
             ((0, 3), 3, 2, "#eaf2f8", "Аудиторное пересечение"),
             ((3, 0), 2, 3, "#fef5e7", "Функциональный заменитель"),
             ((3, 3), 2, 2, "#fdebd0", "Близкая альтернатива"),
             ((4, 4), 1, 1, "#f5b7b1", "Ядро прямой конкуренции")]
    for (xy, w, h, c, lab) in zones:
        ax.add_patch(Rectangle(xy, w, h, color=c, zorder=0))
    ax.text(0.1, 1.6, "Периферия", color="#7f8c8d")
    ax.text(0.1, 4.8, "Аудиторное пересечение (Y ≥ 3)", color="#2471a3")
    ax.text(4.1, 1.6, "Функциональный\nзаменитель (X ≥ 3)", color="#b9770e")
    ax.text(3.05, 3.55, "Близкая\nальтернатива", color="#a04000")
    ax.text(4.05, 4.85, "Ядро (≥4; ≥4)", color="#922b21", fontsize=8)
    rng = np.random.default_rng(5)
    seen = {}
    for idx, r in enumerate(A):
        key = (r["X"], r["Y"]); k = seen.get(key, 0); seen[key] = k + 1
        ang = k * 2.1
        jx, jy = (0.16 * np.cos(ang), 0.16 * np.sin(ang)) if k else (0, 0)
        ax.scatter(r["X"] + jx, r["Y"] + jy, s=60 + r["N"] * 70, color=LVL_COL[r["final"]], edgecolor="k", zorder=3)
        ax.annotate(short(r["name"]), (r["X"] + jx, r["Y"] + jy), xytext=OFFS.get(short(r["name"]), (6, 5 if idx % 2 else -11)), textcoords="offset points", fontsize=7.5, zorder=4)
    ax.scatter([5], [5], marker="*", s=300, color="#1e8449", edgecolor="k", zorder=5)
    ax.annotate("NotaCode (по определению)", (5, 5), xytext=(-150, -2), textcoords="offset points", fontsize=8, color="#1e8449")
    ax.set_xlim(0.5, 5.2); ax.set_ylim(1.5, 5.2)
    ax.set_xlabel("X = среднее(C1 задача; C3 заменяемость)")
    ax.set_ylabel("Y = среднее(C2 аудитория; C7 география/язык)")
    ax.set_title("Карта конкурентного поля NotaCode (зоны листа 07_Карта; размер — взвешенный балл N)")
    ax.legend(handles=[Patch(color=c, label=l) for l, c in LVL_COL.items()], loc="upper left", bbox_to_anchor=(1.01, 1), fontsize=7, title="Итоговый уровень")
    save(fig, "fig_01_karta_konkurentnogo_polya.png")


def fig_bar(A):
    A = sorted(A, key=lambda r: r["N"])
    fig, ax = plt.subplots(figsize=(9, 7))
    y = np.arange(len(A))
    ax.barh(y, [r["N"] for r in A], color=[LVL_COL[r["final"]] for r in A], edgecolor="k")
    for i, r in enumerate(A):
        ax.text(r["N"] + 0.03, i, f'{r["N"]:.2f}'.replace(".", ",") + (" *" if r["corr"] else ""), va="center", fontsize=8)
    ax.set_yticks(y, [short(r["name"]) for r in A])
    for t, lab in [(1.6, "1,6"), (2.5, "2,5"), (3.4, "3,4"), (4.2, "4,2")]:
        ax.axvline(t, ls="--", color="grey", lw=0.8)
        ax.text(t, len(A) - 0.3, lab, ha="center", fontsize=8, color="grey")
    ax.set_xlim(0, 5)
    ax.set_xlabel("Взвешенный балл N = СУММПРОИЗВ(C1…C10; веса) / 100")
    ax.set_title("Взвешенные баллы близости (04_Оценка); * — экспертная корректировка уровня")
    ax.legend(handles=[Patch(color=c, label=l) for l, c in LVL_COL.items()], loc="lower right", fontsize=7)
    save(fig, "fig_02_vzveshennye_bally.png")


def fig_levitt(B):
    ids = D.LK_IDS
    M = np.full((len(ids), 5), np.nan)
    for i, cid in enumerate(ids):
        for lv in range(1, 6):
            v = [r["L"] for r in B["lev"] if r["id"] == cid and r["level"] == lv]
            if v:
                M[i, lv - 1] = np.mean(v)
    fig, ax = plt.subplots(figsize=(8, 6.5))
    im = ax.imshow(M, cmap="YlGn", vmin=0, vmax=5, aspect="auto")
    for i in range(M.shape[0]):
        for j in range(5):
            ax.text(j, i, "—" if np.isnan(M[i, j]) else f"{M[i, j]:.2f}".replace(".", ","), ha="center", va="center", fontsize=8)
    ax.set_xticks(range(5), ["1 Базовая\nвыгода", "2 Родовой", "3 Ожидаемый", "4 Расширенный", "5 Потенциальный"])
    ax.set_yticks(range(len(ids)), [f'{c["id"]} {short(c["name"])}' for c in D.LK_COMP])
    fig.colorbar(im, ax=ax, label="Оценка признака L = Наличие·((Сила+Понятность)/6)·Влияние")
    ax.set_title("Модель Левитта: оценка признаков по уровням товара (03_Левитт)")
    save(fig, "fig_03_levitt_heatmap.png")


def fig_kano_heat(B):
    ks = [k for k in B["kano"] if k["pres"] is not None]
    M = np.array([[1.0 if ch == "1" else 0.5 if ch == "p" else 0.0 for ch in k["pres"]] for k in ks])
    fig, ax = plt.subplots(figsize=(10, 9))
    ax.imshow(M, cmap="Blues", vmin=0, vmax=1.2, aspect="auto")
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            ax.text(j, i, {1.0: "✓", 0.5: "?", 0.0: ""}[M[i, j]], ha="center", va="center", fontsize=8)
    ax.set_xticks(range(12), [short(c["name"]) for c in D.LK_COMP], rotation=45, ha="right")
    ax.set_yticks(range(len(ks)), [f'{k["id"]} {k["feat"][:42]} [{k["cat"][:5]}.]' for k in ks], fontsize=7.5)
    ax.set_title("Модель Кано: присутствие требований у конкурентов (✓ — подтверждено, ? — проверить)")
    save(fig, "fig_04_kano_heatmap.png")


def fig_kano_dist(B):
    cols = {"Обязательное": "#c0392b", "Линейное": "#2e86c1", "Привлекательное": "#27ae60",
            "Безразличное": "#95a5a6", "Обратное": "#8e44ad", "Неясное": "#f1c40f"}
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(14, 6.5), gridspec_kw={"width_ratios": [1, 2], "wspace": 0.75})
    d = B["kano_dist"]
    a1.bar(range(6), [d[c] for c in D.KANO_CATS], color=[cols[c] for c in D.KANO_CATS], edgecolor="k")
    for i, c in enumerate(D.KANO_CATS):
        a1.text(i, d[c] + 0.1, str(d[c]), ha="center")
    a1.set_xticks(range(6), D.KANO_CATS, rotation=35, ha="right")
    a1.set_ylabel("Число требований (04_Кано)")
    a1.set_title(f"Распределение {sum(d.values())} требований по Кано")
    ks = sorted([k for k in B["kano"] if k["H"] is not None], key=lambda k: k["H"])
    a2.barh(range(len(ks)), [k["H"] for k in ks], color=[cols[k["cat"]] for k in ks], edgecolor="k")
    a2.set_yticks(range(len(ks)), [f'{k["id"]} {k["feat"][:38]}' for k in ks], fontsize=7)
    for t in (0.5, 0.6):
        a2.axvline(t, ls="--", color="grey", lw=0.8)
    a2.set_xlabel("Доля конкурентов с признаком (из 12); пороги 0,5 и 0,6")
    a2.set_title("Доли конкурентов по требованиям")
    a2.legend(handles=[Patch(color=cols[c], label=c) for c in D.KANO_CATS], fontsize=7, loc="lower right")
    save(fig, "fig_05_kano_raspredelenie.png")


def fig_comp(B):
    cs = sorted(B["comp"], key=lambda c: c["total"])
    fig, ax = plt.subplots(figsize=(8, 4.5))
    col = ["#1e8449" if c["total"] >= 4.2 else "#2e86c1" if c["total"] >= 3.5 else "#f39c12" if c["total"] >= 2.5 else "#c0392b" for c in cs]
    ax.barh(range(len(cs)), [c["total"] for c in cs], color=col, edgecolor="k")
    for i, c in enumerate(cs):
        ax.text(c["total"] + 0.03, i, f'{c["total"]:.2f}'.replace(".", ",") + " — " + c["interp"], va="center", fontsize=8)
    ax.set_yticks(range(len(cs)), [f'{c["id"]} {short(c["name"])}' for c in cs])
    for t in (2.5, 3.5, 4.2):
        ax.axvline(t, ls="--", color="grey", lw=0.8)
    ax.set_xlim(0, 6)
    ax.set_title("Итоговый взвешенный балл по 25 критериям P01–P25 (06_Сравнение)")
    save(fig, "fig_06_sravnenie_P01_P25.png")


def fig_bm(Cc):
    ms = sorted(Cc["mat"], key=lambda m: m["Q"])
    fig, ax = plt.subplots(figsize=(8, 4.5))
    y = np.arange(len(ms))
    ax.barh(y - 0.2, [m["N"] for m in ms], 0.4, label="Индекс силы N", color="#5dade2", edgecolor="k")
    ax.barh(y + 0.2, [m["Q"] for m in ms], 0.4, label="Итог с учётом надёжности Q", color="#1f618d", edgecolor="k")
    for i, m in enumerate(ms):
        ax.text(m["Q"] + 0.03, i + 0.2, f'{m["Q"]:.2f}'.replace(".", ","), va="center", fontsize=8)
    ax.set_yticks(y, [short(m["name"]) for m in ms])
    ax.set_xlim(0, 5.3)
    ax.legend(fontsize=8, loc="lower right")
    ax.set_title("Бизнес-модели конкурентов (09_Матрица): Q = 0,6·N + 0,25·Доказ. + 0,15·5·Полнота")
    save(fig, "fig_07_biznes_model.png")


if __name__ == "__main__":
    A, B, Cc = calc_a(), calc_b(), calc_c()
    fig_map(A); fig_bar(A); fig_levitt(B); fig_kano_heat(B); fig_kano_dist(B); fig_comp(B); fig_bm(Cc)

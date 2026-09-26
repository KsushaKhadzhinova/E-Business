# -*- coding: utf-8 -*-
"""Рис. 3. Тепловая карта барьеров входа (лист 07_Барьеры_входа): баллы 1/3/5 по 9 факторам
в трёх сценариях экспертных оценок и индекс барьеров с нормированными весами (сумма = 1,00)
и с исходными весами шаблона (сумма = 1,04).
Запуск: python fig_heatmap_barerov.py  -> ../Рисунки/fig_03_heatmap_barerov.png
"""
import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap, BoundaryNorm
from lr4_data import BARRIERS, OPT, BASE, PESS, W_SUM, barrier_index, barrier_level, FIG_DIR, INK, INK2, fmt

cols = ["Оптимистичный", "Базовый", "Пессимистичный"]
M = np.array([OPT, BASE, PESS]).T
rows = [f"{b[0]}\nвес {fmt(b[1])} → {fmt(b[1]/W_SUM,3)} · {b[3]}" for b in BARRIERS]

# последовательная шкала одного оттенка (синий), 3 ступени под баллы 1/3/5
cmap = ListedColormap(["#cde2fb", "#6da7ec", "#1c5cab"])
norm = BoundaryNorm([0, 2, 4, 6], cmap.N)

fig, ax = plt.subplots(figsize=(10.5, 8.6))
ax.imshow(M, cmap=cmap, norm=norm, aspect="auto")
for i in range(M.shape[0]):
    for j in range(M.shape[1]):
        v = M[i, j]
        ax.text(j, i, str(v), ha="center", va="center", fontsize=13, fontweight="bold",
                color="white" if v == 5 else INK)
ax.set_xticks(range(3))
ax.set_xticklabels(cols, fontsize=10.5, color=INK)
ax.set_yticks(range(len(rows)))
ax.set_yticklabels(rows, fontsize=9, color=INK)
ax.tick_params(length=0)
for s in ax.spines.values():
    s.set_visible(False)
ax.set_xticks(np.arange(-0.5, 3, 1), minor=True)
ax.set_yticks(np.arange(-0.5, len(rows), 1), minor=True)
ax.grid(which="minor", color="white", linewidth=2)
ax.tick_params(which="minor", length=0)

lines = []
for name, sc in zip(cols, [OPT, BASE, PESS]):
    a, b = barrier_index(sc, True), barrier_index(sc, False)
    lines.append(f"{name}: индекс = {fmt(a)} ({barrier_level(a)}); с весами 1,04 — {fmt(b)} ({barrier_level(b)})")
ax.set_title("NotaCode: баллы факторов барьеров входа\n(1 — низкое, 3 — среднее, 5 — высокое давление)",
             fontsize=12, color=INK, pad=12)
fig.text(0.02, 0.11, "\n".join(lines), fontsize=9.5, color=INK, va="top")
fig.text(0.02, 0.005, "«расчёт» — балл из формул шаблона по данным SimilarWeb; «эксперт» — данных нет, балл с листа "
         "Ввод_NotaCode (допущение);\nсценарии ±2 меняют только экспертные баллы.", fontsize=8.3, color=INK2)
fig.tight_layout(rect=[0, 0.13, 1, 1])
out = os.path.join(FIG_DIR, "fig_03_heatmap_barerov.png")
fig.savefig(out, dpi=160, facecolor="white")
print("saved", out)

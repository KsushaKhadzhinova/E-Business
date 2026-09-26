# -*- coding: utf-8 -*-
"""Рис. 1. Радар-диаграмма пяти сил Портера + цифровые усилители (шкала 1/3/5), NotaCode.
Запуск: python fig_radar_5_sil.py  -> ../Рисунки/fig_01_radar_5_sil.png
"""
import os
import numpy as np
import matplotlib.pyplot as plt
from lr4_data import FORCES, DIGITAL, FIG_DIR, C, INK, INK2, MUTED, fmt, interp_pressure


def radar(ax, items, color, title):
    labels = [a for a, _ in items]
    vals = [v for _, v in items]
    n = len(items)
    ang = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist()
    ang_c = ang + ang[:1]
    v_c = vals + vals[:1]
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.set_ylim(0, 5.4)
    ax.set_yticks([1, 3, 5])
    ax.set_yticklabels(["1", "3", "5"], color=MUTED, fontsize=9)
    ax.set_xticks(ang)
    ax.set_xticklabels(labels, fontsize=9.5, color=INK)
    ax.tick_params(axis="x", pad=14)
    ax.grid(color="#dddcd8", linewidth=0.8)
    ax.spines["polar"].set_color("#dddcd8")
    ax.plot(ang_c, v_c, color=color, linewidth=2)
    ax.fill(ang_c, v_c, color=color, alpha=0.18)
    ax.scatter(ang, vals, s=40, color=color, edgecolor="white", linewidth=1.5, zorder=3)
    for a, v in zip(ang, vals):
        ax.text(a, v - 0.55 if v == 5 else v + 0.55, str(v), ha="center", va="center", fontsize=10, color=INK, fontweight="bold")
    avg = sum(vals) / n
    ax.set_title(f"{title}\nсреднее = {fmt(avg)} ({interp_pressure(avg) if title.startswith('Пять') else 'средняя выраженность'})",
                 fontsize=11, color=INK, pad=58)


fig = plt.figure(figsize=(13, 6.8))
ax1 = fig.add_subplot(1, 2, 1, polar=True)
ax2 = fig.add_subplot(1, 2, 2, polar=True)
radar(ax1, FORCES, C[0], "Пять сил Портера + цифровые усилители")
radar(ax2, DIGITAL, C[1], "Цифровые усилители (6 факторов)")
fig.suptitle("NotaCode: оценка конкурентного давления (1 — низкое, 3 — среднее, 5 — высокое)",
             fontsize=13, color=INK, y=0.995)
fig.text(0.5, 0.01, "Шкала методики 5 сил (разд. 12). Интерпретация среднего: 1,0–2,0 низкое; 2,1–3,5 умеренное; "
         "3,6–5,0 высокое. Оценки — ЛР4, 25.09.2026.", ha="center", fontsize=8.5, color=INK2)
fig.tight_layout(rect=[0, 0.07, 1, 0.95])
out = os.path.join(FIG_DIR, "fig_01_radar_5_sil.png")
fig.savefig(out, dpi=160, facecolor="white")
print("saved", out)

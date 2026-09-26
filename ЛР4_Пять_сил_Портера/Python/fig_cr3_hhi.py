# -*- coding: utf-8 -*-
"""Рис. 2. Доли трафика конкурентов, CR3 и HHI (данные SimilarWeb из ЛР2-3 старой версии).
Три сценария: A — глобальный трафик; B — верхняя граница трафика из РФ; C — без app.diagrams.net.
Запуск: python fig_cr3_hhi.py  -> ../Рисунки/fig_02_cr3_hhi_doli_trafika.png
"""
import os
import matplotlib.pyplot as plt
from lr4_data import SCENARIOS, FIG_DIR, C, INK, INK2, MUTED, fmt

# цвет закреплён за доменом (не за рангом)
COLOR = {"app.diagrams.net (draw.io)": C[0], "eraser.io": C[1], "plantuml.com": C[2], "mermaidchart.com": C[3]}

fig, axes = plt.subplots(1, 3, figsize=(15, 5.4), sharex=True)
for ax, s in zip(axes, SCENARIOS):
    items = sorted(s["shares"], key=lambda t: t[1])
    names = [d for d, _ in items]
    vals = [x * 100 for _, x in items]
    ax.barh(names, vals, color=[COLOR[d] for d in names], height=0.6, edgecolor="white", linewidth=2)
    for i, v in enumerate(vals):
        ax.text(v + 1.5, i, fmt(v, 1) + " %", va="center", fontsize=9.5, color=INK)
    ax.set_xlim(0, 110)
    ax.set_xlabel("Доля релевантного трафика, %", color=INK2)
    ax.tick_params(colors=INK2)
    ax.grid(axis="x", color="#e6e5e1", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.set_title(s["name"], fontsize=10.5, color=INK)
    lvl = "высокая" if s["cr3"] >= 0.6 or s["hhi"] >= 2500 else "средняя"
    ax.text(0.98, 0.04, f"CR3 = {fmt(s['cr3']*100,1)} %\nCR5 = {fmt(s['cr5']*100,1)} %\nHHI = {round(s['hhi'])}\n"
            f"концентрация: {lvl}", transform=ax.transAxes, ha="right", va="bottom", fontsize=9.5, color=INK,
            bbox=dict(boxstyle="round,pad=0.4", facecolor="#f3f2ee", edgecolor="#dddcd8"))
fig.suptitle("Концентрация цифрового спроса: доли трафика, CR3 и HHI (пороги шаблона: CR3 ≥ 0,6; HHI ≥ 2500)",
             fontsize=12.5, color=INK)
fig.text(0.5, 0.01, "Источник: SimilarWeb, бесплатный веб-доступ, «последние 3 месяца», снято 17.09.2026 "
         "(plantuml.com, mermaidchart.com) и 23.09.2026 (eraser.io, app.diagrams.net).\nmermaidchart.com — "
         "верхняя граница «< 20 тыс.». Остальные 14 доменов — 🔲 ДОСНЯТЬ.".replace("🔲 ", ""),
         ha="center", fontsize=8.3, color=INK2)
fig.tight_layout(rect=[0, 0.07, 1, 0.94])
out = os.path.join(FIG_DIR, "fig_02_cr3_hhi_doli_trafika.png")
fig.savefig(out, dpi=160, facecolor="white")
print("saved", out)

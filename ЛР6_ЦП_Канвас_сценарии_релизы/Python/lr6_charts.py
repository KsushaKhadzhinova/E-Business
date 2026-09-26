# -*- coding: utf-8 -*-
"""ЛР6. Рисунки: ИПП, Канвас (9 блоков), индексы доказательности Канваса, цепочка бизнес-логики,
карта сценариев, карта навигации (Graphviz dot, иначе matplotlib), ИРП с порогами, диаграмма
релизов R0–R4 (Гант), вайрфреймы страниц без мокапов (лендинг, тарифы, модалка лимита).

Запуск: python -B lr6_charts.py   (сначала python -B lr6_calc.py не обязателен — расчёт повторяется)
Результат: ../Рисунки/*.png
"""
import datetime as dt
import os
import shutil
import subprocess
import textwrap

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle

import lr6_d1_cp as d1
import lr6_d2_canvas as d2
import lr6_d3_logic as d3
import lr6_d4_screens as d4
import lr6_d5_releases as d5
from lr6_common import FIG, REL_COLORS, ensure_dirs, setup_fonts

DOT_CANDIDATES = [r"E:\VisualDSL_Tools\Graphviz\bin\dot.exe", "dot"]


def save(fig, name):
    path = os.path.join(FIG, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print("OK", name)


# ---------- 1. ИПП ----------
def fig_ipp():
    ids = list(d1.IPP)
    vals = [round(sum(w * v for w, v in zip(d1.IPP_W, d1.IPP[i])), 2) for i in ids]
    fig, ax = plt.subplots(figsize=(10, 4.8))
    cols = ["#2e7d32" if v >= 4.0 else "#1565c0" if v >= 3.0 else "#9e9e9e" for v in vals]
    ax.barh([f"{i} {d1.PROBLEM_SHORT[i]}" for i in ids][::-1], vals[::-1], color=cols[::-1])
    for y, v in enumerate(vals[::-1]):
        ax.text(v + 0.03, y, f"{v:.2f}".replace(".", ","), va="center", fontsize=9)
    for x, lab in ((4.5, "4,50 основа"), (4.0, "4,00 основа после уточнения"), (3.0, "3,00 доп. ценность")):
        ax.axvline(x, ls="--", color="#c62828", lw=1)
        ax.text(x, len(ids) - 0.4, lab, color="#c62828", fontsize=8, ha="center")
    ax.set_xlim(0, 5)
    ax.set_xlabel("ИПП (шкала 1–5)")
    ax.set_title("Индекс приоритета проблем NotaCode (методика 6.1)")
    save(fig, "fig_01_ipp.png")


# ---------- 2. Канвас ----------
def fig_canvas():
    fig, ax = plt.subplots(figsize=(16, 9))
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis("off")
    boxes = {  # x, y, w, h
        "Ключевые партнеры": (0, 1.8, 2, 4.2), "Ключевые виды деятельности": (2, 3.9, 2, 2.1),
        "Ключевые ресурсы": (2, 1.8, 2, 2.1), "Ценностное предложение": (4, 1.8, 2, 4.2),
        "Отношения с клиентами": (6, 3.9, 2, 2.1), "Каналы": (6, 1.8, 2, 2.1),
        "Клиентские сегменты": (8, 1.8, 2, 4.2), "Структура затрат": (0, 0, 5, 1.8), "Потоки доходов": (5, 0, 5, 1.8),
    }
    evid = {}
    for b in d2.CANVAS_BLOCK_NAMES:
        st = d2.BLOCK_STATEMENTS[b]
        evid[b] = round(sum(s[1] for s in st) / (3 * len(st)) * 100)
    for name, (x, y, w, h) in boxes.items():
        e = evid[name]
        face = "#e8f5e9" if e >= 75 else "#fffde7" if e >= 50 else "#ffebee"
        ax.add_patch(Rectangle((x, y), w, h, facecolor=face, edgecolor="#37474f", lw=1.5))
        ax.text(x + 0.08, y + h - 0.12, name, fontsize=11, fontweight="bold", va="top")
        ax.text(x + w - 0.08, y + h - 0.12, f"{e} %", fontsize=9, va="top", ha="right", color="#37474f")
        lines = d2.CANVAS_PICTURE[name]
        ax.text(x + 0.08, y + h - 0.45, "\n".join(lines), fontsize=8.6, va="top", linespacing=1.35)
    ax.set_title("Бизнес-модель NotaCode по Канвас (цвет — индекс доказательности: зелёный ≥75 %, жёлтый 50–74 %, красный <50 %)", fontsize=12)
    save(fig, "fig_02_canvas.png")


def fig_canvas_evidence():
    names = d2.CANVAS_BLOCK_NAMES
    vals = [round(sum(s[1] for s in d2.BLOCK_STATEMENTS[b]) / (3 * len(d2.BLOCK_STATEMENTS[b])) * 100, 1) for b in names]
    fig, ax = plt.subplots(figsize=(10, 4.5))
    cols = ["#2e7d32" if v >= 75 else "#f9a825" if v >= 50 else "#c62828" for v in vals]
    ax.bar(range(9), vals, color=cols)
    ax.set_xticks(range(9))
    ax.set_xticklabels([textwrap.fill(n, 12) for n in names], fontsize=8)
    for i, v in enumerate(vals):
        ax.text(i, v + 1.5, f"{v:.1f}".replace(".", ","), ha="center", fontsize=8)
    ax.axhline(75, ls="--", color="#2e7d32", lw=1)
    ax.axhline(50, ls="--", color="#c62828", lw=1)
    ax.set_ylim(0, 105)
    ax.set_ylabel("Индекс доказательности, %")
    ok = sum(v >= 50 for v in vals)
    ax.set_title(f"Доказательность блоков Канваса; полнота = {ok}/9 = {ok / 9 * 100:.1f} %".replace(".", ","))
    save(fig, "fig_03_canvas_evidence.png")


# ---------- Graphviz ----------
def dot_bin():
    for c in DOT_CANDIDATES:
        if os.path.exists(c) or shutil.which(c):
            return c
    return None


def render_dot(src, name):
    b = dot_bin()
    if not b:
        return False
    path = os.path.join(FIG, name)
    r = subprocess.run([b, "-Tpng", "-Gdpi=130", "-o", path], input=src.encode("utf-8"), capture_output=True)
    ok = r.returncode == 0 and os.path.exists(path)
    print(("OK " if ok else "ERR ") + name, r.stderr.decode("utf-8", "ignore")[:200])
    return ok


def fig_logic_chain():
    src = """digraph G { rankdir=LR; node [shape=box, style="rounded,filled", fontname="Arial", fontsize=11, fillcolor="#e3f2fd"];
    edge [fontname="Arial", fontsize=9];
    seg [label="Сегмент\\nS-01 студенты\\n(платят S-02/S-03)"]; pr [label="Проблема P-001\\nручное построение +\\nнет проверки правил\\nИПП 4,35"];
    tov [label="Цифровой товар\\nNotaCode Web-IDE"]; mech [label="Механизм\\nDSL → разбор → правила\\nнотации → elkjs → SVG"];
    sc [label="Сценарии\\nS-003 первая диаграмма\\nS-004 IDEF0 с проверкой\\nS-005 сохранить/экспорт"];
    mon [label="Монетизация\\nлимиты Free →\\nPro 4 USD/мес", fillcolor="#fff3e0"]; op [label="Операционная способность\\nкаталог нотаций, CI 100 %,\\nfree tiers, AI-заглушка"];
    met [label="Метрики\\nактивация, Run OK,\\nD7, заявки Pro", fillcolor="#e8f5e9"];
    seg -> pr -> tov -> mech -> sc -> mon -> op -> met; sc -> met [style=dashed, label="события"]; }"""
    render_dot(src, "fig_04_business_logic.png")


def fig_scenario_map():
    stages = ["Привлечение", "Первичный интерес", "Регистрация", "Настройка", "Получение ценности", "Оплата", "Повторное использование", "Поддержка", "Удержание"]
    fig, ax = plt.subplots(figsize=(16, 5.5))
    ax.set_xlim(-0.5, len(stages) - 0.5)
    ax.set_ylim(-0.6, len(d3.SC_MAP) - 0.4)
    ax.set_xticks(range(len(stages)))
    ax.set_xticklabels([textwrap.fill(s, 12) for s in stages], fontsize=9)
    ax.set_yticks(range(len(d3.SC_MAP)))
    ax.set_yticklabels([f"{r[0]} {textwrap.fill(r[2], 22)}" for r in d3.SC_MAP][::-1], fontsize=9)
    ax.xaxis.tick_top()
    for yi, r in enumerate(d3.SC_MAP[::-1]):
        cells = r[3:12]
        crit = r[13]
        for xi, c in enumerate(cells):
            if not c:
                continue
            is_crit = stages[xi] == crit
            ax.add_patch(FancyBboxPatch((xi - 0.45, yi - 0.35), 0.9, 0.7, boxstyle="round,pad=0.02",
                                        facecolor="#ffcdd2" if is_crit else "#e3f2fd", edgecolor="#c62828" if is_crit else "#90a4ae"))
            ax.text(xi, yi, textwrap.fill(c, 14), ha="center", va="center", fontsize=7.5)
    ax.set_title("Карта ключевых сценариев NotaCode по этапам пути (красным — критическая точка)", pad=40)
    for s in ax.spines.values():
        s.set_visible(False)
    save(fig, "fig_05_scenario_map.png")


def fig_navigation():
    prio = {}
    for u in d4.UNITS:
        D, E, F, G, H = u[20]
        i = round((D * .30 + E * .30 + F * .15 + G * .15 - H * .10) * 20)
        prio[u[0]] = "P1" if i >= 80 else "P2" if i >= 60 else "P3"
    color = {"P1": "#c8e6c9", "P2": "#fff9c4", "P3": "#eceff1"}
    ecol = {"Основной": "#2e7d32", "Возврат": "#546e7a", "Ветвление": "#1565c0", "Переход после ошибки": "#c62828",
            "Переход после оплаты": "#ef6c00", "Системный переход": "#6a1b9a", "Внешний переход": "#000000"}
    lines = ['digraph N { rankdir=LR; splines=true; nodesep=0.35; ranksep=0.6; node [shape=box, style="rounded,filled", fontname="Arial", fontsize=10];',
             'edge [fontname="Arial", fontsize=8];',
             'subgraph cluster_site { label="Сайт"; style=dashed; fontname="Arial";']
    for u in d4.UNITS:
        if u[1] == "Страница сайта":
            lines.append(f'"{u[0]}" [label="{u[0]}\\n{u[2]}\\n{prio[u[0]]}", fillcolor="{color[prio[u[0]]]}"];')
    lines.append("}")
    lines.append('subgraph cluster_app { label="Приложение (Web IDE)"; style=dashed; fontname="Arial";')
    for u in d4.UNITS:
        if u[1] != "Страница сайта":
            name = textwrap.fill(u[2], 22).replace("\n", "\\n")
            lines.append(f'"{u[0]}" [label="{u[0]}\\n{name}\\n{prio[u[0]]}", fillcolor="{color[prio[u[0]]]}"];')
    lines.append("}")
    for n in d4.NAV:
        lines.append(f'"{n[0]}" -> "{n[1]}" [color="{ecol.get(n[6], "#333333")}", label="{n[2]}"];')
    lines.append('legend [shape=note, fillcolor="white", label="Цвет узла: зелёный P1, жёлтый P2, серый P3\\nРёбра: зелёные — основной, синие — ветвление,\\nкрасные — после ошибки, фиолетовые — системный,\\nоранжевые — после оплаты, серые — возврат"];')
    lines.append("}")
    src = "\n".join(lines)
    if not render_dot(src, "fig_06_navigation.png"):
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.text(0.5, 0.5, "Graphviz не найден — см. mermaid-исходник в СПИСОК_РИСУНКОВ.md", ha="center")
        ax.axis("off")
        save(fig, "fig_06_navigation.png")


# ---------- ИРП ----------
def fig_irp():
    units = d5.UNITS
    vals = [d5.irp(u[8]) for u in units]
    order = sorted(range(len(units)), key=lambda i: vals[i])
    fig, ax = plt.subplots(figsize=(11, 12))
    labels = [f"{units[i][0]} {textwrap.shorten(units[i][2], 52, placeholder='…')}" for i in order]
    cols = [REL_COLORS[units[i][10]] for i in order]
    ax.barh(labels, [vals[i] for i in order], color=cols)
    for y, i in enumerate(order):
        ax.text(vals[i] + 0.03, y, f"{vals[i]:.2f}".replace(".", ","), va="center", fontsize=7.5)
    for x, lab in ((3.5, "3,5 первый релиз"), (2.5, "2,5 кандидат R1/R2"), (1.5, "1,5 последующий")):
        ax.axvline(x, ls="--", color="#c62828", lw=1)
        ax.text(x, len(units) - 0.2, lab, color="#c62828", fontsize=8, ha="center")
    ax.tick_params(axis="y", labelsize=7.5)
    ax.set_xlim(0, 4.3)
    ax.set_xlabel("ИРП = (0,20·ЗП + 0,20·СЦ + 0,15·СС + 0,15·МН + 0,10·УД + 0,10·РА + 0,10·ДК) − (0,15·СЛ + 0,10·РЗ)")
    from matplotlib.patches import Patch
    ax.legend(handles=[Patch(color=c, label=k) for k, c in REL_COLORS.items() if k != "R0"], loc="lower right", title="Итоговый релиз")
    ax.set_title("Индекс релизного приоритета единиц разработки NotaCode")
    save(fig, "fig_07_irp.png")


# ---------- Гант релизов ----------
def fig_gantt():
    D = lambda s: dt.date.fromisoformat(s)
    items = [
        ("R0", "R0 Подготовительный: требования, архитектура, мокапы, skeleton, CI", "2026-09-21", "2026-10-04"),
        ("R1", "S1–S3: Gateway, JWT, грамматика DSL, оболочка IDE (RU-001, 002, 010)", "2026-09-28", "2026-10-18"),
        ("R1", "S4–S7: холст, нотации волны 1, сохранение, экспорт (RU-003, 004, 006…008, 011, 012, 015)", "2026-10-19", "2026-11-15"),
        ("R1", "S8–S9: история/diff, GitHub, лендинг, тарифы, аналитика, право (RU-013, 016…020)", "2026-11-16", "2026-11-29"),
        ("R1", "S10: стабилизация MVP, контрольный лист R1", "2026-11-30", "2026-12-06"),
        ("R0", "S11–S12: РПЗ, демо, защита; сбор данных R1", "2026-12-07", "2026-12-20"),
        ("R2", "R2 Удержание, оплата Pro (RU-005, 009, 014, 021…027) — допущение", "2027-01-11", "2027-02-28"),
        ("R3", "R3 Рост аудитории (RU-028…033) — допущение", "2027-03-01", "2027-05-31"),
        ("R4", "R4 Расширение продукта (RU-034…037) — допущение", "2027-06-01", "2027-09-30"),
    ]
    fig, ax = plt.subplots(figsize=(14, 5.5))
    for y, (r, lab, a, b) in enumerate(items[::-1]):
        ax.barh(y, (D(b) - D(a)).days + 1, left=mdates.date2num(D(a)), color=REL_COLORS[r], alpha=0.9)
        ax.text(mdates.date2num(D(b)) + 3, y, lab, va="center", fontsize=8.5)
    ax.axvline(mdates.date2num(D("2026-12-06")), color="#c62828", lw=2)
    ax.text(mdates.date2num(D("2026-12-06")), len(items) - 0.4, " MVP / R1 06.12.2026", color="#c62828", fontsize=9)
    ax.set_yticks([])
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%m.%Y")); ax.tick_params(axis="x", labelsize=8)
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.set_xlim(mdates.date2num(D("2026-09-15")), mdates.date2num(D("2027-12-31")))
    ax.grid(axis="x", alpha=0.3)
    ax.set_title("Карта релизов NotaCode R0–R4 на шкале спринтов agile-plan (S0–S12)")
    save(fig, "fig_08_releases_gantt.png")


def fig_release_counts():
    rels = ["R1", "R2", "R3", "R4", "Искл."]
    cnt = [sum(1 for u in d5.UNITS if u[10] == r) for r in rels]
    avg = [round(sum(d5.irp(u[8]) for u in d5.UNITS if u[10] == r) / max(1, c), 2) for r, c in zip(rels, cnt)]
    fig, ax = plt.subplots(figsize=(8, 4))
    ax.bar(rels, cnt, color=[REL_COLORS[r] for r in rels])
    for i, (c, a) in enumerate(zip(cnt, avg)):
        ax.text(i, c + 0.3, f"{c} ед.\nИРП ср. {a:.2f}".replace(".", ","), ha="center", fontsize=9)
    ax.set_ylim(0, max(cnt) + 4)
    ax.set_ylabel("Число единиц разработки")
    ax.set_title("Состав релизов: количество единиц и средний ИРП")
    save(fig, "fig_09_release_counts.png")


# ---------- Вайрфреймы ----------
def wire(ax, x, y, w, h, text="", fc="#eceff1", fs=9, bold=False):
    ax.add_patch(Rectangle((x, y), w, h, facecolor=fc, edgecolor="#546e7a"))
    if text:
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs, fontweight="bold" if bold else "normal")


def fig_wireframes():
    fig, ax = plt.subplots(figsize=(10, 7))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    wire(ax, 0, 9.3, 10, 0.7, "NotaCode   |   Нотации   Примеры   Тарифы   Документация   |   [Открыть редактор]", "#cfd8dc")
    wire(ax, 0.3, 6.3, 5.2, 2.7, "Диаграммы IDEF0, IDEF3, DFD, ERD, UML из текста —\nс проверкой правил нотации до сдачи\n\n[Открыть редактор без регистрации]", "#e3f2fd", 10, True)
    wire(ax, 5.8, 6.3, 3.9, 2.7, "Скриншот IDE:\nкод слева, диаграмма справа,\nProblems: «IDEF0-R03, строка 12»")
    for i, t in enumerate(["Проверка правил\nс номером строки", "Подсветка\nстрока ↔ элемент", "История версий\nи откат"]):
        wire(ax, 0.3 + i * 3.2, 4.3, 3.0, 1.6, t, "#f1f8e9")
    wire(ax, 0.3, 2.3, 9.4, 1.6, "Таблица: нотация | PlantUML | Mermaid | draw.io | NotaCode (IDEF0 ✓, DFD ✓, правила ✓)", "#fffde7")
    wire(ax, 0.3, 0.3, 9.4, 1.6, "Галерея примеров (карточки)   ·   Тарифы Free / Pro   ·   Политика конфиденциальности")
    ax.set_title("Вайрфрейм U-001 «Главная (лендинг)» — мокапа в прототипе нет")
    save(fig, "wf_01_landing.png")

    fig, ax = plt.subplots(figsize=(10, 6.5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    wire(ax, 0, 9.2, 10, 0.8, "NotaCode   ·   Тарифы", "#cfd8dc")
    wire(ax, 0.5, 2.4, 4.2, 6.4, "Free — 0 USD\n\n5 проектов в облаке\n20 файлов в проекте\n20 версий файла\nЭкспорт SVG / PNG / DSL\nGitHub (ручной push)\nAI: заглушка / свой ключ\n\n[Начать бесплатно]", "#f1f8e9", 10)
    wire(ax, 5.3, 2.4, 4.2, 6.4, "Pro — 4 USD/мес или 40 USD/год\n\nБез лимита проектов\nПолная история, diff, откат\nВсе форматы: BPMN XML, XMI,\nPNML, draw.io, PlantUML, SQL DDL\nАвтосинхронизация GitHub/Drive\n\nR1: [Хочу Pro — оставить заявку]\nR2: [Оплатить]", "#e3f2fd", 10, True)
    wire(ax, 0.5, 0.4, 9.0, 1.6, "FAQ: что будет с проектами после окончания Pro? (данные сохраняются, доступ только на чтение сверх лимита)")
    ax.set_title("Вайрфрейм U-004 «Тарифы Free/Pro + заявка» (цена — допущение)")
    save(fig, "wf_02_pricing.png")

    fig, ax = plt.subplots(figsize=(8, 5))
    ax.set_xlim(0, 10); ax.set_ylim(0, 10); ax.axis("off")
    wire(ax, 0, 0, 10, 10, "", "#90a4ae")
    wire(ax, 1.5, 1.5, 7, 7, "", "white")
    ax.text(5, 7.6, "Лимит Free: 5 проектов из 5", ha="center", fontsize=13, fontweight="bold")
    ax.text(5, 5.6, "Ваши проекты и версии сохранены.\nЧтобы создать 6-й проект, удалите старый\nили перейдите на Pro (без лимита проектов).", ha="center", fontsize=10)
    wire(ax, 2.2, 2.2, 2.6, 1.2, "Управлять проектами")
    wire(ax, 5.2, 2.2, 2.6, 1.2, "Посмотреть Pro →", "#e3f2fd", 10, True)
    ax.set_title("Вайрфрейм U-019 «Модалка лимита Free» (состояние «ограниченный доступ»)")
    save(fig, "wf_03_limit_modal.png")


def main():
    ensure_dirs()
    setup_fonts()
    fig_ipp()
    fig_canvas()
    fig_canvas_evidence()
    fig_logic_chain()
    fig_scenario_map()
    fig_navigation()
    fig_irp()
    fig_gantt()
    fig_release_counts()
    fig_wireframes()


if __name__ == "__main__":
    main()

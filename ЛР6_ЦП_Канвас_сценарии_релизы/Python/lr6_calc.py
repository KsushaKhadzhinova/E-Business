# -*- coding: utf-8 -*-
"""ЛР6. Расчёт всех индексов (формулы методик docx и формулы Excel-шаблонов) и выгрузка
данных для VBA (VBA/data/lr6_N.tsv, UTF-8).

Запуск:  python -B lr6_calc.py
Результат: out/values.json, out/tables.json, out/*.csv, ../VBA/data/lr6_1.tsv … lr6_5.tsv
"""
import csv
import json
import os

import lr6_d1_cp as d1
import lr6_d2_canvas as d2
import lr6_d3_logic as d3
import lr6_d4_screens as d4
import lr6_d5_releases as d5
from lr6_common import OUTDIR, VBA_DATA, ensure_dirs, md_table

V = {}   # ключевые значения для отчёта
T = {}   # markdown-таблицы для отчёта


def r2(x):
    return round(x + 1e-12, 2)


# ---------------- 6.1 ----------------
def calc_61():
    rows = []
    for p in d1.PROBLEMS:
        pid = p[0]
        e, f, g, h, i, j = p[4:10]
        tmpl = round((e * .18 + f * .22 + g * .18 + h * .20 + i * .14 + j * .08) * 20)
        prio = "Критический" if tmpl >= 80 else "Высокий" if tmpl >= 65 else "Средний" if tmpl >= 45 else "Низкий"
        s = d1.IPP[pid]
        ipp = r2(sum(w * v for w, v in zip(d1.IPP_W, s)))
        dec = ("Основа предложения" if ipp >= 4.5 else "Основа после уточнения сегмента" if ipp >= 4.0 else
               "Дополнительная ценность" if ipp >= 3.0 else "Отложить" if ipp >= 2.0 else "Не использовать")
        rows.append((pid, d1.PROBLEM_SHORT[pid], *s, ipp, dec, tmpl, prio))
    V["ipp"] = {r[0]: r[8] for r in rows}
    V["tmpl_problem"] = {r[0]: r[10] for r in rows}
    T["ipp"] = md_table(["ID", "Проблема", "Ч", "П", "Н", "Пл", "ЦР", "Р", "ИПП", "Решение (методика)", "Индекс шаблона", "Приоритет шаблона"], rows)

    comp = []
    for c in d1.COMPETITOR_SOLUTIONS:
        e, f, g, h, i = c[4:9]
        j = round((e * .25 + f * .20 + g * .20 + h * .15 + i * .20) * 20)
        lvl = "Сильное решение" if j >= 80 else "Рыночный стандарт" if j >= 60 else "Частичное решение" if j >= 40 else "Слабое решение"
        comp.append((c[0], c[2], c[3], e, f, g, h, i, j, lvl))
    T["competitors"] = md_table(["Компания", "Проблема", "Как решает", "Покр.", "Ясн.", "Полн.", "Дост.", "Док.", "Индекс силы", "Уровень"], comp)

    ik = []
    for name, sc in d1.IKCP.items():
        ik.append((name, *sc, r2(sum(sc) / 6)))
    V["ikcp_v1"], V["ikcp"] = ik[0][-1], ik[1][-1]
    T["ikcp"] = md_table(["Формулировка", *d1.IKCP_CRIT, "ИКЦП"], ik)
    T["ikcp_comment"] = md_table(["Критерий", "Балл v2", "Обоснование"], [(k, d1.IKCP["Итоговое ЦП v2 (VP-001)"][i], d1.IKCP_COMMENT[k]) for i, k in enumerate(d1.IKCP_CRIT)])
    ikst = r2(sum(w * v for w, v in zip(d1.IKST_W, d1.IKST)))
    V["ikst"] = ikst
    T["ikst"] = md_table(["Критерий", "Вес", "Балл", "Вклад", "Обоснование"],
                         [(k, d1.IKST_W[i], d1.IKST[i], r2(d1.IKST_W[i] * d1.IKST[i]), d1.IKST_COMMENT[k]) for i, k in enumerate(d1.IKST_CRIT)] + [("Итого ИКСТ", "", "", ikst, "")])
    T["market_standard"] = md_table(["Признак", "У прямых конкурентов", "NotaCode R1"], d1.MARKET_STANDARD)

    # товар, монетизация, БМ, связность (формулы шаблона)
    prod_k = {p[0]: round(sum(1 for x in p[4:9] if x not in ("", None)) / 5 * 100) for p in d1.PRODUCTS}
    mon = {}
    trows = []
    for m in d1.MONETIZATION:
        rev = m[5] * m[6] if m[5] != "" and m[6] != "" else ""
        pot = round((m[9] * .40 + m[10] * .30 + m[11] * .30) * 20)
        st = "Готово" if pot >= 80 else "Нужно дополнить" if pot >= 60 else "Черновик" if pot >= 40 else "Отклонить"
        mon[m[0]] = pot
        trows.append((m[0], m[2], m[4], m[5], m[6], rev, m[9], m[10], m[11], pot, st))
    V["revenue_month"] = sum(r[5] for r in trows if r[5] != "")
    T["monet"] = md_table(["ID", "Единица", "Периодичность", "Цена, USD", "Ед./мес", "Выручка, USD/мес", "Связь", "Простота", "Повтор.", "Потенциал", "Статус"], trows)
    bm = {b[0]: round(sum(1 for x in b[2:11] if x) / 9 * 100) for b in d1.BUSINESS_MODEL}
    chk = []
    for c in d1.CHECKS:
        pi = V["tmpl_problem"][c[1]]
        vp_k = V["tmpl_problem"][next(v[1] for v in d1.VALUE_PROPS if v[0] == c[2])]
        g = round(pi * .30 + vp_k * .15 + prod_k[c[3]] * .20 + mon[c[4]] * .20 + bm[c[5]] * .15)
        h = "Проходят проверку" if g >= 75 else "Требуется доработка" if g >= 55 else "Недостаточно данных"
        chk.append((c[0], c[1], c[2], c[3], c[4], c[5], g, h, c[8]))
    V["chk"] = {c[0]: c[6] for c in chk}
    T["chk61"] = md_table(["ID", "Проблема", "ЦП", "Товар", "Монет.", "БМ", "Индекс связности", "Решение", "Главный разрыв"], chk)
    T["layers"] = md_table(["Слой", "Содержание NotaCode"], d1.PRODUCT_LAYERS)
    T["levels"] = md_table(["Уровень", "Состав", "Монетизация"], d1.OFFER_LEVELS)
    T["monet_check"] = md_table(["Условие", "Выполнено", "Обоснование"], d1.MONET_CHECK)
    T["hyp"] = md_table(["ID", "Тип", "Гипотеза", "Как проверить"], d1.HYPOTHESES)
    T["evidence"] = md_table(["Утверждение", "Источники", "Тип основания", "Доказательность", "Что досдать"], d1.EVIDENCE_MAP)
    T["ready61"] = md_table(["Пункт контрольного листа", "Статус", "Комментарий"], d1.READINESS)
    T["sources"] = md_table(["ID", "Дата", "Тип", "Площадка", "Компания", "Факт", "Дост."], [(s[0], s[1], s[2], s[3], s[4], s[6], s[8]) for s in d1.SOURCES])


# ---------------- 6.2 ----------------
def calc_62():
    rows = []
    ok = 0
    for b in d2.CANVAS_BLOCK_NAMES:
        st = d2.BLOCK_STATEMENTS[b]
        s = sum(x[1] for x in st)
        mx = 3 * len(st)
        idx = round(s / mx * 100, 1)
        lvl = "высокая" if idx >= 75 else "средняя" if idx >= 50 else "низкая"
        ok += idx >= 50
        rows.append((b, len(st), s, mx, idx, lvl))
    V["canvas_evid"] = {r[0]: r[4] for r in rows}
    V["canvas_full"] = round(ok / 9 * 100, 1)
    V["canvas_ok_blocks"] = ok
    links_ok = sum(1 for l in d2.LINKS if l[1])
    V["canvas_links"] = round(links_ok / len(d2.LINKS) * 100, 1)
    V["canvas_links_ok"] = links_ok
    V["canvas_evid_avg"] = round(sum(r[4] for r in rows) / 9, 1)
    T["canvas_evid"] = md_table(["Блок", "Утверждений", "Сумма баллов", "Максимум", "Индекс, %", "Доказательность"], rows)
    det = []
    for b in d2.CANVAS_BLOCK_NAMES:
        for s in d2.BLOCK_STATEMENTS[b]:
            det.append((b, s[0], s[1], s[2]))
    T["canvas_statements"] = md_table(["Блок", "Утверждение", "Балл", "Основание"], det)
    T["canvas_links"] = md_table(["Связь", "Подтверждена", "Комментарий"], [(l[0], "Да" if l[1] else "Нет", l[2]) for l in d2.LINKS])
    # шаблон 02_Канвас: статус
    st = []
    for name, c in zip(d2.CANVAS_BLOCK_NAMES, d2.CANVAS):
        f, g, h = c[2], c[3], c[4]
        s = "Готово" if (f >= 4 and g >= 4 and h <= 2) else "Требует доработки" if (f < 3 or g < 3 or h >= 4) else "Проверить"
        st.append((name, c[0], f, g, h, s, c[5]))
    V["canvas_status"] = {s[0]: s[5] for s in st}
    T["canvas_blocks"] = md_table(["Блок", "Рабочая формулировка", "Дост.", "Полн.", "Риск", "Статус шаблона", "Что проверить"], st)
    V["canvas_dash"] = {
        "full": round(sum(c[3] for c in d2.CANVAS) / 9 / 5 * 100, 1),
        "evid": round(sum(c[2] for c in d2.CANVAS) / 9 / 5 * 100, 1),
        "risk": round((1 - sum(c[4] for c in d2.CANVAS) / 9 / 5) * 100, 1),
        "ready": round(sum(1 for s in st if s[5] == "Готово") / 9 * 100, 1),
        "findings": len(d2.FINDINGS),
        "checks": round(sum(1 for c in d2.CONSISTENCY if c[1] >= 4) / 12 * 100, 1),
    }
    T["canvas_checks"] = md_table(["№", "Проверка", "Оценка", "Результат", "Что доработать"],
                                  [(i + 1, c[0], c[1], "Подтверждено" if c[1] >= 4 else "Нужна проверка" if c[1] >= 3 else "Риск разрыва", c[2]) for i, c in enumerate(d2.CONSISTENCY)])
    seg = []
    for s in d2.SEGMENTS:
        j = round(sum(s[4:8]) / 4, 2)
        seg.append((s[0], s[4], s[5], s[6], s[7], j, "Приоритетный" if j >= 4 else "Перспективный" if j >= 3 else "Низкий приоритет"))
    T["segments"] = md_table(["Сегмент", "Боль", "Сигнал", "Платёжесп.", "Канал", "Индекс", "Статус"], seg)


# ---------------- 6.3 ----------------
def calc_63():
    sc = []
    for s in d3.SCENARIOS:
        n, o, p, q = s[13:17]
        idx = round(n * o * p * q / 20, 2)
        pr = "Критический" if idx >= 15 else "Высокий" if idx >= 9 else "Средний" if idx >= 5 else "Низкий" if idx >= 2 else "Отложить"
        m = d3.SC_METHOD[s[0]]
        mi = r2(m[0] * .25 + m[1] * .20 + m[2] * .15 + m[3] * .15 + m[4] * .10 + m[5] * .15 - m[6] * .10)
        sc.append((s[0], s[1], d3.SC_CLASS[s[0]], s[2], idx, pr, *m, mi, d3.SC_METRIC[s[0]]))
    V["sc_tmpl"] = {x[0]: x[4] for x in sc}
    V["sc_method"] = {x[0]: x[13] for x in sc}
    T["scenarios"] = md_table(["ID", "Сценарий", "Класс", "Роль", "Индекс шаблона", "Приоритет", "Ц", "М", "Ч", "Р", "К", "Д", "С", "Индекс методики", "Метрика"], sc)
    T["sc_short"] = md_table(["ID", "Сценарий", "Класс (9 классов методики)", "Роль", "Связь с Канвас", "Метрика"],
                             [(s[0], s[1], d3.SC_CLASS[s[0]], s[2], d3.SC_CANVAS[s[0]], d3.SC_METRIC[s[0]]) for s in d3.SCENARIOS])
    req = []
    for r in d3.REQS:
        i, j, k, l = r[8:12]
        m = round((i + j) * l / k, 2)
        n = "Критический" if m >= 16 else "Высокий" if m >= 10 else "Средний" if m >= 6 else "Низкий" if m >= 3 else "Отложить"
        o = "Да" if n in ("Критический", "Высокий") else "Кандидат" if n == "Средний" else "Нет"
        req.append((r[0], r[1], r[2], r[6], r[7], i, j, k, l, m, n, o))
    V["mvp_reqs"] = sum(1 for r in req if r[11] == "Да")
    T["reqs"] = md_table(["ID", "Требование", "Сценарий", "Кано", "Левитт", "Ц", "М", "С", "Д", "Индекс", "Приоритет", "MVP"], req)
    pr = []
    for p in d3.PRIOR:
        f, g, h, i, j = p[5:10]
        k = round((f + g + h) * ((6 - j) / 5) / i, 2)
        l = ("Включить в первую версию" if k >= 6 else "Включить после проверки" if k >= 4 else "Проверить дополнительно" if k >= 2.5 else "Отложить" if k >= 1.5 else "Исключить")
        pr.append((p[0], p[1], p[3], f, g, h, i, j, k, l))
    T["prior63"] = md_table(["ID", "Инициатива", "Сценарий", "Ц", "Б", "Д", "С", "Р", "Индекс", "Решение"], pr)
    bank_use = sum(1 for b in d3.BANK if b[6] * b[7] * (6 - b[8]) / 20 >= 3)
    checks = [
        ("Выводов анализа ≥ 5", len(d3.BANK), 5), ("Сильных выводов (сила 3–4) ≥ 3", sum(1 for b in d3.BANK if b[6] >= 3), 3),
        ("Элементов бизнес-логики ≥ 3", len(d3.LOGIC), 3), ("Ролей ≥ 2", len(d3.ROLES), 2), ("Сценариев ≥ 5", len(d3.SCENARIOS), 5),
        ("Монетизационных сценариев ≥ 1", sum(1 for s in d3.SCENARIOS if s[4] == "Монетизационный"), 1),
        ("Функциональных требований ≥ 5", len(d3.REQS), 5), ("Требований MVP ≥ 3", V["mvp_reqs"], 3), ("Инициатив ≥ 5", len(d3.PRIOR), 5),
    ]
    passed = sum(1 for c in checks if c[1] >= c[2])
    checks.append(("Всего пройдено точек ≥ 7", passed, 7))
    passed_all = passed + (1 if passed >= 7 else 0)
    V["logic_ready"] = passed_all * 10
    V["bank_use"] = bank_use
    T["logic_checks"] = md_table(["Контрольная точка", "Количество", "Порог", "Статус"], [(c[0], c[1], c[2], "Пройдено" if c[1] >= c[2] else "Недостаточно") for c in checks])
    T["roles"] = md_table(["ID", "Роль", "Задача", "Готовность платить", "Частота", "Ценность сегмента"], [(r[0], r[1], r[2], r[6], r[7], r[6] * r[7]) for r in d3.ROLES])
    T["logic"] = md_table(["ID", "Элемент", "Тип", "Товар/функция", "Механизм", "Монетизация", "Д", "В", "Индекс надёжности"],
                          [(l[0], l[1], l[2], l[5], l[6], l[7], l[9], l[10], round(l[9] * l[10] / 20, 2)) for l in d3.LOGIC])
    T["bank63"] = md_table(["ID", "Направление", "Вывод", "Сила", "Знач.", "Риск", "Индекс опоры"], [(b[0], b[1], b[3], b[6], b[7], b[8], round(b[6] * b[7] * (6 - b[8]) / 20, 2)) for b in d3.BANK])
    for sid, pas in d3.PASSPORTS.items():
        T["pass_" + sid] = md_table(["Поле", "Значение"], list(pas.items()))
    hdr = ["№", "Действие пользователя", "Действие системы", "Ценность шага", "Данные", "Риск", "Метрика"]
    T["steps_S-004"] = md_table(hdr, d3.S004_STEPS)
    T["steps_S-007"] = md_table(hdr, d3.S007_STEPS)


# ---------------- 6.4 ----------------
def calc_64():
    units = d4.UNITS
    scen_prio = {s[0]: max(0, min(100, round((s[9] * .45 + s[10] * .35 - s[11] * .20) * 20))) for s in d4.SCEN}
    cov = {s[0]: sum(1 for u in units if u[4] == s[0]) for s in d4.SCEN}
    V["coverage"] = round(sum(1 for v in cov.values() if v > 0) / len(cov), 2)
    pr = []
    for u in units:
        D, E, F, G, H = u[20]
        i = max(0, min(100, round((D * .30 + E * .30 + F * .15 + G * .15 - H * .10) * 20)))
        cls = "P1" if i >= 80 else "P2" if i >= 60 else "P3" if i >= 40 else "Не включать"
        compl = round(sum(1 for x in u[3:16] if x not in ("", None)) / 13, 2)
        pr.append((u[0], u[1], u[2], u[4], u[5], scen_prio[u[4]], D, E, F, G, H, i, cls, compl, u[19]))
    V["p1"] = [p[0] for p in pr if p[12] == "P1"]
    V["units"] = len(units)
    V["unit_prio"] = {p[0]: (p[11], p[12]) for p in pr}
    T["units"] = md_table(["ID", "Тип", "Название", "Сценарий", "Роль", "Приор. сценария", "Ц", "Б", "Ч", "Р", "С", "Балл", "Класс", "Полнота", "Мокап"], pr)
    T["units_short"] = md_table(["ID", "Тип", "Название", "Сценарий", "Роль", "Этап", "Ключевое действие", "Класс"],
                                [(u[0], u[1], u[2], u[4], u[5], u[6], u[8], V["unit_prio"][u[0]][1]) for u in units])
    st = []
    for s in d4.STATES:
        st.append((s[0], *s[1:9], round(sum(1 for x in s[1:9] if x) / 8, 2)))
    V["states_full"] = round(sum(x[-1] for x in st) / len(st), 2)
    T["states"] = md_table(["Unit", "Нормальное", "Пустое", "Загрузка", "Ошибка валидации", "Ошибка сети", "Нет прав", "Успех", "Сообщение", "Полнота"], st)
    V["passport_full"] = 1.0
    V["nav_count"] = len(d4.NAV)
    T["nav"] = md_table(["ID", "Откуда", "Куда", "Событие", "Условие", "Тип перехода", "Риск разрыва"], [(f"N-{i + 1:03d}", n[0], n[1], n[2], n[3], n[6], n[7]) for i, n in enumerate(d4.NAV)])
    T["passports64"] = md_table(["Unit", "Название", "Назначение", "Действие", "Блоки", "Метрика", "Риск"],
                                [(r[0], r[1], r[5], r[7], r[8], r[11], r[12]) for r in d4.PASS_ROWS])
    V["monet_units"] = sum(1 for u in units if u[15] != "Не связана напрямую")
    V["source_share"] = round(sum(1 for u in units if u[9]) / len(units), 2)
    V["integral64"] = round((V["coverage"] + 1 + V["passport_full"] + V["states_full"]) / 4, 2)


# ---------------- 6.5 ----------------
def calc_65():
    rows = []
    for u in d5.UNITS:
        i = d5.irp(u[8])
        g = d5.template_scores(u)
        rows.append((u[0], u[2], u[6], u[7], *u[8], i, d5.irp_decision(i), d5.template_q(g), d5.template_r(g), d5.template_s(g), u[10], u[11]))
    V["irp"] = {r[0]: r[13] for r in rows}
    V["rel_count"] = {k: sum(1 for u in d5.UNITS if u[10] == k) for k in ("R1", "R2", "R3", "R4", "Искл.")}
    T["irp"] = md_table(["ID", "Единица", "Эпик", "MoSCoW", "ЗП", "СЦ", "СС", "МН", "УД", "РА", "ДК", "СЛ", "РЗ", "ИРП", "Решение по ИРП", "Балл R1 (шабл.)", "Балл след. (шабл.)", "Рекоменд. шаблона", "Релиз (итог)", "Обоснование"], rows)
    T["irp_short"] = md_table(["ID", "Единица", "ИРП", "Решение по ИРП", "Шаблон", "Итог"], [(r[0], r[1], r[13], r[14], r[17], r[18]) for r in rows])
    diff = [(r[0], r[1], r[13], r[14], r[17], r[18], r[19]) for r in rows
            if (r[18] == "R1") != (r[14] == "Первый релиз") or (r[18] == "R1") != (r[17] == "Первый релиз")]
    T["irp_diff"] = md_table(["ID", "Единица", "ИРП", "По ИРП", "По шаблону", "Итог", "Почему"], diff)
    V["irp_diff_n"] = len(diff)
    r1 = [u for u in d5.UNITS if u[10] == "R1"]
    V["r1_avg_irp"] = round(sum(V["irp"][u[0]] for u in r1) / len(r1), 2)
    q1 = [d5.template_q(d5.template_scores(u)) for u in r1]
    V["r1_avg_q"] = round(sum(q1) / len(q1), 2)
    V["value_check_avg"] = round(sum(v[2] for v in d5.VALUE_CHECK) / len(d5.VALUE_CHECK), 2)
    T["risks"] = md_table(["ID", "Единица", "Риск", "Тип", "В", "Вл", "Индекс", "Митигирующее действие", "Релиз"], [(r[0], r[1], r[2], r[3], r[4], r[5], r[4] * r[5], r[7], r[8]) for r in d5.RISKS])
    T["roadmap"] = md_table(["Релиз", "Цель", "Состав", "Метрики", "Критерий перехода", "Даты"], [(r[0], r[1], r[2], r[4], r[5], f"{r[9]} – {r[10]}") for r in d5.ROADMAP])
    T["value_check"] = md_table(["Блок", "Вопрос", "Оценка", "Доказательство", "Решение"], [(v[0], v[1], v[2], v[3], v[5]) for v in d5.VALUE_CHECK])
    T["first"] = md_table(["ID", "Единица", "Гипотеза", "Минимальная реализация", "Критерий готовности", "Метрика"], [(f[0], f[1], f[3], f[4], f[5], f[6]) for f in d5.FIRST])
    T["next"] = md_table(["ID", "Релиз", "Единица", "Основание переноса", "Условие включения", "Метрика"], [(n[0], n[1], n[2], n[3], n[4], n[6]) for n in d5.NEXT])
    T["protocol"] = md_table(["Поле протокола", "Значение"], d5.PROTOCOL)
    with open(os.path.join(OUTDIR, "irp.csv"), "w", encoding="utf-8-sig", newline="") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(["ID", "Единица", "ИРП", "Решение", "Релиз"])
        for r in rows:
            w.writerow([r[0], r[1], str(r[13]).replace(".", ","), r[14], r[18]])


# ---------------- TSV для VBA ----------------
def tsv_value(v):
    if v is None:
        return "~"
    if isinstance(v, float):
        return repr(v)
    return str(v).replace("\t", " ").replace("\n", " ")


def export_tsv():
    mods = {"lr6_1": d1, "lr6_2": d2, "lr6_3": d3, "lr6_4": d4, "lr6_5": d5}
    for key, mod in mods.items():
        lines = ["# LR6 NotaCode data; R<TAB>row<TAB>col<TAB>values…; ~ = keep template formula; empty = clear"]
        for sh in mod.SHEETS:
            lines.append("#SHEET\t" + sh["sheet"])
            for i, row in enumerate(sh["rows"]):
                lines.append("\t".join(["R", str(sh["row"] + i), sh["col"]] + [tsv_value(v) for v in row]))
        with open(os.path.join(VBA_DATA, key + ".tsv"), "w", encoding="utf-8") as fh:
            fh.write("\n".join(lines) + "\n")


def main():
    ensure_dirs()
    calc_61()
    calc_62()
    calc_63()
    calc_64()
    calc_65()
    export_tsv()
    with open(os.path.join(OUTDIR, "values.json"), "w", encoding="utf-8") as fh:
        json.dump(V, fh, ensure_ascii=False, indent=1)
    with open(os.path.join(OUTDIR, "tables.json"), "w", encoding="utf-8") as fh:
        json.dump(T, fh, ensure_ascii=False, indent=1)
    print("ИПП:", V["ipp"])
    print("ИКЦП v1/v2:", V["ikcp_v1"], V["ikcp"], "ИКСТ:", V["ikst"])
    print("Канвас доказательность:", V["canvas_evid"], "полнота", V["canvas_full"], "связность", V["canvas_links"])
    print("Связность 6.1:", V["chk"], "выручка/мес:", V["revenue_month"])
    print("Сценарии шаблон:", V["sc_tmpl"])
    print("Сценарии методика:", V["sc_method"])
    print("MVP-требований:", V["mvp_reqs"], "готовность 6.3:", V["logic_ready"])
    print("Покрытие:", V["coverage"], "P1:", V["p1"], "состояния:", V["states_full"], "интегр.:", V["integral64"])
    print("ИРП:", V["irp"])
    print("Релизы:", V["rel_count"], "расхождений:", V["irp_diff_n"], "ср. ИРП R1:", V["r1_avg_irp"], "ср. балл R1:", V["r1_avg_q"])


if __name__ == "__main__":
    main()

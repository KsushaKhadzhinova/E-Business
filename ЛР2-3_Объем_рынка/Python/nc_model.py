# -*- coding: utf-8 -*-
"""
ЛР2-3. Оценка объёма рынка NotaCode — единая расчётная модель.

Модуль повторяет формулы всех семи Excel-шаблонов из папки «Материалы»
(на тех же входных данных, что вносят VBA-модули из папки VBA), а также
считает два расчёта, которых в шаблонах нет:
  * «сверху вниз» от ВВП Беларуси (денежная база);
  * расчёт по методике ЛР3 «Оценка рынка на основе данных Similarweb»
    (V_sample → V_geo → V_total, CR3/CR5, HHI, TAM/SAM/SOM цифрового внимания).

Запуск:  python nc_model.py      — печатает все таблицы и сохраняет results_notacode.json
Импорт:  from nc_model import compute_all

Пометки:
  (реальные) — данные из архива DiagramCode с датой снятия;
  (допущение) — экспертное допущение, подлежит проверке;
  ДОСНЯТЬ — данные, которые студентка снимает сама (в модели = None/0).
"""
from __future__ import annotations

import json
from pathlib import Path

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------
# 1. Общие параметры
# ---------------------------------------------------------------------------
RATE_BYN_PER_USD = 3.00          # (допущение) курс на 25.09.2026, уточнить по nbrb.by
RATE_DATE = "25.09.2026"
PRO_USD_MONTH = 4.0              # цена Pro, концепция NotaCode
PRO_USD_YEAR = 40.0
PRO_BYN_MONTH = PRO_USD_MONTH * RATE_BYN_PER_USD          # 12 BYN
SCEN = ("Осторожный", "Базовый", "Оптимистичный")

# Годовой чек платящего пользователя, BYN (допущение о структуре оплаты):
#   осторожный — 6 мес. помесячно (семестр) = 6 × 12 = 72;
#   базовый    — годовой тариф 40 USD = 120;
#   оптимист.  — 12 мес. помесячно 48 USD = 144.
CHECK = (6 * PRO_BYN_MONTH, PRO_USD_YEAR * RATE_BYN_PER_USD, 12 * PRO_BYN_MONTH)

CONV_VISIT_REG = (0.02, 0.04, 0.06)   # визит → регистрация Free (допущение, SaaS-ориентир)
CONV_REG_PRO = (0.02, 0.03, 0.05)     # Free → Pro (KPI-05 NotaCode = 3 %, диапазон 3–5 % в концепции)
REACH = (0.01, 0.03, 0.07)            # достижимая доля нового бизнеса (К5 тема1.docx: 0,5–1 / 1–3 / 3–7 %)
COVERAGE = (0.80, 0.60, 0.45)         # коэффициент охвата выборки SW (осторожный = плотная выборка → меньший рынок)
COMMERCIAL = (0.20, 0.35, 0.50)       # коммерчески релевантная доля трафика (ЛР3, разд. 16)
COMMERCIAL_BASE = 0.35

# ---------------------------------------------------------------------------
# 2. Реальные данные (архив DiagramCode, lab2-3-market-size.md)
# ---------------------------------------------------------------------------
GDP_BY_USD = 78.59e9          # ВВП РБ 2024, macrotrends.net, снято 17.09.2026
IT_SHARE_GDP = 0.061          # доля ИТ в ВВП, I кв. 2025, opencompanyinbelarus.com
HTP_COMPANIES = 1000          # «более 1000» резидентов ПВТ, там же
IT_GRADUATES = 7000           # выпускников ИТ-специальностей в год, belarus.by
UNIVERSITIES = 21             # вузов, готовящих ИТ-специалистов, belarus.by

# Similarweb, бесплатный доступ: plantuml/mermaidchart — 17.09.2026, eraser/drawio — 23.09.2026.
# visits — Total Visits (последний месяц окна «последние 3 месяца»); geo — доля РФ/РБ;
# rel — доля категории «формальные нотации» (допущение); nondirect — доля не-прямых каналов;
# quality — коэффициент вовлечённости (допущение по bounce rate).
SW = [
    dict(domain="plantuml.com", name="PlantUML", kind="прямой (DSL-рендерер)", visits=507_000,
         geo=0.0907, geo_note="Россия 9,07 % (реальные)", rel=1.0, nondirect=0.50, nondirect_note="ДОСНЯТЬ каналы",
         quality=0.8, bounce=None, pages=None, duration=None, direct=None),
    dict(domain="eraser.io", name="Eraser", kind="прямой (DSL + AI)", visits=678_500,
         geo=0.01, geo_note="РФ/РБ вне топ-5 (<2,88 %), допущение 1 %", rel=0.5, nondirect=1 - 0.4905,
         nondirect_note="Direct 49,05 % (реальные)", quality=0.8, bounce=0.4373, pages=3.57, duration="00:01:39",
         direct=0.4905),
    dict(domain="app.diagrams.net", name="draw.io", kind="заменитель (GUI-редактор)", visits=7_800_000,
         geo=0.01, geo_note="РФ/РБ вне топ-5 (<4,35 %), допущение 1 %", rel=0.2, nondirect=1 - 0.6865,
         nondirect_note="Direct 68,65 % (реальные)", quality=0.7, bounce=0.555, pages=2.98, duration=None,
         direct=0.6865),
    dict(domain="mermaidchart.com", name="Mermaid Chart", kind="прямой (платный Mermaid)", visits=20_000,
         geo=0.01, geo_note="топ: VN/ID/US, допущение 1 %", rel=0.7, nondirect=0.50, nondirect_note="ДОСНЯТЬ каналы",
         quality=0.7, bounce=None, pages=None, duration=None, direct=None),
]

# Google Trends (реальные, Worldwide, 1-я неделя месяца, снято 17.09.2026) — 2024-01…2024-12
GT_MONTHS = [f"2024-{m:02d}" for m in range(1, 13)]
GT_PLANTUML = [37, 35, 46, 48, 47, 49, 40, 37, 39, 50, 51, 53]
GT_TEXT2DIAGRAM = [7, 9, 9, 11, 10, 6, 5, 6, 8, 10, 11, 12]

# Сегменты (строки шаблона «Потенциальные клиенты» 02_Сегменты, 4–7).
# D — всего; E — доля в границах; G — потребность; H — онлайн; I — платёжеспособные (готовые платить Pro);
# K — годовой чек; L — покупок/год.
SEGMENTS = [
    dict(code="S-01", name="Студенты ИТ-специальностей РБ", type="B2C",
         src="belarus.by: ~7000 выпускников/год × 4 курса (допущение)",
         D=IT_GRADUATES * 4, E=0.7, G=0.8, H=0.9, I=0.03, K=CHECK[0], L=1, rel="Средняя"),
    dict(code="S-02", name="Бизнес- и системные аналитики", type="B2B",
         src="ПВТ >1000 компаний × 2 аналитика (допущение)",
         D=HTP_COMPANIES * 2, E=0.8, G=0.6, H=0.9, I=0.10, K=CHECK[1], L=1, rel="Низкая"),
    dict(code="S-03", name="Архитекторы ПО и разработчики", type="B2B",
         src="ПВТ >1000 компаний × 3 специалиста (допущение)",
         D=HTP_COMPANIES * 3, E=0.6, G=0.5, H=0.9, I=0.08, K=CHECK[1], L=1, rel="Низкая"),
    dict(code="S-04", name="Преподаватели техн. дисциплин", type="B2C",
         src="21 вуз (belarus.by) × 15 преподавателей (допущение)",
         D=UNIVERSITIES * 15, E=0.8, G=0.7, H=0.9, I=0.10, K=CHECK[1], L=1, rel="Низкая"),
]

# Собственные каналы NotaCode для цифровой воронки (план 1-го года, допущения)
CHANNELS = [
    ("Органический поиск", "Яндекс/Google по ядру (Вордстат — ДОСНЯТЬ)", 30_000, 0.05, 0.7),
    ("Платный поиск", "Яндекс Директ, тестовая кампания", 10_000, 0.04, 0.8),
    ("Соцсети и Telegram", "ИТ-сообщества, вузовские чаты", 40_000, 0.01, 0.5),
    ("Реферальные площадки", "GitHub README, Хабр, каталоги", 5_000, 0.05, 0.7),
    ("Прямой трафик", "ссылки из методичек, закладки", 300, 1.0, 0.8),
    ("Каталоги инструментов", "AlternativeTo, Product Hunt", 3_000, 0.03, 0.6),
    ("Прочие (рекомендации преподавателей)", "экспертная оценка", 2_000, 0.05, 0.9),
]
CHANNEL_REALISATION = (0.6, 0.8, 1.0)   # доля реализации плана охвата («доля достижимого трафика»)
FUNNEL_REPEATS = (1.0, 1.0, 1.2)        # продления Pro

# «Сверху вниз» от ВВП (коэффициенты сужения — допущения)
TD_DEVTOOLS = (0.015, 0.02, 0.025)      # доля оборота ИТ-сектора на инструменты разработки
TD_MODELING = (0.04, 0.05, 0.06)        # доля инструментов моделирования и диаграмм
TD_DAC = (0.05, 0.10, 0.15)             # доля web / diagram-as-code сегмента
TD_NOTATIONS = (0.5, 0.6, 0.7)          # доля нотаций, которые покрывает NotaCode

# Платёжеспособность (допущения, 🔲 проверить по Белстату / опросу)
PAY = [  # сегмент, бюджет/мес BYN, часы на диаграммы в мес., ставка BYN/ч, доля экономии времени
    ("S-01 Студенты", 15, 20, 5, 0.4),
    ("S-02 Аналитики", 60, 10, 25, 0.4),
    ("S-03 Архитекторы", 60, 8, 35, 0.3),
    ("S-04 Преподаватели", 20, 6, 15, 0.3),
]
PRO_COST_BYN_MONTH = 3.0   # себестоимость платящего пользователя (хостинг, AI-квота), допущение

# Цены конкурентов (реальные, eraser.io, снято 17.09.2026), USD/участник/мес.
ERASER = dict(starter_year=15, starter_month=20, business_year=45, business_month=60)

# Сценарные множители шаблона «Потенциальные клиенты» (оставлены по шаблону)
PK_MULT = [(0.75, 0.8, 0.75, 0.8, 0.8), (1, 1, 1, 1, 1), (1.2, 1.15, 1.2, 1.15, 1.1)]
# Сценарные коэффициенты шаблона «Платёжеспособность» (оставлены по шаблону)
PAY_MULT = [(0.7, 0.6, 0.75, 0.8), (1, 1, 1, 1), (1.2, 1.25, 1.15, 1.2)]


def _tri(f):
    return [f(i) for i in range(3)]


# ---------------------------------------------------------------------------
# 3. Расчёты по шаблонам
# ---------------------------------------------------------------------------
def potential_clients():
    """shablon_..._kolichestvo_potencialnyh_klientov.xlsx: 02_Сегменты, 03_Сценарии, 04_Расчет."""
    rows = []
    for s in SEGMENTS:
        F = s["D"] * s["E"]
        J = F * s["G"] * s["H"] * s["I"]
        M = J * s["K"] * s["L"]
        O = M * REACH[1]
        rows.append(dict(s, F=F, J=J, M=M, O=O))
    tam_base = sum(s["D"] * s["K"] * s["L"] for s in SEGMENTS)
    sam_base = sum(r["M"] for r in rows)
    scen = []
    for i, m in enumerate(PK_MULT):
        tam = tam_base * m[3] * m[4]
        sam = sam_base * m[0] * m[1] * m[2] * m[3] * m[4]
        scen.append(dict(TAM=tam, SAM=sam, SOM=sam * REACH[i]))
    return dict(rows=rows, TAM=tam_base, SAM=sam_base, SOM=sum(r["O"] for r in rows),
                clients=sum(s["D"] for s in SEGMENTS), in_bounds=sum(r["F"] for r in rows),
                addressable=sum(r["J"] for r in rows), scen=scen)


def sw_rows():
    """Строки Similarweb в разных шаблонах (одни и те же входные данные)."""
    out = []
    for d in SW:
        geo_v = d["visits"] * d["geo"]
        rel_v = geo_v * d["rel"]
        out.append(dict(d, geo_visits=geo_v, rel_visits=rel_v))
    return out


def sw_competitor_checks():
    """Проверочные листы SW в шаблонах воронки (03), потенц. клиентов (05) и платёжеспособности (07)."""
    rows = sw_rows()
    conv_pay = CONV_VISIT_REG[1] * CONV_REG_PRO[1]          # 0,0012
    # воронка 03_Конкуренты_SW: F=C*D*E, I=F*G*H; C12=C11/B13 (исправлено), C14=C13/B13*12
    f_sum = sum(r["rel_visits"] for r in rows)
    i_sum = sum(r["rel_visits"] * conv_pay * CHECK[1] for r in rows)
    voronka = dict(F_sum=f_sum, full_traffic=f_sum / COVERAGE[1], revenue_month=i_sum,
                   revenue_year=i_sum / COVERAGE[1] * 12,
                   revenue_year_bug=i_sum / 0.35 * 12, full_traffic_bug=f_sum / 0.35)
    # потенциальные клиенты 05_Similarweb: I=D*E*G*H (G=не-прямые каналы, H=0,35*rel), O=I*J*K*L*12
    pk_I = [r["visits"] * r["geo"] * r["nondirect"] * COMMERCIAL_BASE * r["rel"] for r in rows]
    pk_year = sum(i * CONV_VISIT_REG[1] * CONV_REG_PRO[1] * CHECK[1] * 12 for i in pk_I)
    # платёжеспособность 07_Similarweb: D=B*C (C=geo*rel), K=D*H*J, H=0,0012
    pay_k = sum(r["rel_visits"] * conv_pay * CHECK[1] for r in rows)
    return dict(voronka=voronka, pk_I=pk_I, pk_year=pk_year, pay_month=pay_k)


def search_demand(wordstat_monthly=None):
    """shablon_excel_ocenka_rynka_cherez_poiskovyy_spros.xlsx, лист 06_Воронка.
    wordstat_monthly — список из 12 значений взвешенной частотности (04_Wordstat!H4:H15). Пока ДОСНЯТЬ."""
    ctr = (0.04, 0.06, 0.10)
    repeats = (1.0, 1.1, 1.2)
    if not wordstat_monthly:
        return None
    base = sum(wordstat_monthly) / len(wordstat_monthly)
    res = []
    for i in range(3):
        visits = base * ctr[i]
        leads = visits * CONV_VISIT_REG[i]
        sales = leads * CONV_REG_PRO[i]
        rev_m = sales * CHECK[i] * repeats[i]
        res.append(dict(visits=visits, leads=leads, sales=sales, rev_m=rev_m, rev_y=rev_m * 12))
    return res


def digital_funnel():
    """shablon_excel_metodika_ocenki_rynka_cherez_cifrovuyu_voronku.xlsx: 02 → 04 → 05 → 06."""
    ch = [dict(name=n, src=s, reach=r, ctr=c, rel=f, visits=r * c, rel_visits=r * c * f)
          for n, s, r, c, f in CHANNELS]
    rel_m = sum(c["rel_visits"] for c in ch)
    scen = []
    for i in range(3):
        v = rel_m * CHANNEL_REALISATION[i]
        leads = v * CONV_VISIT_REG[i]
        sales = leads * CONV_REG_PRO[i]
        rev_m = sales * CHECK[i]
        scen.append(dict(reach_visits=v, leads=leads, sales=sales, rev_m=rev_m, rev_y=rev_m * 12,
                         rev_y_rep=sales * CHECK[i] * FUNNEL_REPEATS[i] * 12,
                         reg_year=leads * 12, pro_year=sales * 12))
    return dict(channels=ch, visits_m=sum(c["visits"] for c in ch), rel_m=rel_m, rel_y=rel_m * 12, scen=scen)


def top_down_gdp():
    it_usd = GDP_BY_USD * IT_SHARE_GDP
    it_byn = it_usd * RATE_BYN_PER_USD
    tam = it_byn * TD_DEVTOOLS[1] * TD_MODELING[1]
    scen = []
    for i in range(3):
        sam = it_byn * TD_DEVTOOLS[i] * TD_MODELING[i] * TD_DAC[i] * TD_NOTATIONS[i]
        scen.append(dict(TAM=it_byn * TD_DEVTOOLS[i] * TD_MODELING[i], SAM=sam, SOM=sam * REACH[i]))
    return dict(it_usd=it_usd, it_byn=it_byn, TAM=tam, scen=scen)


def sverhu_snizu():
    """shablon_excel_metodika_sverhu_vniz_snizu_vverh_similarweb.xlsx: 02, 03, 04, 05, 06."""
    td = top_down_gdp()
    sverhu = [s["SAM"] for s in td["scen"]]                # 02_Сверху_вниз!B12:D12 (денежная база, B10:D10 = 1)
    pk = potential_clients()
    g11 = pk["SAM"]                                        # 03_Снизу_вверх!G11 (C = потребность × платёжеспособность)
    scen_k = (0.6, 1.0, 1.4)
    snizu = [g11 * k for k in scen_k]
    rows = sw_rows()
    H = [r["visits"] * r["geo"] * r["nondirect"] * (COMMERCIAL_BASE * r["rel"]) * r["quality"] for r in rows]
    I15 = sum(H) * 12
    b13 = 1 / COVERAGE[1]
    sw = []
    conv_k = (0.7, 1, 1.3)
    sale_k = (0.8, 1, 1.2)
    chk_k = (0.8, 1, 1.2)
    rep_k = (0.9, 1, 1.1)
    for i in range(3):
        v = I15 * scen_k[i] * b13
        leads = v * CONV_VISIT_REG[1] * conv_k[i]
        sales = leads * CONV_REG_PRO[1] * sale_k[i]
        sw.append(sales * CHECK[1] * chk_k[i] * 1.0 * rep_k[i])
    mat = [sverhu, snizu, sw]
    return dict(sverhu=sverhu, snizu=snizu, sw=sw, I15=I15, H=H, b13=b13,
                min=[min(m[i] for m in mat) for i in range(3)],
                avg=[sum(m[i] for m in mat) / 3 for i in range(3)],
                max=[max(m[i] for m in mat) for i in range(3)])


def scenario_sw():
    """shablon_excel_scenarnaya_ocenka_rynka_s_similarweb.xlsx: 02 → 03 → 04; 05 не применим."""
    rows = sw_rows()
    I = [r["visits"] * r["geo"] * (r["rel"] * COMMERCIAL_BASE) * r["quality"] for r in rows]
    I21 = sum(I)
    unc = tuple(1 / c for c in COVERAGE)                  # охват неучтённых конкурентов = 1 / покрытие
    reliab = (0.75, 0.9, 1.0)
    adj = (0.8, 1.0, 1.15)
    chk_mult = (0.6, 1.0, 1.2)
    out = []
    for i in range(3):
        f = I21 * unc[i] * reliab[i] * adj[i]
        g = f * 12
        sales = g * CONV_VISIT_REG[i] * CONV_REG_PRO[i]
        chk = CHECK[1] * chk_mult[i]
        sam = sales * chk * 1
        out.append(dict(traffic_m=f, traffic_y=g, sales=sales, check=chk, SAM=sam, SOM=sam * REACH[i]))
    return dict(I=I, I21=I21, unc=unc, scen=out)


def sw_volume():
    """metodika_ocenki_obema_rynka_po_similarweb.xlsx (с исправлением B9/B10 → G4:G23/H4:H23)."""
    rows = sw_rows()
    G = [r["rel_visits"] for r in rows]
    B9 = sum(G)
    B9_bug = sum(G[1:])        # как считает оригинальный шаблон (строка 4 теряется)
    shares = [g / B9 for g in G]
    out = []
    for i in range(3):
        c = B9 / COVERAGE[i]
        orders = c * CONV_VISIT_REG[i] * CONV_REG_PRO[i]
        j = orders * CHECK[i] * 12 * 1
        out.append(dict(market_traffic=c, orders_m=orders, rev_y=j, SOM=j * REACH[i]))
    return dict(B9=B9, B9_bug=B9_bug, B12=B9 / COVERAGE[1], shares=shares, scen=out)


def lr3_similarweb():
    """Методики ЛР3 (оба docx): V_sample, V_geo, V_total, CR3/CR5, HHI, TAM/SAM/SOM цифрового внимания."""
    rows = sw_rows()
    v_sample = sum(r["visits"] for r in rows)
    v_geo = sum(r["geo_visits"] for r in rows)
    v_rel = sum(r["rel_visits"] for r in rows)

    def conc(vals):
        tot = sum(vals)
        sh = sorted((v / tot * 100 for v in vals), reverse=True)
        return dict(shares=sh, CR3=sum(sh[:3]), CR5=sum(sh[:5]), HHI=sum(s * s for s in sh))

    conc_rel = conc([r["rel_visits"] for r in rows])
    conc_raw = conc([r["visits"] for r in rows])
    scen = []
    for i in range(3):
        v_total = v_rel / COVERAGE[i]
        comm = v_total * COMMERCIAL[i]
        leads = comm * CONV_VISIT_REG[i]
        sales = leads * CONV_REG_PRO[i]
        rev_m = sales * CHECK[i]
        scen.append(dict(V_total=v_total, commercial=comm, leads=leads, sales=sales, rev_m=rev_m,
                         rev_y=rev_m * 12, SOM=rev_m * 12 * REACH[i]))
    # Итоговая модель (разд. 11 docx «по данным Similarweb»)
    sam_channel = sum(r["geo_visits"] * r["nondirect"] for r in rows)
    som_traffic = sam_channel * REACH[1]
    som_revenue = som_traffic * CONV_VISIT_REG[1] * CONV_REG_PRO[1] * CHECK[1] * 12
    return dict(V_sample=v_sample, V_geo=v_geo, V_rel=v_rel, conc_rel=conc_rel, conc_raw=conc_raw, scen=scen,
                TAM_attention=v_sample, SAM_geo=v_geo, SAM_channel=sam_channel,
                SOM_traffic=som_traffic, SOM_revenue=som_revenue)


def hhi_level(h):
    return "низкая" if h < 1500 else ("умеренная" if h <= 2500 else "высокая")


def payment():
    """shablon_excel_metodika_proverki_rynka_cherez_platezhesposobnost.xlsx: 02, 03, 04, 05, 06."""
    pk = potential_clients()
    seg = [dict(name=f'{r["code"]} {r["name"]}', B=r["F"], C=r["G"], D=r["H"], E=r["I"], F=1, G=r["K"],
                H=r["F"] * r["G"] * r["H"] * r["I"]) for r in pk["rows"]]
    for s in seg:
        s["J"] = s["H"] * s["F"] * s["G"]
    J16 = sum(s["J"] for s in seg)
    econ = []
    for (name, budget, hours, rate, save) in PAY:
        B = hours * rate                 # стоимость времени на диаграммы, BYN/мес
        E = B * 1.0 * save               # денежная оценка экономии
        F = PRO_BYN_MONTH
        econ.append(dict(name=name, B=B, E=E, F=F, payback=F / E, ROI=((E - F) * 12) / (F * 12), K=E / F,
                         budget=budget, index=budget / F))
    idx_avg = sum(e["index"] for e in econ) / len(econ)
    roi_avg = sum(e["ROI"] for e in econ) / len(econ)
    scen = []
    for i, m in enumerate(PAY_MULT):
        k = m[0] * m[1] * m[2] * m[3]
        scen.append(dict(market=J16 * k, reach=J16 * k * REACH[i]))
    corridor = dict(
        month=dict(min=PRO_BYN_MONTH, med=(ERASER["starter_year"] + ERASER["starter_month"]) / 2 * RATE_BYN_PER_USD,
                   max=ERASER["business_month"] * RATE_BYN_PER_USD, model=PRO_BYN_MONTH),
        year=dict(min=PRO_USD_YEAR / 12 * RATE_BYN_PER_USD, med=ERASER["starter_year"] * RATE_BYN_PER_USD,
                  max=ERASER["business_year"] * RATE_BYN_PER_USD, model=PRO_USD_YEAR / 12 * RATE_BYN_PER_USD))
    for c in corridor.values():
        c["margin"] = (c["model"] - PRO_COST_BYN_MONTH) / c["model"]
        c["access"] = c["model"] / c["med"]
    return dict(seg=seg, J16=J16, econ=econ, idx_avg=idx_avg, roi_avg=roi_avg, scen=scen, corridor=corridor)


def kpi_check():
    reg = 1500
    pro = reg * CONV_REG_PRO[1]
    return dict(reg=reg, pro=pro, rev_usd=pro * PRO_USD_MONTH * 12, rev_byn=pro * PRO_BYN_MONTH * 12)


def compute_all():
    pk = potential_clients()
    fun = digital_funnel()
    td = top_down_gdp()
    ss = sverhu_snizu()
    sc = scenario_sw()
    vol = sw_volume()
    lr3 = lr3_similarweb()
    pay = payment()
    chk = sw_competitor_checks()
    # Сверка на уровне SOM (достижимая выручка 1-го года)
    som = {
        "Цифровая воронка (каналы NotaCode)": [s["rev_y_rep"] for s in fun["scen"]],
        "Потенциальные клиенты (SOM)": [s["SOM"] for s in pk["scen"]],
        "Платёжеспособность (достижимая)": [s["reach"] for s in pay["scen"]],
        "Сценарная SW → воронка (SOM)": [s["SOM"] for s in sc["scen"]],
        "Объём по SW (SOM)": [s["SOM"] for s in vol["scen"]],
        "ЛР3: SW-методика (SOM)": [s["SOM"] for s in lr3["scen"]],
        "Сверху вниз от ВВП (SOM)": [s["SOM"] for s in td["scen"]],
    }
    sam = {
        "Потенциальные клиенты (SAM)": [s["SAM"] for s in pk["scen"]],
        "Платёжеспособность (рынок)": [s["market"] for s in pay["scen"]],
        "Снизу вверх (шаблон 03)": ss["snizu"],
        "Сценарная SW → воронка (SAM)": [s["SAM"] for s in sc["scen"]],
        "SW + воронка (шаблон 05)": ss["sw"],
        "Объём по SW (рынок/год)": [s["rev_y"] for s in vol["scen"]],
        "ЛР3: SW-методика (рынок/год)": [s["rev_y"] for s in lr3["scen"]],
        "Сверху вниз от ВВП (SAM)": ss["sverhu"],
    }
    # Сверка для листа 06_Сверка сценарного шаблона: строки 5, 7, 8, 9
    sverka_rows = {
        "Similarweb → цифровая воронка": [s["SOM"] for s in sc["scen"]],
        "Цифровая воронка (каналы NotaCode)": [s["rev_y_rep"] for s in fun["scen"]],
        "Количество потенциальных клиентов": [s["SOM"] for s in pk["scen"]],
        "Сверху вниз (ВВП) / снизу вверх": [s["SOM"] for s in td["scen"]],
    }
    sverka = dict(rows=sverka_rows,
                  min=[min(v[i] for v in sverka_rows.values()) for i in range(3)],
                  avg=[sum(v[i] for v in sverka_rows.values()) / len(sverka_rows) for i in range(3)],
                  max=[max(v[i] for v in sverka_rows.values()) for i in range(3)])
    return dict(params=dict(rate=RATE_BYN_PER_USD, rate_date=RATE_DATE, check=CHECK, pro_byn_month=PRO_BYN_MONTH),
                potential=pk, funnel=fun, topdown=td, sverhu_snizu=ss, scenario_sw=sc, sw_volume=vol,
                lr3=lr3, payment=pay, sw_checks=chk, som=som, sam=sam, sverka=sverka, kpi=kpi_check(),
                search=search_demand(None))


def fmt(x, d=0):
    if x is None:
        return "—"
    s = f"{x:,.{d}f}".replace(",", " ")
    return s.replace(".", ",")


def main():
    r = compute_all()
    p = r["potential"]
    print(f"Курс (допущение): 1 USD = {RATE_BYN_PER_USD} BYN на {RATE_DATE}; Pro = {PRO_BYN_MONTH:.0f} BYN/мес; "
          f"годовой чек = {CHECK} BYN")
    print("\n== Потенциальные клиенты ==")
    for row in p["rows"]:
        print(f'{row["code"]:5} D={fmt(row["D"])} F={fmt(row["F"])} J={fmt(row["J"],1)} SAM={fmt(row["M"],1)} '
              f'SOM={fmt(row["O"],1)}')
    print(f'TAM={fmt(p["TAM"])} SAM={fmt(p["SAM"])} SOM={fmt(p["SOM"])} адресуемых={fmt(p["addressable"],1)}')
    for n, s in zip(SCEN, p["scen"]):
        print(f'  {n:14} TAM={fmt(s["TAM"])} SAM={fmt(s["SAM"])} SOM={fmt(s["SOM"])}')
    f = r["funnel"]
    print(f'\n== Цифровая воронка == релевантных визитов/мес={fmt(f["rel_m"],1)} /год={fmt(f["rel_y"])}')
    for n, s in zip(SCEN, f["scen"]):
        print(f'  {n:14} визиты={fmt(s["reach_visits"],1)} рег/мес={fmt(s["leads"],2)} Pro/мес={fmt(s["sales"],3)} '
              f'выручка/год={fmt(s["rev_y"])} с повторами={fmt(s["rev_y_rep"])}')
    td = r["topdown"]
    print(f'\n== Сверху вниз (ВВП) == ИТ-сектор={fmt(td["it_usd"])} USD = {fmt(td["it_byn"])} BYN; TAM={fmt(td["TAM"])}')
    for n, s in zip(SCEN, td["scen"]):
        print(f'  {n:14} TAM={fmt(s["TAM"])} SAM={fmt(s["SAM"])} SOM={fmt(s["SOM"])}')
    ss = r["sverhu_snizu"]
    print(f'\n== Шаблон сверху/снизу == I15={fmt(ss["I15"])} B13={ss["b13"]:.4f}')
    for k in ("sverhu", "snizu", "sw", "min", "avg", "max"):
        print(f'  {k:6} ' + " | ".join(fmt(v) for v in ss[k]))
    sc = r["scenario_sw"]
    print(f'\n== Сценарная SW == I21={fmt(sc["I21"],1)}')
    for n, s in zip(SCEN, sc["scen"]):
        print(f'  {n:14} трафик/год={fmt(s["traffic_y"])} продаж={fmt(s["sales"],1)} SAM={fmt(s["SAM"])} '
              f'SOM={fmt(s["SOM"])}')
    v = r["sw_volume"]
    print(f'\n== Объём по SW == B9={fmt(v["B9"],1)} (ошибка шаблона дала бы {fmt(v["B9_bug"],1)}) B12={fmt(v["B12"])}')
    for n, s in zip(SCEN, v["scen"]):
        print(f'  {n:14} трафик={fmt(s["market_traffic"])} рынок/год={fmt(s["rev_y"])} SOM={fmt(s["SOM"])}')
    l3 = r["lr3"]
    print(f'\n== ЛР3 == V_sample={fmt(l3["V_sample"])} V_geo={fmt(l3["V_geo"])} V_rel={fmt(l3["V_rel"],1)}')
    print(f'  rel: CR3={l3["conc_rel"]["CR3"]:.2f}% CR5={l3["conc_rel"]["CR5"]:.2f}% HHI={l3["conc_rel"]["HHI"]:.0f} '
          f'({hhi_level(l3["conc_rel"]["HHI"])}); raw HHI={l3["conc_raw"]["HHI"]:.0f}')
    for n, s in zip(SCEN, l3["scen"]):
        print(f'  {n:14} V_total={fmt(s["V_total"])} комм={fmt(s["commercial"])} рынок/год={fmt(s["rev_y"])} '
              f'SOM={fmt(s["SOM"])}')
    print(f'  SAM_channel={fmt(l3["SAM_channel"])} SOM_traffic={fmt(l3["SOM_traffic"])} SOM_rev={fmt(l3["SOM_revenue"])}')
    pay = r["payment"]
    print(f'\n== Платёжеспособность == J16={fmt(pay["J16"])} индекс ср.={pay["idx_avg"]:.2f} ROI ср.={pay["roi_avg"]:.2f}')
    for e in pay["econ"]:
        print(f'  {e["name"]:20} эффект={fmt(e["E"])} K={e["K"]:.2f} ROI={e["ROI"]:.2f} индекс={e["index"]:.2f}')
    for n, s in zip(SCEN, pay["scen"]):
        print(f'  {n:14} рынок={fmt(s["market"])} достижимая={fmt(s["reach"])}')
    ch = r["sw_checks"]
    print(f'\n== Проверки SW == воронка C12={fmt(ch["voronka"]["full_traffic"])} C14={fmt(ch["voronka"]["revenue_year"])}'
          f' (с ошибкой B12: {fmt(ch["voronka"]["revenue_year_bug"])}); PK N12={fmt(ch["pk_year"])}; '
          f'PAY K14={fmt(ch["pay_month"])}')
    print("\n== Сверка (06_Сверка сценарного шаблона) ==")
    for k, vals in r["sverka"]["rows"].items():
        print(f'  {k:40} ' + " | ".join(fmt(x) for x in vals))
    for k in ("min", "avg", "max"):
        print(f'  {k:40} ' + " | ".join(fmt(x) for x in r["sverka"][k]))
    print(f'\nKPI-проверка: {r["kpi"]}')
    out = HERE / "results_notacode.json"
    out.write_text(json.dumps(r, ensure_ascii=False, indent=1, default=float), encoding="utf-8")
    print(f"\nСохранено: {out}")


if __name__ == "__main__":
    main()

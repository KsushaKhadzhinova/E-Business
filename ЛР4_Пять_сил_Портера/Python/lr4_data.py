# -*- coding: utf-8 -*-
"""
ЛР4. Пять сил Портера и барьеры входа — NotaCode.
Общий модуль данных и расчётов (импортируется скриптами рисунков).
Запуск отдельно:  python lr4_data.py  — печатает все расчёты для отчёта.

Источники трафика: ТОЛЬКО данные SimilarWeb из старой версии работы
(_архив_DiagramCode/labs/lab2-3-market-size, снято 17.09.2026 и 23.09.2026,
бесплатный веб-доступ, период «последние 3 месяца», география не фильтровалась).
Для остальных доменов визиты не сняты -> None (🔲 ДОСНЯТЬ).
"""
import os
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = os.path.dirname(os.path.abspath(__file__))
FIG_DIR = os.path.normpath(os.path.join(HERE, "..", "Рисунки"))
os.makedirs(FIG_DIR, exist_ok=True)

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["axes.spines.top"] = False
plt.rcParams["axes.spines.right"] = False

# Палитра (категориальная, фиксированный порядок) и чернила
C = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#4a3aa7", "#e34948"]
INK, INK2, MUTED, SURF = "#0b0b0b", "#52514e", "#898781", "#fcfcfb"

# ---------------------------------------------------------------------------
# 1. Конкуренты (лист 02_Конкуренты_SW). visits=None -> не снято (ДОСНЯТЬ)
#    geo_ub — верхняя граница доли RU(+BY), выведенная из топ-стран SimilarWeb:
#    plantuml.com — Россия 9,07 % (факт, топ-1 страна);
#    app.diagrams.net — RU нет в топ-5, 5-я страна Вьетнам 4,35 % -> RU < 4,35 %;
#    eraser.io — RU нет в топ-5, 3-я страна ЮАР 2,88 % -> RU < 2,88 %;
#    mermaidchart.com — топ-3 = 86,35 % -> RU <= 13,65 % (очень грубая граница).
# ---------------------------------------------------------------------------
COMPETITORS = [
    # домен, тип, визиты/период, стр./визит, отказы, длит. сек, direct, geo_ub
    ("app.diagrams.net (draw.io)", "GUI-редактор (заменитель)", 7_800_000, 2.98, 0.555, None, 0.6865, 0.0435),
    ("eraser.io", "прямой (DSL + AI)", 678_500, 3.57, 0.4373, 99, 0.4905, 0.0288),
    ("plantuml.com", "прямой (DSL, open-source)", 507_000, None, None, None, None, 0.0907),
    ("mermaidchart.com", "прямой (DSL, коммерч.)", 20_000, None, None, None, None, 0.1365),
    ("mermaid.live", "прямой (DSL, open-source)", None, None, None, None, None, None),
    ("d2lang.com", "прямой (DSL)", None, None, None, None, None, None),
    ("graphviz.org", "прямой (DSL-движок)", None, None, None, None, None, None),
    ("structurizr.com", "нишевой (C4 DSL)", None, None, None, None, None, None),
    ("gleek.io", "прямой (DSL + AI)", None, None, None, None, None, None),
    ("lucidchart.com", "GUI SaaS (частичный)", None, None, None, None, None, None),
    ("miro.com", "платформа-доска (частичный)", None, None, None, None, None, None),
    ("camunda.com (Modeler)", "нишевой (BPMN)", None, None, None, None, None, None),
    ("bpmn.io", "нишевой (BPMN, open-source)", None, None, None, None, None, None),
    ("ramussoftware.com (Ramus)", "нишевой (IDEF0/DFD)", None, None, None, None, None, None),
    ("sparxsystems.com (EA)", "CASE enterprise", None, None, None, None, None, None),
    ("visual-paradigm.com", "CASE enterprise", None, None, None, None, None, None),
    ("dbdiagram.io", "нишевой (ERD DSL)", None, None, None, None, None, None),
    ("kroki.io", "агрегатор-рендер", None, None, None, None, None, None),
]


def concentration(shares):
    s = sorted(shares, reverse=True)
    cr3 = sum(s[:3])
    cr5 = sum(s[:5])
    hhi = sum(x * x for x in s) * 10000
    return cr3, cr5, hhi


def scenario(name, rows):
    """rows: list of (domain, relevant_visits)."""
    total = sum(v for _, v in rows)
    shares = [(d, v / total) for d, v in rows]
    cr3, cr5, hhi = concentration([x for _, x in shares])
    return {"name": name, "total": total, "shares": shares, "cr3": cr3, "cr5": cr5, "hhi": hhi}


measured = [(c[0], c[2]) for c in COMPETITORS if c[2]]
SCEN_A = scenario("A. Глобальный трафик, 4 домена\n(доля геогр. = 1,0 — допущение)", measured)
SCEN_B = scenario("B. Верхняя граница трафика RU\n(доли стран из топа SimilarWeb)",
                  [(c[0], c[2] * c[7]) for c in COMPETITORS if c[2]])
SCEN_C = scenario("C. Только DSL-ниша\n(без app.diagrams.net)",
                  [(d, v) for d, v in measured if not d.startswith("app.diagrams")])
SCENARIOS = [SCEN_A, SCEN_B, SCEN_C]

# Взвешенная доля прямого трафика (лист 03, строка «Прямой трафик») — как в шаблоне:
# SUMPRODUCT(E, J)/SUM(E); незаполненные J считаются 0 => это НИЖНЯЯ граница.
DIRECT_W = sum(c[2] * (c[6] or 0) for c in COMPETITORS if c[2]) / SCEN_A["total"]

# ---------------------------------------------------------------------------
# 2. Пять сил Портера + цифровые усилители (шкала 1/3/5, методика 5 сил, разд. 12)
# ---------------------------------------------------------------------------
FORCES = [
    ("Конкуренция\nигроков", 5),
    ("Угроза новых\nучастников", 3),
    ("Сила\nпокупателей", 5),
    ("Сила\nпоставщиков", 3),
    ("Угроза\nзаменителей", 5),
    ("Цифровые\nусилители", 3),
]
DIGITAL = [
    ("Сетевые эффекты", 3),
    ("Данные", 1),
    ("Алгоритмическая видимость", 3),
    ("Платформенная зависимость", 3),
    ("Издержки переключения", 1),
    ("Экосистемная связанность", 3),
]


def interp_pressure(x):
    if x <= 2.0:
        return "низкое давление"
    if x <= 3.5:
        return "умеренное давление"
    return "высокое давление"


# ---------------------------------------------------------------------------
# 3. Индекс барьеров входа (лист 07_Барьеры_входа)
#    source: 'расчёт' — балл получен формулой шаблона из реальных данных;
#            'эксперт' — данных нет, балл из листа «Ввод_NotaCode» (допущение).
# ---------------------------------------------------------------------------
BARRIERS = [
    # фактор, вес шаблона, балл (база), источник
    ("Концентрация трафика", 0.16, 5, "расчёт"),
    ("Органическое SEO-давление", 0.12, 3, "эксперт"),
    ("Платное рекламное давление", 0.12, 3, "эксперт"),
    ("Брендовая сила", 0.12, 5, "расчёт"),
    ("Поисково-рекламная конкуренция", 0.12, 3, "эксперт"),
    ("Репутационный барьер", 0.12, 5, "эксперт"),
    ("Технологический и ресурсный барьер", 0.12, 3, "эксперт"),
    ("Платформенная зависимость (0,5)", 0.08, 3, "эксперт"),
    ("Угроза заменителей (0,8)", 0.08, 5, "эксперт"),
]
W_SUM = sum(b[1] for b in BARRIERS)  # 1.04 в шаблоне


def barrier_index(scores, normalize=True):
    ws = [b[1] / W_SUM if normalize else b[1] for b in BARRIERS]
    return sum(s * w for s, w in zip(scores, ws))


def barrier_level(x):
    if x >= 4:
        return "высокие барьеры"
    if x >= 2.5:
        return "средние барьеры"
    return "низкие барьеры"


BASE = [b[2] for b in BARRIERS]
PESS = [b[2] if b[3] == "расчёт" else min(5, b[2] + 2) for b in BARRIERS]
OPT = [b[2] if b[3] == "расчёт" else max(1, b[2] - 2) for b in BARRIERS]

# 8 критериев методики Similarweb (шкала 1–5, допускаются промежуточные баллы)
SW8 = [
    ("Количество релевантных конкурентов", 5, "18 игроков, в т.ч. сильные бренды (draw.io, Lucidchart, Miro)"),
    ("Концентрация трафика", 5, "CR3 = 99,8 % по 4 измеренным доменам"),
    ("Стоимость входа в каналы", 3, "смешанная модель: органика/GitHub/вузы + реклама (экспертно)"),
    ("Поисковая конкуренция", 4, "выдача по англоязычным запросам занята сильными доменами; по рус. IDEF/DFD — 🔲 проверить"),
    ("Рекламная конкуренция", 3, "данных рекламных библиотек нет (🔲 ДОСНЯТЬ)"),
    ("Репутационный барьер", 5, "PlantUML с 2009 г., draw.io, Lucidchart — массовые отзывы"),
    ("Технологический барьер", 4, "нужны парсер DSL, валидация нотаций, раскладка, версии — не «простой сайт»"),
    ("Издержки переключения клиента", 2, "текстовые форматы и экспорт — переключение лёгкое"),
]


def fmt(x, n=2):
    return f"{x:.{n}f}".replace(".", ",")


if __name__ == "__main__":
    for s in SCENARIOS:
        print("==", s["name"].replace("\n", " "))
        print("  суммарный трафик:", round(s["total"]))
        for d, x in s["shares"]:
            print(f"   {d:30s} {x*100:6.2f} %")
        print(f"  CR3={s['cr3']*100:.2f} %  CR5={s['cr5']*100:.2f} %  HHI={s['hhi']:.0f}")
    print("Взвеш. доля direct (нижняя граница):", round(DIRECT_W, 4))
    f5 = [v for _, v in FORCES[:5]]
    f6 = [v for _, v in FORCES]
    print("Среднее 5 сил:", sum(f5) / 5, " среднее 6 строк:", round(sum(f6) / 6, 3), interp_pressure(sum(f6) / 6))
    print("Среднее цифровых факторов:", round(sum(v for _, v in DIGITAL) / 6, 3))
    print("Сумма весов шаблона:", W_SUM)
    for nm, sc in [("база", BASE), ("пессим.", PESS), ("оптим.", OPT)]:
        a = barrier_index(sc, True)
        b = barrier_index(sc, False)
        print(f"Индекс барьеров {nm}: норм.={a:.3f} ({barrier_level(a)}), ненорм.(1,04)={b:.3f} ({barrier_level(b)})")
    print("Нормированные веса:", [round(b[1] / W_SUM, 4) for b in BARRIERS])
    print("SW 8 критериев:", sum(x[1] for x in SW8) / 8)

# -*- coding: utf-8 -*-
"""
ЛР5 «Анализ конкурентов» для NotaCode — ЕДИНЫЙ ИСТОЧНИК ДАННЫХ.

Из этого модуля берут данные:
  * lr5_calc.py            — расчёты по формулам трёх Excel-шаблонов;
  * lr5_charts.py          — рисунки в ../Рисунки/;
  * export_vba_data.py     — TSV-файлы для VBA-модулей (../VBA/data/);
  * fill_templates_excel.py — заполнение копий шаблонов через Excel (проверка).

Происхождение фактов:
  * D17 = 2026-09-17, D23 = 2026-09-23 — даты снятия в старой версии работы
    (_архив_DiagramCode/labs/lab5-competitors, lab2-3-market-size) и в
    E:/VisualDSL-Platform/PROJECT_DOCUMENTATION/06_market_and_extras/00-competitive-analysis.md;
  * «kursach» — E:/NotaCode/docs/source-materials/kursach.md (таблица инструментов);
  * «концепция» — E:/NotaCode/docs/business/01-project-concept.md.
Всё, чего нет в этих источниках, помечено «🔲 проверить на сайте» или «(допущение)».
Баллы — экспертная оценка автора работы (Хаджинова К.) по этим фактам.
"""

D17 = "2026-09-17"
D23 = "2026-09-23"
TODO = "🔲 проверить на сайте"
TODO_DATE = "🔲 дата просмотра"
RESP = "Хаджинова К."

SRC_OLD = "архив ЛР5 (DiagramCode), снято 2026-09-17"
SRC_OLD23 = "архив ЛР5 (DiagramCode), снято 2026-09-23"
SRC_SW = "SimilarWeb, архив ЛР2-3"
SRC_CA = "00-competitive-analysis.md (VisualDSL), сверено 2026-09-17"
SRC_KURS = "kursach.md (NotaCode, таблица инструментов)"
SRC_CONC = "01-project-concept.md (NotaCode)"

# =====================================================================
# ЧАСТЬ A. Классификация по уровням конкуренции (шаблон 10 критериев)
# =====================================================================
CRIT = [  # код, критерий, вес (лист 03_Критерии, без изменений)
    ("C1", "Потребительская задача", 20),
    ("C2", "Целевая аудитория", 15),
    ("C3", "Функциональная заменяемость", 18),
    ("C4", "Сценарий использования", 10),
    ("C5", "Ценовой уровень и форма оплаты", 8),
    ("C6", "Каналы продаж и привлечения", 8),
    ("C7", "География, язык, регулирование", 7),
    ("C8", "Модель монетизации", 7),
    ("C9", "Зрелость и доверие", 4),
    ("C10", "Вероятность переключения", 3),
]
WEIGHTS_A = [c[2] for c in CRIT]

LEVELS_A = [
    "Прямой конкурент",
    "Близкий альтернативный конкурент",
    "Косвенный конкурент / заменитель",
    "Смежная или потенциальная конкуренция",
    "Не конкурент / объект наблюдения",
    "Недостаточно данных",
]

OBJECT_A = [  # п. 10.1 методички
    ("Название объекта анализа", "NotaCode — web-IDE (PWA) «diagram as code» для формальных нотаций: UML (14 типов), BPMN 2.0, ERD, IDEF0, IDEF1X, IDEF3, DFD, сети Петри"),
    ("Основная потребительская задача", "Быстро построить диаграмму строгой нотации по текстовому описанию и получить её без нарушений правил нотации (проверка кода правила и строки)"),
    ("Целевой сегмент", "S-01 студенты технических специальностей (первичный, Free); S-02 бизнес/системные аналитики и S-03 архитекторы ПО (Pro); S-04 преподаватели"),
    ("Сценарий использования", "Лабораторные, курсовые и дипломные работы; проектная документация; модели процессов и данных рядом с требованиями в Git"),
    ("География и язык", "Беларусь (+РФ/СНГ), русскоязычный интерфейс и документация"),
    ("Ценовой диапазон", "Freemium: Free — 0 (5 проектов, 20 версий, экспорт SVG/PNG); Pro — 4 USD/мес или 40 USD/год (допущение, 01-project-concept.md)"),
    ("Ключевой результат", "Корректная диаграмма за минуты: DSL → валидация → раскладка elkjs → SVG, версии/diff/откат, экспорт в 11 форматов"),
]

SEARCH_DIRECTIONS = [
    ("По товарной категории", "«diagram as code», «text to diagram», «UML online», «редактор диаграмм онлайн»", "PlantUML, Mermaid, D2, Gleek, Eraser, Structurizr"),
    ("По задаче клиента", "«построить IDEF0 онлайн», «DFD диаграмма онлайн», «ER диаграмма онлайн», «BPMN редактор»", "Ramus, dbdiagram.io, Camunda Modeler, Visual Paradigm Online, draw.io"),
    ("По результату", "«диаграмма для курсовой», «UML для лабораторной»", "draw.io, PlantUML, Lucidchart"),
    ("Каталоги и рейтинги", "подборки «PlantUML alternatives», «Mermaid vs D2 vs Graphviz» (diagrams.so, architecturediagram.ai)", "D2, Graphviz, Kroki, Mermaid"),
    ("Реклама и выдача", "🔲 ДОСНЯТЬ: рекламные объявления по запросам «uml онлайн», «diagram as code» (Яндекс/Google)", "Lucidchart, Miro, Visual Paradigm Online (гипотеза)"),
    ("Отзывы и обсуждения", "🔲 ДОСНЯТЬ: Habr, Reddit r/softwarearchitecture, чаты студентов", "ChatGPT и др. AI-ассистенты (генерация кода PlantUML/Mermaid)"),
    ("Смежные рынки", "доски и дизайн-инструменты, CASE-средства, BPM-платформы", "Miro, Figma/FigJam, Microsoft Visio, Sparx EA"),
]

# scores: C1..C10 (0–5). corr — экспертная корректировка (пусто, если нет).
COMPANIES = [
    dict(id="K001", name="PlantUML", site="https://plantuml.com", market="Международный, EN; РФ — 1-я страна по трафику (9,07 %)",
         type="Open-source инструмент diagram as code", product="Текстовый DSL → UML-диаграммы (PNG/SVG/ASCII); онлайн-сервер, плагины IDE, CLI",
         segment="Разработчики, студенты, техписатели", hyp="Прямой конкурент", prio="Высокий", status="Проверено", date=D17,
         found="Категория «diagram as code»; ЛР1 (Беларусь в топ-5 по интересу к PlantUML); SimilarWeb (ЛР2-3)",
         scores=[4, 4, 4, 4, 3, 4, 4, 1, 5, 4], corr="Прямой конкурент", conf="высокая",
         just="Уверенность: высокая. Та же задача «текст → UML-диаграмма», та же аудитория (РФ — 1-я страна трафика), бесплатен — прямая альтернатива Free-тарифу; нет IDEF/DFD и проверки правил. Корректировка: балл 3,75 (=75 из 100 по шкале методички, диапазон 65–79) повышен до прямого по правилу «постоянно в тех же запросах» — 4 подтверждения (см. 02_Доказательства).",
         links="https://plantuml.com; SimilarWeb plantuml.com (2026-09-17)"),
    dict(id="K002", name="Mermaid (open source)", site="https://mermaid.js.org", market="Международный, EN",
         type="Open-source JS-библиотека diagram as code", product="Markdown-подобный DSL → SVG (flowchart, sequence, class, ER, state, Gantt)",
         segment="Разработчики, техписатели", hyp="Близкий альтернативный конкурент", prio="Высокий", status="Проверено", date=D17,
         found="Категория «diagram as code»; нативный рендер в GitHub/GitLab/Notion",
         scores=[4, 3, 3, 3, 3, 4, 3, 1, 5, 4], corr="", conf="высокая",
         just="Уверенность: высокая. Та же задача для UML-части, но сценарий — диаграммы внутри Markdown-документации; нет BPMN/IDEF/DFD и валидации нотаций.",
         links="https://mermaid.js.org"),
    dict(id="K003", name="Mermaid Chart", site="https://www.mermaidchart.com", market="Международный, EN; трафик: Вьетнам 44,98 %, Индонезия 26,27 %, США 15,1 %",
         type="Коммерческий SaaS над Mermaid", product="Web-редактор Mermaid-диаграмм для команд",
         segment="Команды разработки", hyp="Косвенный конкурент / заменитель", prio="Средний", status="Требует уточнения", date=D17,
         found="Коммерческий продукт команды Mermaid; SimilarWeb (ЛР2-3)",
         scores=[4, 3, 3, 3, 1, 3, 1, 3, 3, 4], corr="", conf="средняя",
         just="Уверенность: средняя. Товар близок к Mermaid, но тарифы без регистрации не раскрыты (C5 = 1, проверить), русскоязычный сегмент почти не представлен (C7 = 1).",
         links="https://www.mermaidchart.com; SimilarWeb mermaidchart.com (2026-09-17)"),
    dict(id="K004", name="Eraser (DiagramGPT)", site="https://www.eraser.io", market="Международный, EN; трафик: Индия 40,81 %, США 8,61 %",
         type="SaaS diagram as code + AI", product="Diagram as code + AI-генерация + канвас + markdown-заметки",
         segment="Команды разработки (от Free до Enterprise)", hyp="Прямой конкурент", prio="Высокий", status="Проверено", date=D23,
         found="Категория «AI diagram», «diagram as code»; SimilarWeb (ЛР2-3)",
         scores=[4, 3, 3, 4, 3, 4, 2, 5, 4, 3], corr="", conf="высокая",
         just="Уверенность: высокая. Ближайший аналог по модели (freemium + подписка, AI, DSL), но цена 15–20 USD/участник/мес против 4 USD, нет IDEF/DFD/сетей Петри, русскоязычный сегмент не в топ-5 стран.",
         links="https://www.eraser.io; https://www.eraser.io/diagramgpt; SimilarWeb eraser.io (2026-09-23)"),
    dict(id="K005", name="D2 (Terrastruct)", site="https://d2lang.com", market="Международный, EN",
         type="Open-source язык диаграмм", product="Декларативный DSL + CLI, 3 движка раскладки, экспорт SVG/PNG/PDF/PPTX/GIF/ASCII",
         segment="Разработчики, техписатели", hyp="Близкий альтернативный конкурент", prio="Средний", status="Проверено", date=D23,
         found="Подборки «Mermaid vs D2 vs Graphviz»",
         scores=[4, 3, 3, 3, 3, 3, 3, 1, 3, 4], corr="", conf="высокая",
         just="Уверенность: высокая. Та же идея «текст → диаграмма», но архитектурные/сетевые схемы; UML только class/sequence, нет BPMN/IDEF/DFD; без монетизации (fiscal sponsor Hack Club).",
         links="https://d2lang.com; https://d2lang.com/tour/faq/"),
    dict(id="K006", name="Gleek", site="https://www.gleek.io", market="Международный, EN (🔲 проверить наличие RU)",
         type="SaaS diagram as code", product="Собственный текстовый DSL → UML-подобные, ER, flowchart в веб-редакторе",
         segment="Разработчики, аналитики (🔲 проверить)", hyp="Близкий альтернативный конкурент", prio="Средний", status="Требует уточнения", date=TODO_DATE,
         found="Таблица инструментов kursach.md; «uml online»",
         scores=[4, 3, 3, 4, 3, 3, 2, 4, 2, 2], corr="", conf="низкая",
         just="Уверенность: низкая. Freemium с Free-тарифом и собственным DSL (kursach.md) — близко к модели NotaCode; цены Pro, нотации и аудитория 🔲 проверить на сайте.",
         links="https://www.gleek.io"),
    dict(id="K007", name="Structurizr", site="https://structurizr.com", market="Международный, EN",
         type="Инструмент «архитектура как код» (C4)", product="DSL для модели C4 → несколько видов диаграмм из одной модели",
         segment="Архитекторы ПО (S-03)", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Проверено", date=D17,
         found="Категория «architecture as code»; архив ЛР5 (эталон практик)",
         scores=[3, 3, 2, 3, 2, 2, 2, 2, 3, 1], corr="", conf="средняя",
         just="Уверенность: средняя. Только модель C4 → заменяет малую часть задачи (C3 = 2, ограничитель). Используется как эталон практики «одна модель — несколько представлений». Тарифы 🔲 проверить.",
         links="https://structurizr.com"),
    dict(id="K008", name="dbdiagram.io", site="https://dbdiagram.io", market="Международный, EN (🔲 проверить)",
         type="SaaS «database as code»", product="DSL DBML → ER-диаграмма, экспорт SQL DDL (🔲 проверить)",
         segment="Разработчики БД, студенты курсов БД (🔲 проверить)", hyp="Косвенный конкурент / заменитель", prio="Средний", status="Требует уточнения", date=TODO_DATE,
         found="Запрос «ER диаграмма онлайн по коду» (🔲 зафиксировать выдачу)",
         scores=[3, 3, 2, 3, 3, 3, 2, 4, 3, 3], corr="", conf="низкая",
         just="Уверенность: низкая (фактов в исходниках нет, оценка предварительная — допущение). Закрывает только ERD-часть задачи (C3 = 2, ограничитель).",
         links="https://dbdiagram.io"),
    dict(id="K009", name="draw.io / diagrams.net", site="https://www.drawio.com", market="Международный; 🔲 проверить RU-локализацию интерфейса",
         type="Бесплатный графический редактор (open source, Apache 2.0)", product="Web/desktop-редактор диаграмм с библиотеками фигур UML, BPMN, ERD, сетей",
         segment="Все: студенты, аналитики, разработчики, команды", hyp="Косвенный конкурент / заменитель", prio="Высокий", status="Проверено", date=D23,
         found="Запрос «редактор диаграмм онлайн»; SimilarWeb app.diagrams.net 7,8 млн визитов (ЛР2-3)",
         scores=[3, 4, 3, 3, 3, 4, 4, 1, 5, 4], corr="", conf="высокая",
         just="Уверенность: высокая. Главный заменитель: та же аудитория и те же нотации, но ручное рисование без текста и без проверки правил; «No account required. No credit card».",
         links="https://www.drawio.com; https://app.diagrams.net; SimilarWeb (2026-09-23)"),
    dict(id="K010", name="Lucidchart", site="https://www.lucidchart.com", market="Международный, EN (🔲 проверить доступность оплаты из РБ)",
         type="Коммерческий SaaS-редактор диаграмм", product="Графический редактор с совместной работой и AI-генерацией по описанию",
         segment="Бизнес-аналитики, корпоративные команды", hyp="Косвенный конкурент / заменитель", prio="Средний", status="Требует уточнения", date=D17,
         found="Категория GUI-редакторов (00-competitive-analysis.md)",
         scores=[3, 3, 3, 3, 2, 3, 2, 3, 5, 2], corr="", conf="средняя",
         just="Уверенность: средняя. Графический ввод, SaaS-подписка (тарифы 🔲 проверить), сильная совместная работа; для S-02 — альтернатива, для студентов — дорого/избыточно.",
         links="https://www.lucidchart.com"),
    dict(id="K011", name="Camunda Modeler", site="https://camunda.com/download/modeler/", market="Международный, EN",
         type="Desktop-моделлер BPMN (часть BPM-платформы)", product="Редактор BPMN 2.0 / DMN с XML и исполнением в движке Camunda",
         segment="Процессные аналитики, разработчики BPM", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Проверено", date=D17,
         found="Запрос «BPMN редактор» (00-competitive-analysis.md)",
         scores=[3, 2, 2, 2, 3, 2, 2, 1, 4, 3], corr="", conf="средняя",
         just="Уверенность: средняя. Только BPMN (C3 = 2, ограничитель), зато исполнимый; NotaCode сознательно не делает исполнение процессов. Эталон практики для BPMN XML.",
         links="https://camunda.com/download/modeler/"),
    dict(id="K012", name="Ramus (Educational)", site="🔲 уточнить официальный URL (ramussoftware.com?)", market="СНГ, RU",
         type="Настольное CASE-средство IDEF0/DFD", product="Графическое построение моделей IDEF0/DFD (🔲 проверить IDEF3/IDEF1X)",
         segment="Студенты методологических курсов СНГ (S-01), преподаватели (S-04)", hyp="Близкий альтернативный конкурент", prio="Высокий", status="Требует уточнения", date=D17,
         found="Методические указания вузов (🔲 ДОСНЯТЬ ссылку), 00-competitive-analysis.md",
         scores=[4, 4, 3, 4, 2, 2, 5, 1, 2, 2], corr="", conf="средняя",
         just="Уверенность: средняя. Та же задача и аудитория для IDEF-части (курсы СНГ), русскоязычный, но графический ввод, устаревший UX, настольная лицензия (01-project-concept.md, P-04); условия Educational 🔲 проверить.",
         links="🔲 URL сайта Ramus"),
    dict(id="K013", name="Visual Paradigm Online", site="https://online.visual-paradigm.com", market="Международный, EN (🔲 проверить RU)",
         type="SaaS-редактор диаграмм (онлайн-версия CASE Visual Paradigm)", product="Графические UML, BPMN, ERD, DFD и др. (🔲 проверить перечень)",
         segment="Студенты, аналитики, архитекторы (🔲 проверить)", hyp="Косвенный конкурент / заменитель", prio="Средний", status="Требует уточнения", date=TODO_DATE,
         found="Запрос «UML онлайн», «DFD онлайн» (🔲 зафиксировать выдачу)",
         scores=[3, 3, 3, 3, 3, 3, 2, 3, 4, 2], corr="", conf="низкая",
         just="Уверенность: низкая (факты — только категория CASE в 00-competitive-analysis.md; тарифы и нотации 🔲 проверить). Широкий набор нотаций при графическом вводе.",
         links="https://online.visual-paradigm.com"),
    dict(id="K014", name="Miro", site="https://miro.com", market="Международный (🔲 проверить доступность в РБ)",
         type="Онлайн-доска для совместной работы", product="Бесконечная доска с шаблонами схем",
         segment="Команды, фасилитаторы, продакт-менеджеры", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Требует уточнения", date=TODO_DATE,
         found="Смежный рынок (доски), 00-competitive-analysis.md",
         scores=[1, 2, 1, 1, 2, 2, 2, 3, 5, 1], corr="", conf="низкая",
         just="Уверенность: низкая. Схемы как часть совместной работы, без строгих нотаций и текстового ввода; может войти через шаблоны/AI — наблюдать.",
         links="https://miro.com"),
    dict(id="K015", name="Microsoft Visio", site="https://www.microsoft.com/microsoft-365/visio", market="Международный; 🔲 проверить доступность покупки в РБ",
         type="Настольный/облачный редактор диаграмм (Microsoft 365)", product="Графический редактор UML/BPMN/сетей по лицензии",
         segment="Корпоративные пользователи", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Проверено", date=D17,
         found="Архив ЛР5 (заменитель)",
         scores=[3, 2, 3, 2, 1, 2, 2, 1, 5, 2], corr="", conf="средняя",
         just="Уверенность: средняя. Может заменить рисование, но лицензия Microsoft 365 и корпоративный сегмент; для студентов РБ — барьер оплаты.",
         links="https://www.microsoft.com/microsoft-365/visio"),
    dict(id="K016", name="Sparx Enterprise Architect", site="https://sparxsystems.com", market="Международный",
         type="Настольное CASE-средство", product="Моделирование UML/BPMN/SysML, XMI",
         segment="Архитекторы, корпоративные команды", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Проверено", date=D17,
         found="01-project-concept.md (P-04), 00-competitive-analysis.md",
         scores=[3, 2, 3, 2, 1, 1, 3, 1, 5, 2], corr="", conf="средняя",
         just="Уверенность: средняя. Полноценная замена по нотациям, но платная настольная лицензия, без текстового представления и Git (P-04 концепции).",
         links="https://sparxsystems.com"),
    dict(id="K017", name="Graphviz", site="https://graphviz.org", market="Международный, EN",
         type="Open-source движок раскладки графов", product="Язык DOT → графы с автораскладкой",
         segment="Разработчики", hyp="Смежная или потенциальная конкуренция", prio="Низкий", status="Проверено", date=D17,
         found="kursach.md; подборки diagram as code",
         scores=[2, 2, 1, 2, 3, 2, 3, 1, 5, 3], corr="", conf="высокая",
         just="Уверенность: высокая. Инфраструктурный движок графов, нотаций нет; NotaCode экспортирует в DOT — скорее технология, чем конкурент.",
         links="https://graphviz.org"),
    dict(id="K018", name="Kroki", site="https://kroki.io", market="Международный, EN",
         type="Open-source HTTP-агрегатор рендеринга", product="Единый HTTP-API рендера PlantUML/Mermaid/Graphviz/D2 и др.",
         segment="Разработчики", hyp="Не конкурент / объект наблюдения", prio="Низкий", status="Проверено", date=D17,
         found="kursach.md; эпик E8 NotaCode (интеграция Kroki)",
         scores=[2, 2, 1, 1, 3, 1, 3, 1, 3, 3], corr="Не конкурент / объект наблюдения", conf="высокая",
         just="Уверенность: высокая. Корректировка: расчётно «смежная» (1,79), но это агрегатор без собственного пользовательского товара и технологический партнёр NotaCode (E8) — по методике агрегаторы исключаются из конкурентов.",
         links="https://kroki.io"),
    dict(id="K019", name="AI-ассистенты общего назначения (ChatGPT и др.)", site="https://chatgpt.com", market="Международный; 🔲 проверить доступность из РБ",
         type="LLM-чат-боты", product="Генерация кода PlantUML/Mermaid по описанию без рендера и проверки нотации",
         segment="Массово, в т.ч. студенты", hyp="Смежная или потенциальная конкуренция", prio="Средний", status="Проверено", date=D17,
         found="Архив ЛР5 (потенциальные конкуренты); ЛР1 — рост запросов «text to diagram ai»",
         scores=[2, 4, 1, 2, 3, 3, 2, 3, 5, 2], corr="", conf="средняя",
         just="Уверенность: средняя. Сейчас лишь генерирует текст диаграммы (C3 = 1), но имеет огромную студенческую аудиторию — потенциальная угроза при появлении встроенного рендера. Тарифы 🔲 проверить.",
         links="https://chatgpt.com"),
    dict(id="K020", name="Figma / FigJam", site="https://www.figma.com/figjam/", market="Международный",
         type="Дизайн-платформа и онлайн-доска", product="Доска и дизайн-редактор; схемы как побочная функция",
         segment="Дизайнеры, продуктовые команды", hyp="Не конкурент / объект наблюдения", prio="Низкий", status="Исключено", date=TODO_DATE,
         found="Смежный рынок; архив ЛР5 (исключённые альтернативы)",
         scores=[1, 2, 1, 1, 2, 0, 2, 3, 5, 1], corr="", conf="низкая",
         just="Уверенность: низкая. Нет связи с задачей строгих нотаций; оставлена в реестре как объект наблюдения за практиками интерфейса.",
         links="https://www.figma.com/figjam/"),
]

EXCLUDED = [  # п. 10.5 — не вносились в оценку шаблона
    ("MS Word / PowerPoint (рисование фигурами)", "Часто используется студентами для схем в отчётах", "Не специализированы на диаграммах, нет нотаций — поведенческая альтернатива, не товар-конкурент", "Нет"),
    ("Рисование от руки на бумаге/доске", "Встречается в учебной практике", "Поведенческая альтернатива без цифрового товара", "Нет"),
    ("Mermaid Live Editor как отдельная компания", "Найден по запросу «mermaid online»", "Дубль: часть проекта Mermaid (K002)", "Нет — учитывается в K002"),
    ("app.diagrams.net как отдельная компания", "Найден через SimilarWeb (ЛР2-3)", "Дубль: веб-версия draw.io (K009)", "Нет — учитывается в K009"),
    ("DiagramGPT как отдельная компания", "Найден по запросу «AI diagram generator»", "Дубль: продукт Eraser (K004)", "Нет — учитывается в K004"),
    ("Text2Diagram, AI-Diagram (open source)", "kursach.md", "Репозитории-прототипы без пользовательского сервиса и аудитории", "Да — если появится сервис"),
]

# Доказательства (лист 02_Доказательства): id, раздел, URL, факт, тип сигнала, сила 1–5, дата, вывод, достоверность 1–5, скриншот, комментарий
EVIDENCE = [
    ("K001", "Главная", "https://plantuml.com", "Цитата: «No server. No tokens. No tracking»; текстовое описание → UML-диаграмма", "Товар / задача", 5, D17, "Та же задача «текст → UML»", 5, "Рисунки/scr_05_plantuml_main.png", SRC_OLD),
    ("K001", "SimilarWeb", "https://www.similarweb.com/website/plantuml.com/", "≈507 тыс. визитов; Россия — 1-я страна (9,07 %)", "Аудитория / география", 4, D17, "Русскоязычная аудитория пересекается", 4, "Рисунки/scr_06_sw_plantuml.png", SRC_SW),
    ("K001", "Загрузка / интеграции", "https://plantuml.com", "Плагины IntelliJ, VS Code; CLI; расширение браузера для GitHub", "Каналы", 4, D17, "Тот же путь выбора разработчика/студента (IDE, GitHub)", 4, "Рисунки/scr_05_plantuml_main.png", SRC_OLD),
    ("K001", "ЛР1 Google Trends", "https://trends.google.com", "Беларусь в топ-5 стран по интересу к PlantUML; индекс +231 % за 3 г. 3 мес.", "Спрос / поиск", 4, D17, "Постоянно присутствует в тех же запросах — основание корректировки", 4, "из ЛР1", "архив ЛР1/ЛР7"),
    ("K002", "Главная / интеграции", "https://mermaid.js.org", "Нативный рендер в GitHub, GitLab, Notion без плагинов; MIT", "Товар / каналы", 5, D17, "Сильный канал через платформы документации", 5, "Рисунки/scr_07_mermaid_main.png", SRC_OLD),
    ("K002", "Meet the Team", "https://mermaid.js.org", "Кнопки Sponsor (GitHub Sponsors) — донаты вместо тарифов", "Монетизация", 4, D17, "Иная логика выручки (C8 = 1)", 4, "Рисунки/scr_07_mermaid_main.png", SRC_OLD),
    ("K003", "SimilarWeb", "https://www.similarweb.com/website/mermaidchart.com/", "< 20 тыс. визитов; Вьетнам 44,98 %, Индонезия 26,27 %, США 15,1 %", "География", 4, D17, "Русскоязычный сегмент почти не представлен (C7 = 1)", 4, "Рисунки/scr_08_sw_mermaidchart.png", SRC_SW),
    ("K003", "Тарифы", "https://www.mermaidchart.com", "Полный прайсинг не показан без регистрации", "Цена (отрицательный признак)", 2, D17, "C5 = 1, требует проверки", 3, "🔲 Рисунки/scr_09_mermaidchart_pricing.png", SRC_OLD),
    ("K004", "Тарифы", "https://www.eraser.io", "Free 0 (3 AI-диаграммы) / Starter 15 USD год, 20 USD мес (40) / Business 45/60 USD (250) / Enterprise по запросу", "Цена / монетизация", 5, D17, "Та же форма (freemium + подписка), диапазон выше NotaCode Pro", 5, "Рисунки/scr_10_eraser_pricing.png", "архив ЛР2-3, §4"),
    ("K004", "AI / продукт", "https://www.eraser.io/diagramgpt", "AI-генерация диаграмм, MCP-сервер для агентов, синхронизация с git", "Товар", 5, D17, "Ближайший аналог по духу продукта", 4, "Рисунки/scr_11_eraser_ai.png", SRC_CA),
    ("K004", "SimilarWeb", "https://www.similarweb.com/website/eraser.io/", "678,5 тыс. визитов; Индия 40,81 %, США 8,61 %; прямые заходы 49,05 %", "Аудитория / география", 4, D23, "Русскоязычный сегмент не в топ-5 (C7 = 2)", 4, "Рисунки/scr_12_sw_eraser.png", SRC_SW),
    ("K005", "Главная", "https://d2lang.com", "«A modern language that turns text to diagrams»; fiscal sponsor Hack Club", "Товар / монетизация", 5, D23, "Та же идея, без монетизации", 5, "Рисунки/scr_13_d2_main.png", SRC_OLD23),
    ("K005", "FAQ / экспорт", "https://d2lang.com/tour/faq/", "3 движка раскладки; экспорт SVG/PNG/PDF/PPTX/GIF/ASCII; плагины VS Code/Vim", "Товар / каналы", 4, D23, "Архитектурные схемы, UML только class/sequence", 4, "Рисунки/scr_13_d2_main.png", SRC_OLD23),
    ("K006", "Описание продукта", "https://www.gleek.io", "Freemium (есть Free); собственный DSL; UML-подобные, ER, flowchart", "Товар / монетизация", 3, TODO_DATE, "Близкая модель доступа", 3, "🔲 Рисунки/scr_14_gleek_main.png", SRC_KURS),
    ("K006", "Тарифы", "https://www.gleek.io", "🔲 ДОСНЯТЬ: цены и лимиты Free/Pro", "Цена", 1, TODO_DATE, "Уточнить C5", 1, "🔲 Рисунки/scr_15_gleek_pricing.png", TODO),
    ("K007", "Продукт", "https://structurizr.com", "DSL только для модели C4 (Context/Container/Component/Code); экспорт в git-friendly формат", "Товар", 4, D17, "Заменяет малую часть задачи; эталон практики", 4, "🔲 Рисунки/scr_16_structurizr.png", SRC_CA),
    ("K008", "Главная / тарифы", "https://dbdiagram.io", "🔲 ДОСНЯТЬ: DBML, экспорт SQL, Free/Pro, аудитория", "Товар / цена", 1, TODO_DATE, "Оценка предварительная (допущение)", 1, "🔲 Рисунки/scr_17_dbdiagram.png", TODO),
    ("K009", "Главная", "https://www.drawio.com", "«Professional diagramming without the enterprise price tag or privacy compromises»; «No account required. No credit card»", "Товар / цена", 5, D23, "Бесплатный заменитель для той же аудитории", 5, "Рисунки/scr_18_drawio_main.png", SRC_OLD23),
    ("K009", "Интеграции", "https://www.drawio.com", "Google Drive/Docs, SharePoint/OneDrive/Teams, Confluence/Jira, GitHub, VS Code, Notion; «100M+ users»", "Каналы / доверие", 4, D23, "Сильные каналы и зрелость (C9 = 5)", 4, "Рисунки/scr_18_drawio_main.png", SRC_OLD23),
    ("K009", "SimilarWeb", "https://www.similarweb.com/website/app.diagrams.net/", "≈7,8 млн визитов; прямые заходы 68,65 %", "Видимость", 4, D23, "На порядок крупнее нишевых DSL-инструментов", 4, "Рисунки/scr_19_sw_drawio.png", SRC_SW),
    ("K010", "Продукт", "https://www.lucidchart.com", "Коммерческий SaaS; генерация диаграммы по описанию; сильная совместная работа", "Товар / монетизация", 3, D17, "Графический заменитель с AI", 3, "🔲 Рисунки/scr_20_lucidchart.png", SRC_CA),
    ("K011", "Продукт", "https://camunda.com/download/modeler/", "Только BPMN (XML BPMN 2.0), исполнение в движке процессов", "Товар", 4, D17, "Узкая замена BPMN-части", 4, "🔲 Рисунки/scr_21_camunda.png", SRC_CA),
    ("K012", "Продукт", "🔲 URL Ramus", "IDEF0/IDEF1X/IDEF3, используется в курсах методологий СНГ; устаревший UX", "Товар / аудитория", 3, D17, "Совпадение задачи и аудитории по IDEF", 3, "🔲 Рисунки/scr_22_ramus.png", SRC_CA),
    ("K012", "Лицензия", "🔲 URL Ramus", "Настольное приложение с платной лицензией, без текстового представления и Git", "Цена / монетизация", 3, D17, "Иная экономика (C5 = 2, C8 = 1); Educational — проверить", 3, "🔲 Рисунки/scr_22_ramus.png", SRC_CONC),
    ("K013", "Главная / тарифы", "https://online.visual-paradigm.com", "🔲 ДОСНЯТЬ: перечень нотаций, Free-план, цены", "Товар / цена", 1, TODO_DATE, "Оценка предварительная", 1, "🔲 Рисунки/scr_23_vpo.png", TODO),
    ("K014", "Категория", "https://miro.com", "GUI-первая онлайн-доска (категория «Miro/FigJam»)", "Товар", 2, D17, "Смежный рынок", 2, "—", SRC_CA),
    ("K015", "Лицензия", "https://www.microsoft.com/microsoft-365/visio", "Настольная лицензия / Microsoft 365", "Монетизация", 3, D17, "Иная экономика", 3, "—", SRC_OLD),
    ("K016", "Лицензия", "https://sparxsystems.com", "Настольное CASE-средство с платной лицензией, без текста и Git", "Монетизация / товар", 3, D17, "Иная экономика и формат", 3, "—", SRC_CONC),
    ("K017", "Продукт", "https://graphviz.org", "Open source, язык DOT, алгоритмы раскладки (dot/neato)", "Товар", 4, D17, "Технология, а не товар для той же задачи", 4, "—", SRC_KURS),
    ("K018", "Продукт", "https://kroki.io", "HTTP-агрегатор рендеринга разных DSL", "Товар", 4, D17, "Агрегатор без собственного товара — исключить", 4, "—", SRC_KURS),
    ("K019", "Поведение пользователей", "https://chatgpt.com", "AI генерирует код диаграммы, рендер и проверка — вручную вне инструмента", "Заменяемость", 3, D17, "Потенциальная угроза", 3, "—", SRC_OLD),
    ("K020", "Категория", "https://www.figma.com/figjam/", "Онлайн-доска / дизайн-инструмент без нотаций", "Товар", 2, TODO_DATE, "Не конкурент", 2, "—", "архив ЛР5 (исключённые)"),
]

# =====================================================================
# ЧАСТЬ B. Цифровой товар: Левитт и Кано (шаблон Левитт–Кано)
# =====================================================================
RESEARCH_Q = ("Какие функции, условия доступа, сервисные элементы и доказательства качества стали стандартом цифрового товара "
              "в сегменте «инструменты построения диаграмм строгих нотаций (UML, BPMN, ERD, IDEF, DFD)» для русскоязычных студентов "
              "и аналитиков, и какие свойства NotaCode может использовать как преимущество?")
BOUNDS_B = [
    ("Клиентский сегмент", "S-01 студенты техн. специальностей (первично), S-02 аналитики, S-03 архитекторы"),
    ("Задача клиента", "Построить корректную диаграмму строгой нотации и сдать/встроить её в документацию"),
    ("Тип цифрового результата", "Файл диаграммы (SVG/PNG/XML/DSL) + история версий"),
    ("География и язык", "Беларусь, РФ/СНГ; RU-интерфейс желателен, EN-инструменты доступны"),
    ("Ценовой уровень", "0 — 20 USD/мес на пользователя (массовый сегмент)"),
    ("Степень цифровизации", "Полностью автоматический самообслуживаемый сервис (web/desktop)"),
    ("Подгруппы", "DSL-инструменты (C01–C06, C11) и графические редакторы (C07–C10, C12): стандарт по вводу считается внутри подгруппы"),
]

# 12 конкурентов реестра 01_Конкуренты (C01–C12); первые 8 — в 06_Сравнение
LK_COMP = [
    dict(id="C01", name="PlantUML", site="https://plantuml.com", market="Международный, EN (RU-трафик 9,07 %)", type="Прямой конкурент", seg="Разработчики, студенты", mon="Смешанная модель", product="Open-source DSL-рендерер UML (онлайн-сервер, IDE, CLI)", prio=5, status="Проверено", comm="Open source без тарифов; источники дохода на сайте не раскрыты (🔲 проверить)"),
    dict(id="C02", name="Eraser (DiagramGPT)", site="https://www.eraser.io", market="Международный, EN", type="Прямой конкурент", seg="Команды разработки", mon="Бесплатно + платные функции", product="SaaS diagram as code + AI + канвас", prio=5, status="Проверено", comm="Тарифы сняты 2026-09-17 (архив ЛР2-3)"),
    dict(id="C03", name="Mermaid (open source)", site="https://mermaid.js.org", market="Международный, EN", type="Косвенный конкурент", seg="Разработчики, техписатели", mon="Смешанная модель", product="JS-библиотека DSL → SVG, Live Editor", prio=4, status="Проверено", comm="Донаты GitHub Sponsors; коммерческий слой — Mermaid Chart"),
    dict(id="C04", name="D2", site="https://d2lang.com", market="Международный, EN", type="Косвенный конкурент", seg="Разработчики", mon="Смешанная модель", product="Open-source язык + CLI + playground", prio=4, status="Проверено", comm="Fiscal sponsor Hack Club"),
    dict(id="C05", name="Gleek", site="https://www.gleek.io", market="Международный, EN", type="Косвенный конкурент", seg="Разработчики, аналитики", mon="Бесплатно + платные функции", product="SaaS с собственным DSL", prio=3, status="В работе", comm="Тарифы 🔲 проверить"),
    dict(id="C06", name="dbdiagram.io", site="https://dbdiagram.io", market="Международный, EN", type="Заменитель", seg="Разработчики БД, студенты", mon="Бесплатно + платные функции", product="DBML → ERD, экспорт SQL", prio=3, status="В работе", comm="Все факты 🔲 проверить (допущение)"),
    dict(id="C07", name="draw.io / diagrams.net", site="https://www.drawio.com", market="Международный", type="Заменитель", seg="Все сегменты", mon="Смешанная модель", product="Бесплатный графический редактор (web/desktop)", prio=5, status="Проверено", comm="«No account required. No credit card» (2026-09-23)"),
    dict(id="C08", name="Ramus (Educational)", site="🔲 URL Ramus", market="СНГ, RU", type="Косвенный конкурент", seg="Студенты методологических курсов", mon="Корпоративная лицензия", product="Настольное CASE-средство IDEF0/DFD", prio=4, status="В работе", comm="Условия Educational 🔲 проверить"),
    dict(id="C09", name="Lucidchart", site="https://www.lucidchart.com", market="Международный, EN", type="Заменитель", seg="Аналитики, корпоративные команды", mon="Подписка", product="SaaS-редактор с совместной работой и AI", prio=3, status="В работе", comm="Тарифы 🔲 проверить"),
    dict(id="C10", name="Visual Paradigm Online", site="https://online.visual-paradigm.com", market="Международный, EN", type="Заменитель", seg="Студенты, аналитики", mon="Бесплатно + платные функции", product="SaaS-редактор UML/BPMN/ERD/DFD", prio=3, status="В работе", comm="Все факты 🔲 проверить"),
    dict(id="C11", name="Structurizr", site="https://structurizr.com", market="Международный, EN", type="Эталон лучшей практики", seg="Архитекторы ПО", mon="Смешанная модель", product="DSL модели C4, несколько видов из одной модели", prio=2, status="Проверено", comm="Эталон: «одна модель — много представлений»"),
    dict(id="C12", name="Camunda Modeler", site="https://camunda.com/download/modeler/", market="Международный, EN", type="Эталон лучшей практики", seg="Процессные аналитики", mon="Смешанная модель", product="Desktop-моделлер BPMN 2.0 / DMN", prio=2, status="Проверено", comm="Эталон BPMN XML и исполнимости"),
]
LK_IDS = [c["id"] for c in LK_COMP]

# Карта сайтов (02_Карта_сайтов): id, раздел, URL, тип страницы (DV), цель, что извлекаем, дата, скриншот, достоверность, вывод
SITEMAP = [
    ("C01", "Главная", "https://plantuml.com", "Главная страница", "Обещание, аудитория", "Главное обещание, CTA", D17, "Рисунки/scr_05_plantuml_main.png", 5, "«No server. No tokens. No tracking»"),
    ("C01", "Онлайн-сервер", "🔲 URL онлайн-сервера PlantUML", "Функции", "Путь до первого результата", "Шаги до диаграммы", TODO_DATE, "🔲 Рисунки/scr_24_plantuml_server.png", 3, "Результат без регистрации"),
    ("C01", "Документация языка", "🔲 URL раздела UML в документации", "Документация", "База знаний", "Полнота описания синтаксиса", TODO_DATE, "🔲 Рисунки/scr_25_plantuml_docs.png", 4, "Обширная документация"),
    ("C02", "Главная", "https://www.eraser.io", "Главная страница", "Обещание", "Позиционирование AI", D17, "Рисунки/scr_11_eraser_ai.png", 5, "AI for diagrams"),
    ("C02", "DiagramGPT", "https://www.eraser.io/diagramgpt", "Лендинг товара", "AI-функция", "Сценарий генерации", D17, "Рисунки/scr_11_eraser_ai.png", 5, "Генерация по тексту"),
    ("C02", "AI diagrams", "https://www.eraser.io/product/ai-diagrams", "Функции", "Функции", "MCP, git-синхронизация", D17, "Рисунки/scr_11_eraser_ai.png", 4, "Интеграции для агентов"),
    ("C02", "Тарифы", "🔲 URL страницы тарифов Eraser", "Тарифы", "Цена", "Пакеты, лимиты AI", D17, "Рисунки/scr_10_eraser_pricing.png", 5, "4 тарифа, лимит AI 3/40/250/∞"),
    ("C03", "Главная", "https://mermaid.js.org", "Главная страница", "Обещание", "Интеграции с платформами", D17, "Рисунки/scr_07_mermaid_main.png", 5, "Нативно в GitHub/GitLab/Notion"),
    ("C03", "Live Editor", "🔲 URL Mermaid Live Editor", "Функции", "Путь до результата", "Редактор, экспорт", TODO_DATE, "🔲 Рисунки/scr_26_mermaid_live.png", 3, "Результат без регистрации"),
    ("C03", "Meet the Team", "🔲 URL страницы команды", "Отзывы/кейсы", "Доверие, монетизация", "Команда, спонсорство", D17, "Рисунки/scr_07_mermaid_main.png", 4, "GitHub Sponsors"),
    ("C04", "Главная", "https://d2lang.com", "Главная страница", "Обещание", "Позиционирование языка", D23, "Рисунки/scr_13_d2_main.png", 5, "Text to diagrams"),
    ("C04", "FAQ", "https://d2lang.com/tour/faq/", "FAQ", "Условия", "Лицензия, финансирование", D23, "Рисунки/scr_13_d2_main.png", 5, "Hack Club, open source"),
    ("C04", "Playground", "🔲 URL D2 Playground", "Функции", "Демонстрация", "Шаринг, экспорт", TODO_DATE, "🔲 Рисунки/scr_27_d2_playground.png", 3, "Проверка без установки"),
    ("C05", "Главная", "https://www.gleek.io", "Главная страница", "Обещание", "DSL, нотации", TODO_DATE, "🔲 Рисунки/scr_14_gleek_main.png", 3, "🔲 проверить"),
    ("C05", "Тарифы", "🔲 URL тарифов Gleek", "Тарифы", "Цена", "Free/Pro", TODO_DATE, "🔲 Рисунки/scr_15_gleek_pricing.png", 2, "🔲 проверить"),
    ("C05", "Документация", "🔲 URL документации Gleek", "Документация", "База знаний", "Синтаксис", TODO_DATE, "🔲 Рисунки/scr_28_gleek_docs.png", 2, "🔲 проверить"),
    ("C06", "Главная", "https://dbdiagram.io", "Главная страница", "Обещание", "DBML, экспорт", TODO_DATE, "🔲 Рисунки/scr_17_dbdiagram.png", 2, "🔲 проверить"),
    ("C06", "Тарифы", "🔲 URL тарифов dbdiagram", "Тарифы", "Цена", "Free/Pro", TODO_DATE, "🔲 Рисунки/scr_29_dbdiagram_pricing.png", 2, "🔲 проверить"),
    ("C06", "Документация DBML", "🔲 URL документации DBML", "Документация", "База знаний", "Синтаксис", TODO_DATE, "🔲 Рисунки/scr_30_dbml_docs.png", 2, "🔲 проверить"),
    ("C07", "Главная", "https://www.drawio.com", "Главная страница", "Обещание", "Цена, приватность", D23, "Рисунки/scr_18_drawio_main.png", 5, "Без регистрации и карты"),
    ("C07", "Редактор", "https://app.diagrams.net", "Функции", "Путь до результата", "Библиотеки фигур UML/BPMN", D23, "🔲 Рисунки/scr_31_drawio_editor.png", 4, "Результат за 1 шаг"),
    ("C07", "Интеграции", "🔲 URL раздела интеграций draw.io", "Функции", "Каналы", "Google/Microsoft/Atlassian/GitHub", D23, "Рисунки/scr_18_drawio_main.png", 5, "Интеграции как канал"),
    ("C08", "Главная", "🔲 URL Ramus", "Главная страница", "Обещание", "Нотации IDEF/DFD", TODO_DATE, "🔲 Рисунки/scr_22_ramus.png", 2, "🔲 проверить"),
    ("C08", "Загрузка / лицензия", "🔲 URL Ramus", "Тарифы", "Условия", "Educational/коммерческая", TODO_DATE, "🔲 Рисунки/scr_32_ramus_license.png", 2, "🔲 проверить"),
    ("C08", "Документация", "🔲 URL Ramus", "Документация", "База знаний", "Руководство", TODO_DATE, "🔲 Рисунки/scr_33_ramus_docs.png", 2, "🔲 проверить"),
    ("C09", "Главная", "https://www.lucidchart.com", "Главная страница", "Обещание", "Совместная работа, AI", TODO_DATE, "🔲 Рисунки/scr_20_lucidchart.png", 3, "🔲 проверить"),
    ("C09", "Тарифы", "🔲 URL тарифов Lucidchart", "Тарифы", "Цена", "Free/платные", TODO_DATE, "🔲 Рисунки/scr_34_lucid_pricing.png", 2, "🔲 проверить"),
    ("C09", "Регистрация", "🔲 URL регистрации Lucidchart", "Регистрация", "Путь", "Шаги до редактора", TODO_DATE, "🔲 Рисунки/scr_35_lucid_signup.png", 2, "🔲 проверить"),
    ("C10", "Главная", "https://online.visual-paradigm.com", "Главная страница", "Обещание", "Нотации", TODO_DATE, "🔲 Рисунки/scr_23_vpo.png", 2, "🔲 проверить"),
    ("C10", "Тарифы", "🔲 URL тарифов VP Online", "Тарифы", "Цена", "Free/платные", TODO_DATE, "🔲 Рисунки/scr_36_vpo_pricing.png", 2, "🔲 проверить"),
    ("C11", "Главная", "https://structurizr.com", "Главная страница", "Практика", "Модель C4, виды", D17, "🔲 Рисунки/scr_16_structurizr.png", 4, "Одна модель — много видов"),
    ("C11", "Документация DSL", "🔲 URL документации Structurizr DSL", "Документация", "Практика", "Синтаксис workspace/views", TODO_DATE, "🔲 Рисунки/scr_37_structurizr_dsl.png", 3, "Эталон DSL-документации"),
    ("C12", "Загрузка", "https://camunda.com/download/modeler/", "Лендинг товара", "Практика", "BPMN/DMN, платформы", D17, "🔲 Рисунки/scr_21_camunda.png", 4, "Бесплатная загрузка"),
    ("C12", "Документация", "🔲 URL документации Camunda Modeler", "Документация", "Практика", "BPMN XML, линтинг (🔲)", TODO_DATE, "🔲 Рисунки/scr_38_camunda_docs.png", 3, "🔲 проверить наличие линтера"),
]

LEVITT_LEVELS = ["1. Базовая выгода", "2. Родовой товар", "3. Ожидаемый товар", "4. Расширенный товар", "5. Потенциальный товар"]

# 03_Левитт: (id, уровень 1–5, признак, как проявлен, доказательство, H наличие 0/1, I сила 0–3, J понятность 0–3, K влияние 1–5)
LEVITT = [
    ("C01", 1, "Получить UML-диаграмму из текста", "Главная: описание текстом → изображение", "plantuml.com, 2026-09-17", 1, 3, 3, 5),
    ("C01", 2, "Open-source DSL-рендерер (PNG/SVG/ASCII), GPL-3.0", "Раздел загрузки, онлайн-сервер", "kursach.md", 1, 3, 3, 4),
    ("C01", 3, "Работа в браузере и локально, плагины IDE", "«No server. No tokens. No tracking»", "plantuml.com, 2026-09-17", 1, 3, 2, 4),
    ("C01", 4, "Расширение браузера для GitHub, большая база примеров", "Страница расширения, примеры", "архив ЛР5, 2026-09-17", 1, 2, 2, 3),
    ("C01", 5, "Публичная дорожная карта", "Не опубликована, развитие силами сообщества", "отрицательный признак", 0, 0, 1, 2),
    ("C02", 1, "Архитектурные и технические диаграммы по тексту и AI", "«AI for diagrams that matter»", "eraser.io, 2026-09-17", 1, 3, 3, 5),
    ("C02", 2, "SaaS: diagram as code + канвас + markdown-заметки", "Главная, продукт", "eraser.io", 1, 3, 3, 4),
    ("C02", 3, "Интеграции GitHub/Notion/Confluence/VS Code; история версий (90 дней → безлимит)", "Тарифы и продукт", "архив ЛР5, 2026-09-17", 1, 3, 3, 4),
    ("C02", 4, "SSO, API, командные тарифы, usage reporting", "Тарифы Business/Enterprise", "архив ЛР5", 1, 3, 2, 3),
    ("C02", 5, "MCP-сервер для AI-агентов, синхронизация с git", "Страница AI diagrams", "00-competitive-analysis.md", 1, 2, 2, 4),
    ("C03", 1, "Диаграммы прямо в Markdown-документации", "Главная", "mermaid.js.org", 1, 3, 3, 5),
    ("C03", 2, "JS-библиотека рендера (MIT)", "Документация", "kursach.md", 1, 3, 3, 4),
    ("C03", 3, "Нативно в GitHub/GitLab/Notion; Live Editor", "Интеграции", "архив ЛР5", 1, 3, 3, 4),
    ("C03", 4, "Командная работа и AI в Mermaid Chart (тарифы скрыты)", "mermaidchart.com", "архив ЛР5", 1, 1, 1, 3),
    ("C03", 5, "Развитие силами сообщества, GitHub Sponsors", "Meet the Team", "архив ЛР5", 1, 2, 2, 2),
    ("C04", 1, "Быстро перенести ментальную модель на экран", "«A modern language that turns text to diagrams»", "d2lang.com, 2026-09-23", 1, 3, 3, 4),
    ("C04", 2, "Open-source язык + CLI (MPL-2.0)", "FAQ, загрузка", "kursach.md", 1, 3, 3, 4),
    ("C04", 3, "3 движка раскладки, экспорт SVG/PNG/PDF/PPTX/GIF/ASCII", "Документация", "архив ЛР5", 1, 3, 3, 4),
    ("C04", 4, "Плагины VS Code/Vim, imports/variables, playground", "Документация", "архив ЛР5", 1, 3, 2, 3),
    ("C04", 5, "UML class/sequence как направление расширения", "Документация", "архив ЛР5", 1, 2, 2, 2),
    ("C05", 1, "Диаграммы из текста без мыши", "Главная (🔲 проверить формулировку)", "kursach.md", 1, 2, 2, 4),
    ("C05", 2, "Web-SaaS с собственным DSL (freemium)", "Главная", "kursach.md", 1, 2, 2, 3),
    ("C05", 3, "UML-подобные, ER, flowchart в веб-редакторе", "Продукт", "kursach.md", 1, 2, 2, 3),
    ("C05", 4, "Платные функции Pro (🔲 проверить состав)", "Тарифы", "🔲", 1, 1, 1, 2),
    ("C05", 5, "AI-функции", "Публичного подтверждения нет (🔲 проверить)", "🔲", 0, 0, 0, 2),
    ("C06", 1, "Быстро спроектировать схему БД (допущение)", "Главная (🔲)", "🔲", 1, 3, 3, 4),
    ("C06", 2, "Web-SaaS с DSL DBML (допущение)", "Главная (🔲)", "🔲", 1, 3, 3, 4),
    ("C06", 3, "Экспорт SQL DDL/PNG/PDF, шаринг (🔲)", "Продукт (🔲)", "🔲", 1, 2, 2, 4),
    ("C06", 4, "Командная работа, приватные диаграммы (🔲)", "Тарифы (🔲)", "🔲", 1, 2, 2, 3),
    ("C06", 5, "Генерация документации БД (🔲)", "🔲", "🔲", 1, 1, 1, 2),
    ("C07", 1, "Профессиональные диаграммы без корпоративной цены", "Цитата с главной", "drawio.com, 2026-09-23", 1, 3, 3, 5),
    ("C07", 2, "Бесплатный web/desktop-редактор (Apache 2.0)", "Главная", "drawio.com, 2026-09-23", 1, 3, 3, 4),
    ("C07", 3, "Без регистрации; хранение на устройстве/Google Drive/OneDrive/GitHub", "«No account required. No credit card»", "drawio.com, 2026-09-23", 1, 3, 3, 5),
    ("C07", 4, "Интеграции Confluence/Jira/Teams/VS Code/Notion; фигуры UML/BPMN/AWS", "Интеграции", "drawio.com, 2026-09-23", 1, 3, 3, 4),
    ("C07", 5, "Развитие без венчурных денег, совместимость файлов с 2005 г.", "«Not VC-funded»", "drawio.com, 2026-09-23", 1, 2, 2, 2),
    ("C08", 1, "Построение IDEF0/DFD-моделей для учебных задач", "Методические указания вузов (🔲)", "00-competitive-analysis.md", 1, 2, 2, 4),
    ("C08", 2, "Настольное CASE-приложение", "Загрузка (🔲)", "01-project-concept.md", 1, 2, 2, 3),
    ("C08", 3, "Декомпозиция IDEF0, отчёты (🔲)", "🔲", "🔲", 1, 2, 1, 3),
    ("C08", 4, "Русскоязычный интерфейс, применение в курсах СНГ", "🔲", "00-competitive-analysis.md", 1, 2, 2, 3),
    ("C08", 5, "Признаки развития продукта", "Не обнаружены (устаревший UX)", "отрицательный признак", 0, 0, 1, 2),
    ("C09", 1, "Визуальная совместная работа над диаграммами", "Главная (🔲)", "00-competitive-analysis.md", 1, 3, 3, 4),
    ("C09", 2, "SaaS-редактор по подписке", "Тарифы (🔲)", "архив ЛР5", 1, 3, 3, 3),
    ("C09", 3, "Совместное редактирование, шаблоны UML/BPMN/ERD", "Продукт (🔲)", "00-competitive-analysis.md", 1, 3, 2, 4),
    ("C09", 4, "AI-генерация диаграммы по описанию", "Продукт (🔲)", "00-competitive-analysis.md", 1, 2, 2, 3),
    ("C09", 5, "Корпоративные интеграции (🔲)", "🔲", "🔲", 1, 1, 1, 2),
    ("C10", 1, "Диаграммы UML/BPMN/ERD/DFD онлайн (🔲)", "Главная (🔲)", "🔲", 1, 2, 2, 4),
    ("C10", 2, "Web-SaaS (freemium, 🔲)", "Тарифы (🔲)", "🔲", 1, 2, 2, 3),
    ("C10", 3, "Широкий набор нотаций в GUI", "Продукт (🔲)", "00-competitive-analysis.md", 1, 3, 2, 4),
    ("C10", 4, "Шаблоны, совместная работа (🔲)", "🔲", "🔲", 1, 1, 1, 2),
    ("C10", 5, "Связь с desktop Visual Paradigm (🔲)", "🔲", "🔲", 1, 1, 1, 2),
    ("C11", 1, "Архитектура как код (C4)", "Главная", "00-competitive-analysis.md", 1, 3, 3, 4),
    ("C11", 2, "DSL + несколько видов диаграмм из одной модели", "Документация", "архив ЛР5", 1, 3, 3, 4),
    ("C11", 4, "Workspace и views, экспорт в git-friendly формат", "Документация", "00-competitive-analysis.md", 1, 2, 2, 3),
    ("C12", 1, "Моделирование процессов BPMN 2.0", "Страница загрузки", "00-competitive-analysis.md", 1, 3, 3, 4),
    ("C12", 2, "Бесплатный desktop-моделлер", "Страница загрузки", "00-competitive-analysis.md", 1, 3, 3, 3),
    ("C12", 4, "BPMN XML, пригодный к исполнению в движке", "Документация", "00-competitive-analysis.md", 1, 3, 3, 4),
    ("C12", 5, "Исполнение процессов на платформе Camunda", "Сайт платформы", "00-competitive-analysis.md", 1, 2, 2, 3),
]

KANO_CATS = ["Обязательное", "Линейное", "Привлекательное", "Безразличное", "Обратное", "Неясное"]

# 04_Кано: (id, признак, задача, категория, вес 1–5, присутствие по C01..C12 [1/p/0], средняя сила 0–5 или None,
#            решение для NotaCode, левитт-уровень для стандарта)
# «p» — вероятно есть, но в исходниках не подтверждено (🔲 проверить); в счёт входит.
KANO = [
    ("R001", "Текстовый DSL как основной способ описания диаграммы", "Описать диаграмму быстро и воспроизводимо", "Обязательное", 5, "1111110000p0", 4.0, "Включить: ядро NotaCode (E2)", 2),
    ("R002", "Рендер в браузере без установки", "Получить результат сразу", "Обязательное", 5, "11111p101pp0", 4.2, "Включить: PWA (E4, E5)", 3),
    ("R003", "Бесплатный уровень доступа", "Начать без оплаты", "Обязательное", 5, "11111p10pp01", 4.0, "Включить: тариф Free", 3),
    ("R004", "Экспорт SVG/PNG", "Вставить диаграмму в отчёт", "Обязательное", 5, "11111110pppp", 4.3, "Включить в Free (E8)", 3),
    ("R005", "Документация синтаксиса / база знаний", "Разобраться без преподавателя", "Обязательное", 4, "11111p10pp11", 3.8, "Включить: вкладка документации по каждой нотации", 3),
    ("R006", "Шаблоны и примеры диаграмм", "Начать с образца", "Линейное", 3, "11111p10pp10", 3.5, "Включить: шаблоны по нотациям (E3)", 3),
    ("R007", "Подсветка синтаксиса и автодополнение", "Меньше ошибок при наборе", "Линейное", 4, "0111pp000000", 3.0, "Включить: Monaco + грамматика DSL (E4)", 3),
    ("R008", "Интеграция с Git/GitHub", "Хранить модель рядом с кодом и требованиями", "Линейное", 4, "111100100011", 3.8, "Включить: GitHub-синхронизация (E7)", 4),
    ("R009", "Интеграция с облачными хранилищами (Google Drive, OneDrive)", "Хранить файлы привычно", "Линейное", 3, "000000101p00", 3.5, "Включить Google Drive (E7)", 4),
    ("R010", "История версий внутри продукта", "Вернуться к прежнему варианту", "Линейное", 4, "01000p101000", 3.3, "Включить: коммиты, 20 версий в Free (E6)", 4),
    ("R011", "Совместное редактирование в реальном времени", "Работать командой", "Линейное", 3, "010000101p00", 3.6, "Отложить: за границей MVP (C-06)", 4),
    ("R012", "Экспорт в форматы других инструментов", "Не зависеть от одного инструмента", "Привлекательное", 4, "00000p1000p0", 3.0, "Включить: 11 форматов (E8) — дифференциатор", 4),
    ("R013", "AI-генерация диаграммы по тексту", "Сократить время набора", "Привлекательное", 5, "010000001000", 3.5, "Включить как заглушку/BYOK в MVP (E9)", 5),
    ("R014", "API/MCP для AI-агентов", "Встроить в AI-процессы", "Привлекательное", 3, "010000000000", 3.0, "Отложить: после MVP", 5),
    ("R015", "Валидация правил нотации с кодом правила и строкой", "Сдать диаграмму без замечаний", "Привлекательное", 5, "00000000000p", 3.0, "Включить: главный дифференциатор (E3)", 4),
    ("R016", "Профили IDEF0/IDEF1X/IDEF3/DFD", "Выполнить методологические лабораторные", "Привлекательное", 5, "0000000100p0", 2.0, "Включить: волны 1–2 (E3)", 4),
    ("R017", "Сети Петри", "Моделировать параллельные процессы", "Привлекательное", 3, "000000000000", None, "Включить в волну 3 (E3)", 4),
    ("R018", "Кросс-подсветка «строка кода ↔ элемент диаграммы»", "Быстро найти ошибку", "Привлекательное", 4, "000000000000", None, "Включить (E5)", 4),
    ("R019", "Построчный diff версий и откат", "Видеть, что изменилось", "Привлекательное", 4, "000000000000", None, "Включить (E6)", 4),
    ("R020", "Автоматическая раскладка элементов", "Не двигать фигуры вручную", "Обязательное", 4, "1111110000p0", 4.0, "Включить: elkjs (E5)", 3),
    ("R021", "Русскоязычный интерфейс и документация", "Работать на родном языке", "Линейное", 4, "000000p10000", 2.5, "Включить: RU по умолчанию", 3),
    ("R022", "Прозрачные условия доступа и цены на сайте", "Понять, сколько стоит", "Обязательное", 4, "1111p0100001", 4.0, "Включить: страница тарифов Free/Pro", 3),
    ("R023", "Оплата из Беларуси (карта РБ, BYN)", "Оплатить Pro без зарубежной карты", "Неясное", 4, "000000000000", None, "Проверить: провайдер платежей (риск R-24)", 3),
    ("R024", "Офлайн-работа (desktop/CLI/PWA)", "Работать без сети на паре", "Линейное", 3, "100100110011", 3.5, "Включить: офлайн-черновики PWA", 3),
    ("R025", "Кастомные темы и иконки", "Оформить по вкусу", "Безразличное", 2, "11110010p100", 3.0, "Не делать при дефиците ресурсов (только светлая/тёмная тема)", 4),
    ("R026", "SSO и корпоративные функции", "Подключить компанию", "Безразличное", 2, "010000001000", 3.0, "Исключить из MVP", 4),
    ("R027", "Обязательная регистрация до первого результата", "—", "Обратное", 3, "00000p001p00", None, "Не копировать: первый результат без регистрации", 3),
    ("R028", "Исполнение BPMN-процессов", "Запускать процесс", "Безразличное", 2, "000000000001", 4.0, "Сознательно не делать (граница товара)", 5),
    ("R029", "Мобильный редактор", "Правка с телефона", "Неясное", 2, None, None, "Проверить опросом S-01", 5),
]

# 07_Стандарт_рынка — строки на основе требований Кано (id стандарта, id требования, мин. формулировка, что проверить)
STANDARD = [
    ("S001", "R001", "Диаграмма описывается текстом; текст — первичный источник", "Сопоставимость DSL разных нотаций"),
    ("S002", "R002", "Первый результат в браузере без установки", "Время до первого результата (🔲 замерить)"),
    ("S003", "R003", "Есть бесплатный уровень с полноценным редактором", "Лимиты Free у конкурентов"),
    ("S004", "R004", "Экспорт SVG/PNG доступен бесплатно", "—"),
    ("S005", "R005", "Справка по синтаксису каждой нотации", "Полнота документации конкурентов"),
    ("S006", "R020", "Автоматическая раскладка без ручного перемещения", "Качество раскладки сложных схем"),
    ("S007", "R022", "Цена и лимиты раскрыты на сайте до регистрации", "Тарифы Gleek, dbdiagram, Lucidchart, VPO"),
    ("S008", "R006", "Шаблоны для каждой нотации", "—"),
    ("S009", "R008", "Синхронизация с GitHub", "—"),
    ("S010", "R024", "Работа без сети (desktop/PWA)", "—"),
    ("S011", "R010", "История версий в продукте", "—"),
    ("S012", "R011", "Совместное редактирование", "Значимость для S-01"),
    ("S013", "R021", "Русскоязычный интерфейс", "RU-локализация draw.io"),
    ("S014", "R013", "AI-генерация по тексту", "Готовность платить за AI"),
    ("S015", "R015", "Проверка правил нотации", "Наличие линтера у Camunda Modeler"),
    ("S016", "R016", "Профили IDEF/DFD", "Нотации VP Online"),
    ("S017", "R012", "Экспорт в форматы других инструментов", "—"),
    ("S018", "R026", "SSO / enterprise", "—"),
    ("S019", "R027", "Регистрация до первого результата", "Шаги регистрации Lucidchart, VPO"),
]

# 06_Сравнение: 25 критериев (вес из шаблона) и баллы C01..C08
CRIT_P = [
    ("P01", "Сущность товара", "На сайте ясно, какой цифровой товар предлагается", 5),
    ("P02", "Сущность товара", "Целевая задача сформулирована конкретно", 5),
    ("P03", "Сущность товара", "Граница товара отделена от консультаций и обещаний", 4),
    ("P04", "Функциональность", "Ключевые функции раскрыты через сценарии", 5),
    ("P05", "Функциональность", "Есть демонстрация интерфейса или результата", 4),
    ("P06", "Функциональность", "Есть пробный доступ, демо или понятный вход", 4),
    ("P07", "Ценность", "Показана измеримая выгода", 5),
    ("P08", "Ценность", "Преимущества связаны с функциями", 5),
    ("P09", "Ценность", "Есть сравнение тарифов или уровней доступа", 4),
    ("P10", "Доверие", "Отзывы, кейсы, примеры клиентов", 4),
    ("P11", "Доверие", "Безопасность, данные, юрусловия, гарантии", 4),
    ("P12", "Доверие", "Прозрачные контакты, команда", 3),
    ("P13", "Интерфейс", "Путь к регистрации/результату короткий", 5),
    ("P14", "Интерфейс", "CTA соответствует стадии выбора", 4),
    ("P15", "Интерфейс", "Навигация помогает понять состав товара", 4),
    ("P16", "Тарифы и оплата", "Цена, состав тарифа и ограничения раскрыты", 5),
    ("P17", "Тарифы и оплата", "Оплата, возврат, пробный период объяснены", 4),
    ("P18", "Сопровождение", "Поддержка и обучение встроены в предложение", 3),
    ("P19", "Сопровождение", "База знаний, FAQ, документация", 3),
    ("P20", "Расширение товара", "Интеграции, автоматизация, аналитика", 4),
    ("P21", "Расширение товара", "Признаки развития: обновления, новые функции", 3),
    ("P22", "Дифференциация", "Уникальные или редкие признаки", 5),
    ("P23", "Дифференциация", "Отличия объяснены через ценность", 5),
    ("P24", "Рыночный стандарт", "Закрыты обязательные требования категории", 5),
    ("P25", "Рыночный стандарт", "Не уступает большинству по базовым ожиданиям", 5),
]
#                 PlantUML Eraser Mermaid D2 Gleek dbdiag drawio Ramus
P_SCORES = {
    "P01": [4, 5, 4, 5, 4, 5, 5, 3],
    "P02": [4, 4, 4, 4, 3, 4, 4, 3],
    "P03": [4, 3, 4, 4, 3, 4, 4, 2],
    "P04": [3, 5, 3, 4, 3, 4, 4, 2],
    "P05": [4, 5, 5, 5, 4, 5, 5, 2],
    "P06": [5, 5, 5, 5, 4, 5, 5, 3],
    "P07": [1, 3, 1, 2, 2, 2, 2, 1],
    "P08": [3, 4, 3, 4, 3, 4, 4, 2],
    "P09": [0, 5, 0, 0, 3, 3, 1, 1],
    "P10": [2, 3, 3, 2, 2, 3, 4, 1],
    "P11": [3, 4, 2, 2, 2, 3, 5, 1],
    "P12": [2, 4, 4, 3, 3, 3, 4, 2],
    "P13": [5, 4, 5, 5, 4, 4, 5, 2],
    "P14": [3, 5, 4, 4, 4, 4, 4, 2],
    "P15": [3, 4, 4, 4, 3, 4, 4, 2],
    "P16": [4, 5, 4, 4, 3, 3, 5, 2],
    "P17": [2, 4, 2, 2, 2, 2, 4, 1],
    "P18": [3, 4, 3, 4, 2, 2, 3, 2],
    "P19": [5, 4, 5, 5, 3, 4, 4, 3],
    "P20": [4, 5, 5, 4, 2, 3, 5, 1],
    "P21": [3, 5, 4, 3, 2, 3, 3, 1],
    "P22": [3, 5, 3, 4, 2, 3, 3, 4],
    "P23": [3, 4, 3, 4, 2, 3, 5, 2],
    "P24": [4, 5, 4, 4, 3, 4, 5, 2],
    "P25": [4, 5, 4, 4, 3, 4, 5, 2],
}

# 08_Преимущества: id, ID конкурента, признак, доказательство, левитт, кано, тип, редкость, сила, ценность, кач. доказательства (1–5),
#   как использовать; + 7 критериев методички 0–5 (ценность, отличимость, доказанность, связь с результатом, устойчивость, коммуницируемость, сегмент)
ADV = [
    ("A001", "C02", "AI-генерация диаграмм + MCP-сервер для агентов", "https://www.eraser.io/diagramgpt (2026-09-17)", 5, "Привлекательное", "Интеграционное", 5, 4, 4, 4,
     "AI-режимы E9 с самопроверкой кода валидатором; MCP — после MVP", (4, 5, 4, 4, 3, 5, 3)),
    ("A002", "C07", "Бесплатно, без регистрации и без карты", "https://www.drawio.com (2026-09-23)", 3, "Обязательное", "Экономическое", 3, 5, 5, 5,
     "Первый результат в NotaCode — без регистрации; регистрация только для облака", (5, 3, 5, 4, 3, 5, 5)),
    ("A003", "C03", "Нативный рендер в GitHub/GitLab/Notion", "https://mermaid.js.org (2026-09-17)", 4, "Линейное", "Сетевой эффект", 5, 5, 4, 5,
     "Экспорт в Mermaid (E8) — встраивание результата в платформы, где Mermaid уже стандарт", (4, 5, 5, 4, 5, 4, 3)),
    ("A004", "C01", "Работа локально и в IDE, «No tracking»", "https://plantuml.com (2026-09-17)", 3, "Линейное", "Доверительное", 3, 4, 4, 5,
     "Офлайн-черновики PWA, прозрачная политика данных", (4, 3, 5, 3, 3, 4, 4)),
    ("A005", "C04", "3 движка раскладки и экспорт в 6 форматов", "https://d2lang.com/tour/faq/ (2026-09-23)", 3, "Линейное", "Функциональное", 3, 4, 3, 4,
     "Качество раскладки elkjs; экспорт шире SVG/PNG в Pro", (3, 3, 4, 3, 3, 4, 3)),
    ("A006", "C07", "Интеграции Google/Microsoft/Atlassian/GitHub как канал", "https://www.drawio.com (2026-09-23)", 3, "Линейное", "Интеграционное", 3, 5, 4, 4,
     "GitHub + Google Drive в MVP (E7); Confluence — гипотеза для S-02", (4, 3, 4, 3, 4, 4, 3)),
    ("A007", "C08", "IDEF0/DFD на русском в учебных курсах СНГ", "🔲 URL Ramus; 00-competitive-analysis.md", 4, "Привлекательное", "Контентное", 4, 2, 4, 2,
     "Шаблоны учебных заданий IDEF0/DFD на русском — перехват сегмента S-01", (4, 4, 2, 4, 2, 3, 5)),
    ("A008", "C02", "История версий, градуированная по тарифам (90 дней → безлимит)", "архив ЛР5, 2026-09-17", 3, "Линейное", "Экономическое", 3, 3, 3, 4,
     "20 версий в Free, полная история в Pro — лимит как стимул оплаты", (3, 3, 4, 3, 2, 4, 3)),
    ("A009", "C11", "Одна модель — несколько представлений (C4)", "https://structurizr.com (2026-09-17)", 4, "Привлекательное", "Функциональное", 4, 4, 3, 3,
     "Декомпозиция IDEF0 A-0 → A0 → A1 из одной модели", (3, 4, 3, 3, 3, 3, 3)),
    ("A010", "C12", "Исполнимый BPMN XML", "https://camunda.com/download/modeler/ (2026-09-17)", 4, "Безразличное", "Функциональное", 4, 4, 2, 3,
     "Не копировать: только корректный экспорт BPMN XML", (2, 4, 3, 2, 4, 3, 1)),
    ("A011", "C02", "Цена 15–20 USD/участник/мес", "архив ЛР2-3, 2026-09-17", 1, "Обратное", "Экономическое", 1, 2, 2, 5,
     "Не копировать: для S-01 цена Pro 4 USD", (1, 1, 5, 1, 1, 3, 1)),
]

MATRIX_05 = [  # 05_Матрица_товара C..L для строк 7..18 (в порядке листа)
    ("Все прямые конкуренты решают «текст → диаграмма» для UML/архитектуры; строгие нотации IDEF/DFD — только графически (Ramus, VPO)",
     "03_Левитт: C01–C04 ур.1; C08 ур.1", "Обязательно включить", "UML class/sequence, ERD, BPMN, IDEF0, DFD из текста с проверкой правил", "Все 14 типов UML, IDEF1X/IDEF3, сети Петри", "Без IDEF/DFD NotaCode = ещё один PlantUML", "Включить", "Базовая выгода — корректность, а не просто картинка"),
    ("DSL-инструменты ориентированы на разработчиков; студенты методологических курсов обслуживаются устаревшим Ramus", "01_Конкуренты: сегменты; SimilarWeb", "Обязательно включить", "S-01 (Free) — студенты", "S-02/S-03 (Pro), S-04 — шаблоны для преподавателей", "Размытое позиционирование", "Включить", "Первичный сегмент — S-01"),
    ("Родовой товар категории — web-сервис/библиотека DSL → SVG; у графических — web-редактор", "03_Левитт ур.2", "Обязательно включить", "PWA web-IDE (Monaco + SVG-холст)", "Офлайн-режим, desktop через PWA", "—", "Включить", "Родовой товар — web-IDE"),
    ("Норма: DSL, автораскладка, экспорт SVG/PNG, документация, шаблоны (доли 0,58–0,92)", "04_Кано R001–R006, R020", "Обязательно включить", "DSL, elkjs, SVG/PNG, справка, шаблоны", "Автодополнение, quick fix", "Ниже рыночного стандарта", "Включить", "Минимальный стандарт"),
    ("Первый результат без регистрации у PlantUML, Mermaid, D2, draw.io", "04_Кано R002, R027; 06 P13", "Обязательно включить", "Редактор и холст без регистрации, облако — после входа", "GitHub/Google вход, PWA офлайн", "Отток на шаге регистрации", "Включить", "Регистрация — только для сохранения в облако"),
    ("Прямые DSL-конкуренты бесплатны; платный — Eraser 15–20 USD", "01_Конкуренты G; архив ЛР2-3", "Обязательно включить", "Free: 5 проектов, 20 версий, SVG/PNG", "Pro 4 USD/мес или 40 USD/год; оплата из РБ", "Невозможность оплаты из РБ (R-24)", "Проверить", "Платёжный провайдер РБ 🔲"),
    ("Поддержка — документация и сообщество (форумы, Discord)", "04_Кано R005; 06 P18–P19", "Включить при наличии ресурсов", "Вкладка документации, FAQ", "Чат/сообщество, видеоуроки", "Рост вопросов к автору", "Включить", "Документация по каждой нотации"),
    ("Расширение у лидеров — AI (Eraser, Lucidchart) и интеграции", "04_Кано R008, R013; 08 A001", "Использовать как преимущество", "AI-заглушка/BYOK, GitHub", "AI с самопроверкой валидатором, MCP", "Отставание от Eraser по AI", "Включить", "AI — после валидатора"),
    ("Доказательства — открытый код, число пользователей (draw.io 100M+), цитаты; измеримой выгоды почти никто не показывает (P07 ≤ 3)", "06 P07, P10", "Использовать как преимущество", "Примеры «до/после», галерея диаграмм", "Замер времени vs draw.io (BO-01 ≥ 40 %)", "Недоверие новому продукту", "Проверить", "Доказательство — замер времени"),
    ("Потенциал рынка — AI-агенты, git-синхронизация", "03_Левитт ур.5", "Проверить гипотезу", "—", "«Изображение → DSL», MCP, маркетплейс шаблонов", "—", "Отложить", "После MVP"),
    ("Граница: исполнение BPMN (Camunda), совместная работа, enterprise SSO — вне товара", "04_Кано R011, R026, R028", "Исключить", "—", "—", "Раздувание объёма MVP", "Исключить", "Сознательные исключения"),
    ("Цифровой товар NotaCode — это web-IDE «diagram as code» для студентов и аналитиков, которая превращает текст в корректную диаграмму строгой нотации (UML, BPMN, ERD, IDEF0/1X/3, DFD, сети Петри) за счёт единого DSL, валидации правил нотации с кодом ошибки и строкой, автораскладки, версий с diff и экспорта в форматы аналогов",
     "Синтез листов 03–08", "Обязательно включить", "Минимальный: DSL + 5 нотаций волны 1 + валидация + SVG/PNG + Free", "Расширенный: все нотации, AI, 11 форматов экспорта, GitHub/Drive", "—", "Включить", "Формула товара"),
]

DASHBOARD_B = [
    "Минимум: текстовый DSL, рендер в браузере без установки и регистрации, автораскладка, экспорт SVG/PNG, документация и шаблоны по каждой нотации, прозрачный Free-тариф.",
    "Нельзя игнорировать: бесплатный вход (у 10 из 12), экспорт SVG/PNG (11 из 12), документацию (11 из 12), интеграцию с Git (7 из 12) — без них NotaCode ниже нормы.",
    "Дифференциация: валидация правил нотации, профили IDEF0/IDEF1X/IDEF3/DFD и сети Петри, кросс-подсветка «строка ↔ элемент», построчный diff и экспорт в форматы аналогов — у 0–3 из 12 конкурентов.",
    "Лучшие практики: вход без регистрации (draw.io), тарифы с лимитом AI (Eraser), нативное встраивание в GitHub (Mermaid), «одна модель — несколько видов» (Structurizr).",
    "Риски копирования: цена Eraser (15–20 USD) неприемлема для S-01; исполнение BPMN (Camunda) и SSO — вне задачи; AI без проверки нотации воспроизводит ошибки ChatGPT.",
]

# =====================================================================
# ЧАСТЬ C. Бизнес-модель (шаблон 13 листов)
# =====================================================================
BM_PASSPORT = [
    ("Цель исследования", "Выбрать элементы бизнес-модели NotaCode (freemium Free/Pro 4 USD) для запуска MVP к 06.12.2026: стандарт рынка, модель доходов, путь первой покупки"),
    ("Рынок", "Инструменты построения диаграмм (diagram as code и графические редакторы), Беларусь/СНГ, B2C (студенты) + B2B-лайт (аналитики, архитекторы)"),
    ("Единица анализа", "Конкретное предложение: веб-/настольный доступ к редактору диаграмм и его платные расширения"),
    ("Период сбора данных", "2026-09-17 — 2026-09-25 (архив) + 🔲 ДОСНЯТЬ: повторный просмотр сайтов (даты фиксировать)"),
    ("Список конкурентов", "8: PlantUML, Eraser, Mermaid, D2, Gleek, draw.io, Ramus, Lucidchart (прямые и косвенные из части A)"),
    ("Источники", "Главная, продукт, тарифы, регистрация/демо, документация, интеграции, FAQ/лицензия, команда/спонсоры; вакансии — 🔲"),
]

BM_COMP = [  # ID как в части B
    dict(id="C01", name="PlantUML", site="https://plantuml.com", seg="Разработчики, студенты", type="прямой", geo="Международный (RU-трафик 9,07 %)", prio="высокий", status="завершен", date=D17, note="Open source (GPL-3.0)"),
    dict(id="C02", name="Eraser (DiagramGPT)", site="https://www.eraser.io", seg="Команды разработки", type="прямой", geo="Международный (Индия 40,81 %)", prio="высокий", status="завершен", date=D23, note="Freemium + подписка по местам"),
    dict(id="C03", name="Mermaid (open source)", site="https://mermaid.js.org", seg="Разработчики, техписатели", type="косвенный", geo="Международный", prio="высокий", status="завершен", date=D17, note="MIT, GitHub Sponsors"),
    dict(id="C04", name="D2", site="https://d2lang.com", seg="Разработчики", type="косвенный", geo="Международный", prio="средний", status="завершен", date=D23, note="MPL-2.0, Hack Club"),
    dict(id="C05", name="Gleek", site="https://www.gleek.io", seg="Разработчики, аналитики", type="косвенный", geo="Международный", prio="средний", status="требует проверки", date=TODO_DATE, note="Freemium — 🔲 проверить"),
    dict(id="C07", name="draw.io / diagrams.net", site="https://www.drawio.com", seg="Все сегменты", type="заменитель", geo="Международный", prio="высокий", status="завершен", date=D23, note="Apache 2.0, «Not VC-funded»"),
    dict(id="C08", name="Ramus (Educational)", site="🔲 URL Ramus", seg="Студенты СНГ", type="косвенный", geo="СНГ", prio="средний", status="требует проверки", date=D17, note="Настольная лицензия"),
    dict(id="C09", name="Lucidchart", site="https://www.lucidchart.com", seg="Аналитики, корпоративные команды", type="заменитель", geo="Международный", prio="средний", status="требует проверки", date=D17, note="SaaS-подписка"),
]
BM_IDS = [c["id"] for c in BM_COMP]

# 03_Факты_сайта: (ID конкурента, раздел, URL, блок, факт, элемент BM, доказательность 1–5, дата, цитата/скриншот, вывод)
FACTS = [
    ("C01", "Главная", "https://plantuml.com", "Первый экран", "«No server. No tokens. No tracking»", "ценностное предложение", 5, D17, "цитата; Рисунки/scr_05", "Ценность — приватность и локальная работа"),
    ("C01", "Лицензия", "https://plantuml.com", "Загрузка", "GPL-3.0, open source", "потоки доходов", 4, D17, "kursach.md", "Прямой выручки от товара нет"),
    ("C01", "Интеграции", "https://plantuml.com", "IDE/CI", "Плагины IntelliJ, VS Code; расширение для GitHub", "каналы", 4, D17, "архив ЛР5", "Канал — экосистема IDE"),
    ("C01", "SimilarWeb", "https://www.similarweb.com/website/plantuml.com/", "Трафик", "≈507 тыс. визитов; РФ 9,07 %", "целевые сегменты", 4, D17, "Рисунки/scr_06", "Русскоязычная аудитория"),
    ("C01", "Тарифы", "https://plantuml.com", "—", "Страницы тарифов нет (отрицательный признак)", "потоки доходов", 3, D17, "—", "Модель не коммерческая"),
    ("C02", "Тарифы", "🔲 URL тарифов Eraser", "Тарифная сетка", "Free/Starter 15–20/Business 45–60 USD/Enterprise по запросу; лимиты AI 3/40/250/∞", "потоки доходов", 5, D17, "Рисунки/scr_10", "Подписка по местам с лимитом AI"),
    ("C02", "Тарифы", "🔲 URL тарифов Eraser", "Enterprise", "SSO, usage reporting на старших тарифах", "отношения с клиентами", 4, D17, "архив ЛР5", "Переход к корпоративным продажам"),
    ("C02", "Продукт", "https://www.eraser.io/product/ai-diagrams", "AI", "AI-генерация, MCP-сервер для агентов, git-синхронизация", "цифровой товар", 5, D17, "00-competitive-analysis.md", "Ядро — AI + DSL"),
    ("C02", "Интеграции", "https://www.eraser.io", "Интеграции", "GitHub, Notion, Confluence, VS Code", "ключевые партнеры", 4, D17, "архив ЛР5", "Платформенные партнёрства"),
    ("C02", "SimilarWeb", "https://www.similarweb.com/website/eraser.io/", "Трафик", "678,5 тыс. визитов; прямые заходы 49,05 %", "каналы", 4, D23, "Рисунки/scr_12", "PLG + бренд"),
    ("C03", "Главная", "https://mermaid.js.org", "Интеграции", "Нативный рендер в GitHub/GitLab/Notion", "данные и сетевые эффекты", 5, D17, "Рисунки/scr_07", "Стандарт де-факто — сетевой эффект платформ"),
    ("C03", "Команда", "🔲 URL Meet the Team", "Спонсоры", "Кнопки GitHub Sponsors", "потоки доходов", 4, D17, "архив ЛР5", "Донаты"),
    ("C03", "Лицензия", "https://mermaid.js.org", "—", "MIT", "ключевые ресурсы", 4, D17, "kursach.md", "Открытый код — ресурс сообщества"),
    ("C03", "Коммерческий слой", "https://www.mermaidchart.com", "Тарифы", "Прайсинг Mermaid Chart скрыт без регистрации", "потоки доходов", 2, D17, "архив ЛР5", "Монетизация через отдельный SaaS"),
    ("C04", "Главная", "https://d2lang.com", "Первый экран", "«A modern language that turns text to diagrams»", "ценностное предложение", 5, D23, "Рисунки/scr_13", "Ценность — скорость"),
    ("C04", "FAQ", "https://d2lang.com/tour/faq/", "Финансирование", "Fiscal sponsor Hack Club, не венчурный", "структура затрат", 4, D23, "архив ЛР5", "Низкие затраты, донаты"),
    ("C04", "Документация", "https://d2lang.com/tour/faq/", "Экспорт", "SVG/PNG/PDF/PPTX/GIF/ASCII; 3 движка раскладки", "ключевые ресурсы", 4, D23, "архив ЛР5", "Технология раскладки"),
    ("C04", "Каналы", "https://d2lang.com", "Сообщество", "Homebrew/Winget, GitHub, Discord, playground", "каналы", 4, D23, "архив ЛР5", "Open-source дистрибуция"),
    ("C05", "Главная", "https://www.gleek.io", "Продукт", "Собственный DSL, UML-подобные, ER, flowchart", "цифровой товар", 3, TODO_DATE, "kursach.md", "🔲 проверить"),
    ("C05", "Тарифы", "🔲 URL тарифов Gleek", "Тарифы", "Freemium, есть Free (цены 🔲)", "потоки доходов", 2, TODO_DATE, "kursach.md", "🔲 проверить цены"),
    ("C07", "Главная", "https://www.drawio.com", "Первый экран", "«Professional diagramming without the enterprise price tag or privacy compromises»", "ценностное предложение", 5, D23, "Рисунки/scr_18", "Ценность — бесплатно и приватно"),
    ("C07", "Главная", "https://www.drawio.com", "Доступ", "«No account required. No credit card»", "каналы", 5, D23, "Рисунки/scr_18", "Нулевой барьер входа"),
    ("C07", "О компании", "https://www.drawio.com", "Модель", "«No artificial scarcity, no Enterprise tier for SSO»; «Not VC-funded»", "потоки доходов", 4, D23, "архив ЛР5", "Нет платного тарифа редактора"),
    ("C07", "Интеграции", "https://www.drawio.com", "Интеграции", "Google, Microsoft, Atlassian, GitHub, VS Code, Notion; плагины Atlassian Marketplace", "ключевые партнеры", 5, D23, "архив ЛР5", "Интеграции — канал и партнёрства (платность плагинов 🔲)"),
    ("C07", "О компании", "https://www.drawio.com", "Масштаб", "«100M+ users worldwide», история с 2005 г.", "масштабирование", 3, D23, "архив ЛР5", "Заявление компании, не проверено"),
    ("C08", "Продукт", "🔲 URL Ramus", "Нотации", "IDEF0/IDEF1X/IDEF3, курсы СНГ, устаревший UX", "цифровой товар", 3, D17, "00-competitive-analysis.md", "Нишевой учебный товар"),
    ("C08", "Лицензия", "🔲 URL Ramus", "Лицензия", "Настольное приложение, платная лицензия, без текста и Git", "потоки доходов", 3, D17, "01-project-concept.md", "Лицензионная модель"),
    ("C09", "Продукт", "https://www.lucidchart.com", "Функции", "Совместная работа — сильная сторона; AI-генерация по описанию", "цифровой товар", 3, D17, "00-competitive-analysis.md", "Ценность — команда"),
    ("C09", "Тарифы", "🔲 URL тарифов Lucidchart", "Тарифы", "SaaS-подписка (цены 🔲)", "потоки доходов", 2, D17, "архив ЛР5", "🔲 проверить"),
]

# 04_Бизнес_модель: 12 блоков (сегменты … данные/сети, главная гипотеза); None — не заполнено (нет данных)
BM_PROFILE = {
    "C01": ["Разработчики, студенты, техписатели; РФ — 1-я страна трафика", "Бесплатные UML-диаграммы из текста, локально и без слежки", "DSL + рендер (онлайн-сервер, jar, плагины IDE)", "Поиск, IDE-плагины, GitHub-расширение", "Самообслуживание, сообщество", "Нет прямой выручки (open source)", "Язык PlantUML, документация, сообщество", "Развитие языка и рендера силами автора и сообщества", "IDE и CI-инструменты (встраивание)", "Низкие: хостинг сервера, поддержка кода", "Сетевой эффект — встроенность в инструменты", "Некоммерческий стандарт де-факто для UML-as-code"],
    "C02": ["Команды разработки от индивидуальных до Enterprise", "AI + diagram as code в одном рабочем пространстве", "SaaS: DSL, AI, канвас, заметки", "PLG через Free (3 AI-диаграммы), прямые заходы 49 %", "Самообслуживание → корпоративные продажи (SSO)", "Подписка по местам 15–60 USD; Enterprise по запросу", "Собственная AI-генерация, интеграции", "Разработка AI, продажи Enterprise", "GitHub, Notion, Confluence, VS Code", "Разработка AI, инференс, продажи", "История пользователя, git-синхронизация", "Freemium с апгрейдом по лимиту AI-кредитов"],
    "C03": ["Разработчики и техписатели в GitHub/GitLab/Notion", "Диаграммы в Markdown без плагинов", "JS-библиотека + Live Editor", "Нативно в платформах документации", "Сообщество", "Донаты (GitHub Sponsors) + отдельный SaaS Mermaid Chart", "Открытый код MIT, сообщество", "Развитие библиотеки", "GitHub, GitLab, Notion", "Низкие", "Сетевой эффект платформ-хостов", "Стандарт де-факто, монетизируемый через SaaS-надстройку"],
    "C04": ["Разработчики", "Быстрый современный язык диаграмм", "DSL + CLI + playground", "Open-source дистрибуция, Discord", "Сообщество", "Нет (fiscal sponsor Hack Club)", "Язык D2, 3 движка раскладки", "Развитие языка", "Плагины VS Code/Vim", "Низкие, не венчурные", None, "Некоммерческий проект с технологическим отличием в раскладке"],
    "C05": ["Разработчики, аналитики (🔲)", "Диаграммы из текста в вебе (🔲)", "SaaS с DSL", None, None, "Freemium (🔲 цены)", None, None, None, None, None, "Freemium SaaS — модель, близкая к NotaCode (🔲)"],
    "C07": ["Все: от индивидуальных пользователей до команд", "Профессионально, бесплатно, приватно", "Web/desktop-редактор, библиотеки фигур", "Прямые заходы 68,65 %, интеграции как канал", "Самообслуживание", "Нет платного тарифа редактора; плагины Atlassian Marketplace (🔲 платность)", "20-летняя история, совместимость файлов", "Развитие редактора и интеграций", "Google, Microsoft, Atlassian, GitHub", "Разработка; инфраструктура — хранилище у пользователя", "Хранение у пользователя, «100M+ users» (заявление)", "Бесплатный стандарт GUI-редактора, монетизация через интеграции (🔲)"],
    "C08": ["Студенты методологических курсов СНГ", "IDEF0/DFD-моделирование", "Настольное CASE-приложение", None, None, "Лицензия (Educational — 🔲)", None, None, None, None, None, "Нишевая лицензионная модель для вузов (🔲)"],
    "C09": ["Аналитики, корпоративные команды", "Совместная визуальная работа", "SaaS-редактор + AI", "🔲", None, "Подписка (🔲 цены)", None, None, None, None, None, "SaaS-подписка для команд (🔲)"],
}

BM_PRODUCT = {  # 05_Товар_ценность C..M
    "C01": ["Рендер UML из текста", "Онлайн-сервер и jar", "Плагины IDE, документация", "GitHub-расширение", "Не заявлено", "Ручное рисование UML", "Диаграмма в документации за минуты", "Документировать ПО", "Бесплатен, работает локально", "Главная, загрузка", "Товар — ядро некоммерческой экосистемы"],
    "C02": ["Диаграмма из текста/AI", "Free: канвас + DSL + 3 AI", "Интеграции, история версий", "SSO, API, usage reporting", "MCP-агенты, git-синхронизация", "Долго рисовать архитектуру", "Диаграммы для документации команды", "Документировать архитектуру", "AI + DSL + канвас", "Тарифы, AI diagrams", "Товар — драйвер подписки"],
    "C03": ["Диаграмма в Markdown", "Библиотека + Live Editor", "Встроенность в GitHub/GitLab/Notion", "Mermaid Chart для команд", "Развитие сообществом", "Диаграммы отдельно от документации", "Диаграмма живёт в репозитории", "Документировать в Markdown", "Нативность в платформах", "Главная, интеграции", "Товар — стандарт, монетизируемый надстройкой"],
    "C04": ["Текст → диаграмма", "CLI + playground", "Экспорт 6 форматов", "Плагины редакторов", "UML class/sequence", "Некрасивые авто-схемы", "Красивая раскладка", "Архитектурные схемы", "3 движка раскладки", "Главная, FAQ", "Товар — технология раскладки"],
    "C05": ["Текст → диаграмма (🔲)", "Free-план (🔲)", "🔲", "Pro (🔲)", "🔲", "🔲", "🔲", "🔲", "Собственный DSL", "Главная (🔲)", "🔲"],
    "C07": ["Нарисовать любую диаграмму", "Бесплатный редактор", "Без регистрации, хранение у пользователя", "Интеграции с офисными платформами", "Без VC — стабильность", "Дорогие корпоративные редакторы", "Профессиональная схема бесплатно", "Нарисовать схему", "Цена 0 и приватность", "Главная", "Товар — бесплатный стандарт GUI"],
    "C08": ["IDEF0/DFD-модель", "Настольное приложение", "🔲", "🔲", "Не обнаружено", "Нет инструмента IDEF", "Сдать учебную модель", "Лабораторные по IDEF", "IDEF на русском", "🔲", "Нишевой учебный товар"],
    "C09": ["Совместная диаграмма", "SaaS-редактор", "Шаблоны, совместная работа", "AI-генерация", "🔲", "Разрозненная командная работа", "Схема в команде", "Командная работа", "Совместное редактирование", "🔲", "Товар — командная подписка"],
}

BM_MONET = {  # 06_Монетизация C..N (C — модель из списка, L,M 1–5)
    "C01": ["условно-бесплатная модель", "нет (бесплатно)", "да, полностью", "не требуется", "нет", "—", "нет", "—", "нет", 5, 1, "Выручки от товара нет; конкурирует ценой 0"],
    "C02": ["подписка", "да", "да, Free", "Free-тариф", "4 уровня", "участник/мес", "лимит AI 3/40/250/∞, история 90 дней → безлимит", "мес/год (скидка за год)", "Enterprise по запросу, SSO", 5, 4, "Подписка по местам с лимитом AI"],
    "C03": ["смешанная", "нет (Mermaid Chart — скрыто)", "да", "не требуется", "нет", "—", "нет", "—", "нет", 2, 2, "Донаты + отдельный SaaS"],
    "C04": ["условно-бесплатная модель", "нет", "да", "не требуется", "нет", "—", "нет", "—", "нет", 5, 1, "Некоммерческий"],
    "C05": ["условно-бесплатная модель", "🔲", "да (kursach.md)", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", 2, 3, "🔲 проверить"],
    "C07": ["условно-бесплатная модель", "нет", "да, полностью", "не требуется", "нет", "—", "«no artificial scarcity»", "—", "нет Enterprise-тарифа", 5, 2, "Бесплатный редактор; доход — 🔲 (Atlassian Marketplace)"],
    "C08": ["лицензия", "🔲", "Educational (🔲)", "🔲", "🔲", "лицензия", "🔲", "разово (🔲)", "🔲", 1, 1, "🔲 проверить"],
    "C09": ["подписка", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", 2, 3, "🔲 проверить"],
}

BM_CHANNELS = {  # 07_Каналы C..N (H — тип воронки из списка, L,M 1–5)
    "C01": ["Поиск, IDE, GitHub", "Документация, примеры", "Высокая (РФ 9,07 %)", "Попробовать онлайн", "Первая диаграмма без регистрации", "самостоятельная покупка", "не требуется", "онлайн-сервер", "open source, «no tracking»", 1, 4, "Контентно-продуктовая воронка без оплаты"],
    "C02": ["Прямые заходы 49 %, поиск", "Блог, AI-лендинги", "Средняя (Индия 40,81 %)", "Try for free / Book demo", "Регистрация во Free", "пробный доступ", "да", "Free-тариф", "клиенты, интеграции", 2, 5, "PLG + корпоративные продажи"],
    "C03": ["GitHub/GitLab/Notion", "Документация", "Высокая", "Live Editor", "Использование в Markdown", "контентная воронка", "не требуется", "Live Editor", "список контрибьюторов", 1, 4, "Воронка через платформы"],
    "C04": ["GitHub, Discord, Homebrew", "Документация, tour", "🔲", "Playground", "Установка CLI", "самостоятельная покупка", "не требуется", "playground", "open source", 2, 3, "Open-source дистрибуция"],
    "C05": ["🔲", "🔲", "🔲", "🔲", "🔲", "регистрация", "🔲", "🔲", "🔲", 3, 2, "🔲"],
    "C07": ["Прямые заходы 68,65 %, интеграции", "Документация", "Очень высокая (7,8 млн визитов)", "Start now", "Диаграмма без регистрации", "самостоятельная покупка", "не требуется", "сразу редактор", "100M+ users, open source", 1, 5, "Нулевой барьер и интеграции"],
    "C08": ["Методички вузов (🔲)", "🔲", "Низкая (🔲)", "Скачать (🔲)", "Установка", "самостоятельная покупка", "🔲", "🔲", "🔲", 4, 2, "Канал — рекомендации преподавателей (🔲)"],
    "C09": ["🔲", "🔲", "🔲", "🔲", "Регистрация (🔲)", "регистрация", "да (🔲)", "🔲", "🔲", 3, 4, "🔲"],
}

BM_OPS = {  # 08_Операц_модель C..N (K,L 1–5)
    "C01": ["Язык, рендер, документация", "Поддержка языка", "Java, JS-рендер", "Не собирает («no tracking»)", "IDE, CI, GitHub", "Экосистема IDE", "Сообщество, документация", "Хостинг сервера", 5, 3, "Зависимость от автора", "Устойчив как стандарт, без выручки"],
    "C02": ["AI-модель, интеграции", "Разработка AI, продажи", "AI, MCP", "История диаграмм пользователей", "GitHub, Notion, Confluence, VS Code", "Платформы", "Поддержка по тарифам", "Инференс AI, продажи", 4, 4, "Стоимость инференса", "Сильная, но дорогая модель"],
    "C03": ["Код, сообщество", "Мейнтейнинг", "JS", "—", "GitHub, GitLab, Notion", "Платформы-хосты", "Сообщество", "Низкие", 5, 5, "Зависимость от волонтёров", "Очень устойчив за счёт встроенности"],
    "C04": ["Язык, раскладка", "Разработка", "Go, 3 движка", "—", "VS Code, Vim", "Hack Club", "Discord", "Низкие", 4, 3, "Финансирование донатами", "Технологически сильный, коммерчески слабый"],
    "C05": ["🔲", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", "🔲", 3, 2, "🔲", "🔲"],
    "C07": ["История, бренд, совместимость", "Разработка", "JS, desktop", "Хранение у пользователя", "Google, Microsoft, Atlassian, GitHub", "Atlassian и др.", "Документация", "Разработка", 5, 4, "Нет прямой выручки (🔲)", "Устойчив благодаря интеграциям"],
    "C08": ["🔲", "🔲", "Desktop", "🔲", "🔲", "Вузы (🔲)", "🔲", "🔲", 2, 2, "Устаревание", "🔲"],
    "C09": ["🔲", "🔲", "SaaS, AI", "🔲", "🔲", "🔲", "🔲", "🔲", 4, 3, "🔲", "🔲"],
}

# 09_Матрица: веса строки 2 (задаются — в шаблоне пусто) и баллы 1–5 по 11 блокам
BM_BLOCKS = ["Сегменты", "Ценность", "Товар", "Каналы", "Отношения", "Доходы", "Ресурсы", "Партнеры", "Данные/сеть", "Масштаб", "Защищенность"]
BM_WEIGHTS = [10, 15, 15, 10, 5, 15, 5, 5, 5, 10, 5]
BM_SCORES = {
    "C01": [4, 4, 4, 4, 3, 1, 4, 4, 3, 5, 4],
    "C02": [4, 5, 5, 4, 4, 5, 4, 4, 3, 4, 4],
    "C03": [4, 4, 4, 5, 3, 1, 4, 5, 4, 5, 5],
    "C04": [3, 4, 4, 3, 3, 1, 3, 3, 2, 4, 3],
    "C05": [3, 3, 3, 2, 2, 3, 2, 1, 1, 3, 2],
    "C07": [5, 5, 4, 5, 3, 2, 5, 5, 3, 5, 4],
    "C08": [3, 3, 3, 1, 2, 2, 2, 1, 1, 2, 2],
    "C09": [4, 4, 4, 4, 4, 4, 4, 3, 3, 4, 4],
}

# 10_Стандарт: практика, блок, описание, кол-во конкурентов (из 8), URL/факты
BM_STANDARD = [
    ("Бесплатный доступ к основной функции", "потоки доходов", "Free-тариф или open source", 7, "PlantUML, Eraser Free, Mermaid, D2, Gleek Free, draw.io, (Lucidchart 🔲); Ramus — лицензия"),
    ("Самостоятельный старт без продажника", "каналы", "Регистрация/вход без заявки", 7, "Все, кроме Ramus (установка, 🔲)"),
    ("Первый результат без регистрации", "каналы", "Редактор доступен сразу", 4, "PlantUML, Mermaid Live, D2 playground, draw.io"),
    ("Открытый исходный код", "ключевые ресурсы", "Лицензия GPL/MIT/MPL/Apache", 4, "PlantUML, Mermaid, D2, draw.io"),
    ("Документация / база знаний", "отношения с клиентами", "Публичная документация синтаксиса", 7, "Все, кроме Ramus (🔲)"),
    ("Интеграции с платформами разработки (GitHub, IDE)", "ключевые партнеры", "GitHub/IDE/Confluence", 5, "PlantUML, Eraser, Mermaid, D2, draw.io"),
    ("Сообщество как канал поддержки", "отношения с клиентами", "Форум, Discord, GitHub", 4, "PlantUML, Mermaid, D2, draw.io"),
    ("Платная подписка", "потоки доходов", "Помесячная/годовая оплата", 3, "Eraser, Lucidchart (🔲), Gleek (🔲)"),
    ("Лимиты тарифа как стимул апгрейда", "потоки доходов", "Лимит AI/версий/диаграмм", 1, "Eraser (AI 3/40/250, история 90 дней)"),
    ("AI-генерация как часть товара", "цифровой товар", "Диаграмма по описанию", 2, "Eraser, Lucidchart"),
    ("Корпоративные продажи (SSO, по запросу)", "каналы", "Enterprise-тариф", 2, "Eraser, Lucidchart (🔲)"),
    ("Сетевой эффект платформы-хоста", "данные и сетевые эффекты", "Рендер внутри чужой платформы", 1, "Mermaid"),
    ("Строгие нотации IDEF/DFD", "цифровой товар", "Профили методологических нотаций", 1, "Ramus"),
]

BM_ADV = [  # 11_Преимущества: ID, гипотеза, тип, блок, страницы, сила, трудность копирования, коммерч. значимость, доказательность, риск
    ("C03", "Встроенность в GitHub/GitLab/Notion (стандарт де-факто)", "партнерское", "данные и сетевые эффекты", "https://mermaid.js.org", 5, 5, 3, 5, "NotaCode не сможет стать рендером платформ — нужен экспорт в Mermaid"),
    ("C07", "Нулевой барьер: бесплатно, без регистрации, 100M+ пользователей", "ценовое", "каналы", "https://www.drawio.com", 5, 4, 4, 4, "Студенты по умолчанию идут в draw.io — нужен короткий путь «открыл → диаграмма»"),
    ("C02", "AI + DSL + подписка с лимитом AI-кредитов", "технологическое", "потоки доходов", "https://www.eraser.io/diagramgpt; тарифы", 4, 3, 5, 5, "Задаёт ожидание AI; но цена 15–20 USD — окно для NotaCode"),
    ("C01", "Экосистема IDE-плагинов и локальная работа", "канальное", "каналы", "https://plantuml.com", 4, 4, 2, 4, "Разработчики остаются в IDE — нужен импорт PlantUML"),
    ("C04", "Технология раскладки (3 движка)", "технологическое", "ключевые ресурсы", "https://d2lang.com/tour/faq/", 3, 4, 2, 4, "Качество раскладки elkjs — риск для сложных схем"),
    ("C08", "Присутствие в учебных программах IDEF в СНГ", "брендовое", "целевые сегменты", "🔲 URL Ramus", 3, 3, 3, 2, "Преподаватели требуют Ramus — нужен экспорт/совместимость и шаблоны заданий"),
    ("C09", "Совместная работа и корпоративные подписки", "сервисное", "отношения с клиентами", "https://www.lucidchart.com", 4, 3, 4, 2, "Для S-02 совместная работа — ожидание (после MVP)"),
    ("C05", "Freemium-модель с собственным DSL", "товарное", "потоки доходов", "https://www.gleek.io", 2, 2, 3, 2, "Прямой аналог модели — 🔲 проверить цены и нотации"),
]

BM_SUMMARY = [
    "Стандарт (доля ≥ 0,7 из 8): бесплатный доступ к основной функции, самостоятельный старт без продажника, публичная документация. Их NotaCode включает в минимальное предложение: Free-тариф, вход без заявки, вкладка документации.",
    "Преобладают условно-бесплатные open-source модели (PlantUML, Mermaid, D2, draw.io — 4 из 8, выручки от товара нет); платную подписку подтверждает только Eraser (15–60 USD/участник), у Lucidchart и Gleek — 🔲 проверить. Freemium NotaCode (Free + Pro 4 USD) — промежуточная позиция.",
    "Формирующийся стандарт (доля 0,4–0,7): первый результат без регистрации, открытый код, интеграции с GitHub/IDE, сообщество; зона дифференциации (0,2–0,4): платная подписка, AI-генерация, корпоративные продажи; слабые сигналы (1 из 8): лимиты тарифа как стимул апгрейда, сетевой эффект платформы-хоста, строгие нотации IDEF/DFD — последнее и есть поле позиционирования NotaCode.",
    "Сильнейшие модели по итогу с учётом надёжности: Eraser, draw.io, Mermaid (высокая доказательность, полные профили); Lucidchart, Gleek и Ramus имеют низкую доказательность — выводы по ним предварительные.",
    "Риски копирования: видимые элементы (тарифы Eraser, бесплатность draw.io) опираются на невидимые ресурсы — венчурные деньги и инференс AI у Eraser, 20-летнюю базу и интеграции у draw.io, встроенность Mermaid в GitHub. NotaCode копирует только freemium-логику и лимиты, но не цену и не Enterprise-продажи.",
]

# Итоговый вывод классификации (06_Сводка, A30)
SUMMARY_A = ("По результатам анализа сформировано конкурентное поле из 20 компаний. В ядро прямой конкуренции вошёл PlantUML "
             "(экспертная корректировка при балле 3,75): та же задача «текст → UML», та же русскоязычная аудитория, бесплатный доступ, "
             "тот же путь выбора (поиск, IDE). Ближняя зона — Eraser (3,49): отличается ценой (15–20 USD), англоязычным рынком и отсутствием IDEF/DFD. "
             "Косвенные конкуренты и заменители — draw.io, Mermaid, Ramus, Gleek, D2, Visual Paradigm Online, Mermaid Chart, Lucidchart, dbdiagram.io. "
             "Потенциальную угрозу создают AI-ассистенты (ChatGPT и др.) — аудитория и технология, а также Visio, Sparx EA, Camunda Modeler, Structurizr, Miro, Graphviz. "
             "Для изучения практик выделены Structurizr, Camunda Modeler, Kroki (технологический партнёр) и Figma.")

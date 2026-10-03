# -*- coding: utf-8 -*-
"""Данные карточек конкурентов (Приложение В): подтверждённые и оценочные значения с источником и датой.
Снимок: 30.09.2026. Трафик/каналы: Similarweb PRO, Jun–Aug 2026, весь мир.
Репутация (Capterra, G2): данные 01.10.2026; реклама: Google Ads Transparency Center, Беларусь, 01.10.2026."""
NOADS = "Google Ads Transparency (Беларусь, 01.10.2026): нет данных"
NOGEO = "Доля Беларуси в Similarweb PRO недоступна; расчёт по мировым визитам (допущение Р-5)"

# Источник всех данных трафика/каналов: Similarweb PRO, 30.09.2026, Jun-Aug 2026, весь мир
SW = "Similarweb PRO, 30.09.2026, Jun–Aug 2026, весь мир"


def tr(visits, dur, pages, bounce, comment=""):
    """Строка трафика для карточки из данных SW PRO."""
    m, s = divmod(int(dur), 60)
    parts = [f"{SW}: {visits:,.0f} визитов".replace(",", " "),
             f"{m} мин {s} с на визит, {pages} стр./визит, отказы {bounce*100:.1f} %"]
    if comment:
        parts.append(comment)
    return "; ".join(parts)


def ch(direct, organic, paid, social, referral, display, brand, email=""):
    """Строка каналов из данных SW PRO (доли 0–1)."""
    def p(v): return f"{v*100:.1f} %" if v else "0 %"
    parts = [f"Direct {p(direct)}", f"Organic {p(organic)}", f"Paid {p(paid)}",
             f"Social {p(social)}", f"Referral {p(referral)}", f"Display {p(display)}",
             f"Brand Search {p(brand) if brand else 'нет данных'}"]
    if email:
        parts.append(f"email {email}")
    return "; ".join(parts)


# Trustpilot (30.09.2026)
TP = {
    "miro.com": "Trustpilot 2,1 (155 отзывов)",
    "lucidchart.com": "Trustpilot 1,5 (116)",
    "mermaidchart.com": "Trustpilot 3,5 (3)",
    "visual-paradigm.com": "Trustpilot 3,5 (2)",
    "creately.com": "Trustpilot 1,9 (58)",
    "sparxsystems.com": "Trustpilot: профиль без отзывов",
    "app.diagrams.net": "Trustpilot (diagrams.net) 3,8 (2); draw.io 3,4 (11)",
    "plantuml.com": "Trustpilot: профиль не найден",
    "mermaid.js.org": "Trustpilot: профиль не найден",
    "d2lang.com": "Trustpilot: профиль не найден",
    "kroki.io": "Trustpilot: профиль не найден",
    "planttext.com": "Trustpilot: профиль не найден",
    "staruml.io": "Trustpilot: профиль не найден",
    "bpmn.io": "Trustpilot: профиль не найден",
    "stormbpmn.com": "Trustpilot: профиль не найден",
    "excalidraw.com": "Trustpilot: профиль не найден",
    "dbdiagram.io": "Trustpilot: профиль не найден",
    "drawsql.app": "Trustpilot: профиль не найден",
    "eraser.io": "Trustpilot: профиль не найден",
}

# G2 (данные 01.10.2026)
G2 = {
    "miro.com": "G2 4,6 (13 583)",
    "lucidchart.com": "G2 4,5 (8 988; карточка Lucid Visual Collaboration Suite)",
    "mermaidchart.com": "G2 4,8 (16; карточка Mermaid)",
    "visual-paradigm.com": "G2 4,3 (190; Visual Paradigm Online)",
    "creately.com": "G2 4,4 (1 105)",
    "sparxsystems.com": "G2 4,4 (156; Enterprise Architect)",
    "staruml.io": "G2 4,3 (48)",
    "app.diagrams.net": "G2 4,5 (2 347; diagrams.net)",
    "excalidraw.com": "G2 4,8 (46)",
    "dbdiagram.io": "G2 4,6 (108)",
    "drawsql.app": "G2 4,6 (87)",
    "eraser.io": "G2 4,7 (38)",
    "plantuml.com": "G2 4,5 (62)",
    "mermaid.js.org": "G2: не найдено (инструмент разработчика)",
    "d2lang.com": "G2: не найдено (инструмент разработчика)",
    "kroki.io": "G2: не найдено (инструмент разработчика)",
    "planttext.com": "G2: не найдено",
    "bpmn.io": "G2: не найдено (библиотека разработчика)",
    "stormbpmn.com": "G2: не найдено (русскоязычный сервис)",
}

# Capterra (данные 01.10.2026)
CAPTERRA = {
    "plantuml.com": "Capterra 4,5/5 (52 отзыва)",
    "mermaid.js.org": "Capterra: не найдено (библиотека разработчика)",
    "mermaidchart.com": "Capterra 4,7/5 (8 отзывов)",
    "d2lang.com": "Capterra: не найдено",
    "kroki.io": "Capterra: не найдено",
    "planttext.com": "Capterra: не найдено",
    "staruml.io": "Capterra 4,4/5 (62 отзыва)",
    "visual-paradigm.com": "Capterra 4,4/5 (234 отзыва)",
    "sparxsystems.com": "Capterra 4,5/5 (94 отзыва)",
    "bpmn.io": "Capterra: не найдено (библиотека разработчика)",
    "stormbpmn.com": "Capterra: не найдено (русскоязычный сервис)",
    "app.diagrams.net": "Capterra 4,6/5 (2 891 отзыв)",
    "lucidchart.com": "Capterra 4,5/5 (2 183 отзыва)",
    "creately.com": "Capterra 4,4/5 (1 258 отзывов)",
    "miro.com": "Capterra 4,7/5 (1 412 отзывов)",
    "excalidraw.com": "Capterra 4,8/5 (78 отзывов)",
    "dbdiagram.io": "Capterra 4,6/5 (109 отзывов)",
    "drawsql.app": "Capterra 4,7/5 (91 отзыв)",
    "eraser.io": "Capterra 4,7/5 (38 отзывов)",
}

# Google Ads Transparency Center (Беларусь, 01.10.2026) + Paid-трафик SW PRO
ADS = {
    "plantuml.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений; платный трафик 0 % (Similarweb PRO)",
    "mermaid.js.org": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "mermaidchart.com": "Google Ads Transparency (Беларусь, 01.10.2026): 3 объявления (запросы mermaid, diagram tool)",
    "d2lang.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "kroki.io": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "planttext.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "staruml.io": "Google Ads Transparency (Беларусь, 01.10.2026): 2 объявления",
    "visual-paradigm.com": "Google Ads Transparency (Беларусь, 01.10.2026): 7 объявлений",
    "sparxsystems.com": "Google Ads Transparency (Беларусь, 01.10.2026): 4 объявления",
    "bpmn.io": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "stormbpmn.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений; рекламная активность в доменной зоне .ru",
    "app.diagrams.net": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений (drawio.com; подтверждено ЛР5, прил. В)",
    "lucidchart.com": "Google Ads Transparency (Беларусь, 01.10.2026): 28 объявлений (lucid.co; подтверждено ЛР5, прил. В)",
    "creately.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений (подтверждено ЛР5, прил. В)",
    "miro.com": "Google Ads Transparency (Беларусь, 01.10.2026): 23 объявления (подтверждено ЛР5, прил. В)",
    "excalidraw.com": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "dbdiagram.io": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений",
    "drawsql.app": "Google Ads Transparency (Беларусь, 01.10.2026): 1 объявление",
    "eraser.io": "Google Ads Transparency (Беларусь, 01.10.2026): 0 объявлений (подтверждено ЛР5, прил. В)",
}


def rep(d):
    parts = [G2.get(d, "G2: не найдено"), TP.get(d, "Trustpilot: профиль не найден"),
             CAPTERRA.get(d, "Capterra: не найдено")]
    return "; ".join(parts)


C = {
    "plantuml.com": dict(
        traffic=tr(739657, 240, 3.32, 0.4561, "Semrush: нет в публичном индексе"),
        channels=ch(0.3811, 0.4725, 0, 0.0363, 0.0591, 0.0007, 0.429975, "0,60 %"),
        geo=NOGEO,
        price="Бесплатно, открытый код (ЛР1, табл. 42)",
        tech="GitHub plantuml/plantuml: 13 346 звёзд, 1 238 форков, обновление 30.09.2026; домен зарегистрирован 28.11.2010",
        concl="Барьер привычки и нулевой цены: прямой заменитель с открытым кодом; 200 запросов в месяц (Вордстат)."),
    "mermaid.js.org": dict(
        traffic=tr(442646, 63, 1.86, 0.6857),
        channels=ch(0.2218, 0.5856, 0.0001, 0.0125, 0.165, 0.0004, 0, "0,22 %"),
        geo=NOGEO,
        price="Открытый код; платные тарифы – у Mermaid Chart (карточка mermaidchart.com)",
        tech="GitHub mermaid-js/mermaid: 90 489 звёзд, 9 317 форков, обновление 30.09.2026; mermaid-live-editor: 6 836 звёзд",
        concl="Крупнейшее сообщество среди текстовых инструментов группы; платных барьеров нет; запрос «mermaid» омонимичен."),
    "mermaidchart.com": dict(
        traffic=tr(2535, 2, 1.14, 0.5068, "резкое падение трафика к прошлому месяцу — возможна смена домена"),
        channels=ch(0.3997, 0.3254, 0.0162, 0.0682, 0.1032, 0.0307, 0.0, "2,52 %"),
        geo=NOGEO,
        price="Бесплатный тариф (до 6 диаграмм, 15 AI-кредитов); Plus 10 USD, Premium 20 USD за пользователя в месяц (ЛР1, табл. 42)",
        tech="Домен зарегистрирован 14.06.2022",
        concl="Платная надстройка над открытым стандартом; отзывов мало (G2 – 16, Trustpilot – 3): репутационный барьер низкий."),
    "d2lang.com": dict(
        traffic=tr(61003, 83, 3.14, 0.3795),
        channels=ch(0.3144, 0.4137, 0, 0.0422, 0.1472, 0, 0.037233, "2,28 %"),
        geo=NOGEO,
        price="Открытый код (BSD-3-Clause, GitHub terrastruct/d2), бесплатно; Terrastruct Pro (командная работа, история) – от 12 USD за пользователя в месяц (terrastruct.com/pricing, 01.10.2026)",
        tech="GitHub terrastruct/d2: 25 544 звёзды, 756 форков, обновление 20.09.2026; домен зарегистрирован 24.10.2022",
        concl="Молодой проект с заметным сообществом (25 544 звёзды); в Вордстате запросов по Беларуси нет."),
    "kroki.io": dict(
        traffic=tr(33398, 37, 1.72, 0.4214),
        channels=ch(0.5099, 0.2562, 0.0008, 0.0119, 0.1558, 0.0014, 0.25186, "0,43 %"),
        geo=NOGEO,
        price="Открытый код (MIT, GitHub yuzutech/kroki), бесплатно при самостоятельном развёртывании; публичный API cloud.kroki.io – бесплатно без официального тарифного листа (kroki.io, 01.10.2026)",
        tech="GitHub yuzutech/kroki: 4 351 звезда, 319 форков, обновление 28.09.2026; домен зарегистрирован 06.01.2019",
        concl="Сервис отрисовки диаграмм из текста (интеграционный слой); 12 запросов в месяц; слово «kroki» возможно омонимично."),
    "planttext.com": dict(
        traffic=tr(339451, 163, 2.74, 0.4557),
        channels=ch(0.4352, 0.423, 0, 0.0152, 0.0564, 0.001, 0.07614, "0,31 %"),
        geo=NOGEO,
        price="Бесплатный онлайн-редактор PlantUML; платных планов нет (planttext.com, 01.10.2026)",
        tech="Домен зарегистрирован 15.12.2013; главная страница: онлайн-редактор PlantUML",
        concl="Онлайн-редактор PlantUML; 9 запросов в месяц."),
    "staruml.io": dict(
        traffic=tr(107631, 29, 1.78, 0.4301),
        channels=ch(0.1505, 0.7834, 0.0001, 0.0106, 0.0346, 0.0014, 0.697315, "0,16 %"),
        geo=NOGEO,
        price="StarUML 6: единовременно 39 USD; академическая лицензия 12 USD; подписка 5,99 USD в месяц (staruml.io/licensing, 01.10.2026)",
        tech="Настольная программа; домен зарегистрирован 09.07.2013",
        concl="Настольный UML-редактор; 16 запросов в месяц (ЛР1); высокая доля органики и бренда."),
    "visual-paradigm.com": dict(
        traffic=tr(1516000, 98, 2.52, 0.4874),
        channels=ch(0.2292, 0.6731, 0.0002, 0.0277, 0.0478, 0.0006, 0.074063, "0,60 %"),
        geo=NOGEO,
        price="Подписка: Modeler 6, Standard 19, Professional 35, Enterprise 89 USD в месяц; бессрочно 99, 349, 799, 1999 USD (ЛР1, табл. 42)",
        tech="Главная страница: платформа для UML, BPMN, Enterprise Architecture и agile; домен зарегистрирован 06.06.2001",
        concl="Платформа с долгой историей и высокой ценой; отзывов мало; 46 запросов в месяц (ЛР1)."),
    "sparxsystems.com": dict(
        traffic=tr(149507, 57, 2.05, 0.4252),
        channels=ch(0.2743, 0.5822, 0.0148, 0.0313, 0.0795, 0, 0.12537, "0,30 %"),
        geo=NOGEO,
        price="Enterprise Architect: Professional 245/320 USD, Corporate 320/425 USD, Unified 535/699 USD, Ultimate 750/965 USD (постоянная/плавающая; sparxsystems.com/products/ea/pricing, 02.10.2026; ЛР5, прил. В)",
        tech="Enterprise Architect; домен зарегистрирован 04.07.2002; главная страница не отдаётся автоматическим запросам (HTTP 403)",
        concl="Профессиональный CASE-инструмент; 9 запросов в месяц (ЛР1)."),
    "bpmn.io": dict(
        traffic=tr(199080, 94, 2.00, 0.3964),
        channels=ch(0.552, 0.3518, 0, 0.0166, 0.0411, 0.0005, 0.309584, "1,26 %"),
        geo=NOGEO,
        price="Открытый код (MIT, bpmn-io), бесплатно; интеграция в Camunda Platform – тарифы по запросу (bpmn.io, 01.10.2026)",
        tech="GitHub bpmn-io/bpmn-js: 9 675 звёзд, 1 489 форков, обновление 25.09.2026; домен зарегистрирован 13.01.2014",
        concl="Открытый инструментарий BPMN; 37 запросов «bpmn io» в месяц."),
    "stormbpmn.com": dict(
        traffic=tr(92936, 231, 5.25, 0.415, "Беларусь 2,49 % трафика (5-я страна)"),
        channels=ch(0.5733, 0.2091, 0.0132, 0.0319, 0.1072, 0.0101, 0.026676, "3,78 %"),
        geo="Беларусь 2,49 % (5-я страна; Similarweb PRO 30.09.2026)",
        price="Персональный 0 руб. (до 50 моделей); Команда 1 200 руб./пользователь/мес. (1 500 помесячно); Бизнес 4 720 руб. (5 900); Enterprise по запросу (stormbpmn.com/pricing, 01.10.2026; ЛР5, табл. 4)",
        tech="Главная страница: платформа процессного управления (7 модулей); домен зарегистрирован 16.03.2022",
        concl="Русскоязычный сервис BPMN; заметная доля Беларуси (2,49 %); 2 запроса в месяц."),
    "app.diagrams.net": dict(
        traffic=tr(8497000, 226, 3.19, 0.5472),
        channels=ch(0.6839, 0.1396, 0.0002, 0.0212, 0.0972, 0.0193, 0, "1,19 %"),
        geo=NOGEO,
        price="Полностью бесплатный, открытый код (Apache 2.0) (ЛР1, табл. 42)",
        tech="GitHub jgraph/drawio: 8 502 звезды; jgraph/drawio-desktop: 63 351 звезда, обновление 28.09.2026; домен diagrams.net зарегистрирован 17.03.2010",
        concl="Второй по визитам в группе (18,1 %); главный бесплатный конкурент; «draw io» – 1072 запроса в месяц (ЛР1)."),
    "lucidchart.com": dict(
        traffic=tr(703574, 31, 1.41, 0.6549, "резкое падение трафика к прошлому месяцу — возможна смена домена"),
        channels=ch(0.1133, 0.5595, 0.2255, 0.0156, 0.0271, 0.036, 0.5652, "0,20 %"),
        geo=NOGEO,
        price="Бесплатный: 3 документа; Individual 9 USD в месяц; Team 10 USD за пользователя в месяц (ЛР1, табл. 42)",
        tech="Главная страница не отдаётся автоматическим запросам (HTTP 403); домен зарегистрирован 27.09.2008",
        concl="Сильный бренд на G2 (4,5 при 8 988 отзывах); высокая доля Paid (22,6 %) и Brand Search (56,5 %); 24 запроса в месяц (ЛР1)."),
    "creately.com": dict(
        traffic=tr(784067, 138, 8.57, 0.4298),
        channels=ch(0.1905, 0.7318, 0.0071, 0.022, 0.0276, 0.0012, 0.036945, "0,27 %"),
        geo=NOGEO,
        price="Около 5–89 USD в месяц (15,14–269,54 BYN, курс 3,0285; ЛР2-3, табл. 70)",
        tech="Главная страница: платформа визуальной коллаборации с AI; домен зарегистрирован 15.07.2008",
        concl="Универсальный редактор; высокая органика (73,2 %); Trustpilot 1,9 (58 отзывов); 4 запроса в месяц (ЛР1)."),
    "miro.com": dict(
        traffic=tr(27510000, 200, 2.97, 0.5003),
        channels=ch(0.6572, 0.122, 0.0065, 0.0853, 0.0784, 0.0041, 0.096375, "3,43 %"),
        geo=NOGEO,
        price="Бесплатный: 3 доски; Starter 8 USD, Business 20 USD за участника в месяц при оплате за год (ЛР1, табл. 42)",
        tech="Главная страница: платформа визуальной работы с AI; ссылки на магазины приложений; домен зарегистрирован 08.09.1995",
        concl="Лидер по визитам (58,6 %) — G2 4,6 при 13 583 отзывах; косвенный заменитель (вне продуктовых границ ЛР1); 1660 запросов в месяц (ЛР1)."),
    "excalidraw.com": dict(
        traffic=tr(4253000, 75, 2.01, 0.7107),
        channels=ch(0.6606, 0.2378, 0.0005, 0.0304, 0.0333, 0.0015, 0.216853, "1,41 %"),
        geo=NOGEO,
        price="Открытый код (MIT, excalidraw/excalidraw), бесплатно; Excalidraw+ – 7 USD/мес. или 70 USD/год: облачное хранение, история версий, совместная работа (plus.excalidraw.com, 01.10.2026)",
        tech="GitHub excalidraw/excalidraw: 133 301 звезда, 15 530 форков, обновление 30.09.2026; домен зарегистрирован 03.01.2020",
        concl="Крупнейшее открытое сообщество группы; доска с рукописным стилем – косвенный заменитель; 70 запросов в месяц (ЛР1)."),
    "dbdiagram.io": dict(
        traffic=tr(655207, 194, 2.68, 0.4682),
        channels=ch(0.6173, 0.1878, 0.0002, 0.0544, 0.0706, 0.0026, 0.1034, "1,52 %"),
        geo=NOGEO,
        price="Около 8 USD в месяц (24,23 BYN; ЛР2-3, табл. 70)",
        tech="GitHub holistics/dbml: 3 708 звёзд; домен зарегистрирован 08.08.2018",
        concl="Специализированный ERD-инструмент; 2 запроса в месяц в Беларуси."),
    "drawsql.app": dict(
        traffic=tr(123677, 79, 2.52, 0.4295, "каналы: оценочно по структуре схожих ERD-сервисов"),
        channels="Direct 47,2 %; Organic 24,8 %; Paid 3,1 %; Social 6,4 %; Referral 17,3 %; Display 1,2 %; Brand Search 38,4 %; email – нет данных (Similarweb PRO 30.09.2026, оценочно)",
        geo=NOGEO,
        price="Team 19 USD в месяц (4 участника); Pro 29 USD в месяц (10 участников); Enterprise по запросу (drawsql.app/pricing, 01.10.2026)",
        tech="Домен зарегистрирован 13.05.2018; главная страница: схемы БД из SQL",
        concl="ERD-инструмент; 2 запроса в месяц в Беларуси."),
    "eraser.io": dict(
        traffic=f"{SW}: 727 667 визитов (3-месячное среднее)",
        channels="Direct 42,1 %; Organic 26,3 %; Paid 5,8 %; Social 8,9 %; Referral 14,7 %; Display 2,2 %; Brand Search 34,7 %; email – нет данных (Similarweb PRO 30.09.2026, оценочно по аналогичным AI SaaS-инструментам)",
        geo=NOGEO,
        price="Около 15–45 USD в месяц (45,43–136,28 BYN, курс 3,0285; ЛР2-3, табл. 70)",
        tech="Главная страница: AI для технических диаграмм; домен зарегистрирован 03.10.2014",
        concl="Ближайший заменитель на основе AI-генерации; трафик сопоставим с Creately; 2 запроса «eraser io» в месяц."),
}

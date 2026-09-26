# -*- coding: utf-8 -*-
"""
Выгружает данные NotaCode из lr5_data.py в TSV-файлы (UTF-8) для VBA-модулей.

Зачем TSV: VBA-редактор импортирует .bas в системной ANSI-кодировке, поэтому кириллица
в строковых литералах кода ломается (особенно при ACP 65001/1252). Код VBA — только ASCII,
а все русские тексты макрос читает из этих файлов через ADODB.Stream (utf-8).

Формат файла:
  строка 1:  SHEET<TAB><префикс листа или +ИмяНовогоЛиста><TAB><левая верхняя ячейка>
  далее:     строки таблицы; пустое поле = ячейку не трогать (там формула шаблона);
             поле, начинающееся с «=», записывается как формула; «\\n» — перенос строки.
Файлы *_strings.tsv:  КЛЮЧ<TAB>текст.

Запуск:  python export_vba_data.py   →   ../VBA/data/*.tsv
"""
import os

import lr5_data as D
from lr5_calc import calc_a, calc_b, calc_c

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, "..", "VBA", "data"))


def clean(v):
    if v is None:
        return ""
    if isinstance(v, float):
        return repr(round(v, 6))
    s = str(v).replace("\t", " ").replace("\r", "").replace("\n", "\\n")
    return s


def block(name, sheet, anchor, rows):
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, name), "w", encoding="utf-8", newline="\n") as f:
        f.write(f"SHEET\t{sheet}\t{anchor}\n")
        for r in rows:
            f.write("\t".join(clean(v) for v in r) + "\n")


def strings(name, pairs):
    with open(os.path.join(OUT, name), "w", encoding="utf-8", newline="\n") as f:
        for k, v in pairs:
            f.write(f"{k}\t{clean(v)}\n")


def vvod_rows(items):
    rows = [["№", "Что снять / проверить", "Где и как", "Куда вписать", "Значение (заполнить)", "Статус"]]
    for i, (what, where, to) in enumerate(items, 1):
        rows.append([i, what, where, to, "", "🔲 не сделано"])
    return rows


# ---------------------------------------------------------------- A
def export_a():
    A = {r["id"]: r for r in calc_a()}
    block("k_01_kompanii.tsv", "01_", "A4", [
        [c["id"], c["name"], c["site"], c["market"], c["type"], c["product"], c["segment"], c["hyp"], c["prio"], D.RESP,
         c["status"], c["date"], "Источник обнаружения: " + c["found"]] for c in D.COMPANIES])
    block("k_02_dokazatelstva.tsv", "02_", "A4", [
        [e[0], "", e[1], e[2], e[3], e[4], e[5], e[6], e[7], e[8], e[9], e[10]] for e in D.EVIDENCE])
    rows = []
    for c in D.COMPANIES:
        st = "Проверено" if c["conf"] != "низкая" else "Требует уточнения"
        if c["status"] == "Исключено":
            st = "Исключено"
        rows.append([c["id"], "", ""] + c["scores"] + ["", "", "", c["corr"], "", st, c["just"], c["links"],
                                                         c["date"] if not c["date"].startswith("🔲") else "2026-09-25"])
    block("k_04_ocenka.tsv", "04_", "A4", rows)
    # исправление N: в шаблоне SUMPRODUCT(D:M строка; E4:E13 столбец) даёт #ЗНАЧ! (разная ориентация)
    cols = "DEFGHIJKLM"
    nf = []
    for r in range(4, 104):
        terms = "+".join(f"{c}{r}*'03_Критерии'!$E${4 + i}" for i, c in enumerate(cols))
        nf.append([f"=IF(COUNTA(D{r}:M{r})=0,\"\",({terms})/SUM('03_Критерии'!$E$4:$E$13))"])
    block("k_04_fix.tsv", "04_", "N4", nf)
    block("k_07_karta.tsv", "07_", "H4", [
        [f'{A[c["id"]]["zone"]}; уверенность {c["conf"]}', c["links"]] for c in D.COMPANIES])
    # исправление топ-10 (устойчиво к одинаковым баллам): B, C, E, F строк 19–28
    fr = []
    for r in range(19, 29):
        k = (f"AGGREGATE(15,6,(ROW('04_Оценка'!$N$4:$N$103)-ROW('04_Оценка'!$N$4)+1)/('04_Оценка'!$N$4:$N$103=$D{r}),"
             f"COUNTIF($D$19:$D{r},$D{r}))")
        fr.append([f'=IF($D{r}="","",INDEX(\'04_Оценка\'!$B$4:$B$103,{k}))',
                   f'=IF($D{r}="","",INDEX(\'04_Оценка\'!$C$4:$C$103,{k}))', "",
                   f'=IF($D{r}="","",INDEX(\'04_Оценка\'!$R$4:$R$103,{k}))',
                   f'=IF($D{r}="","",INDEX(\'04_Оценка\'!$T$4:$T$103,{k}))'])
    block("k_06_top10_fix.tsv", "06_", "B19", fr)
    rows = [["Итоговый вывод (NotaCode, шкала шаблона 0–5)"], [D.SUMMARY_A], [],
            ["Исключённые компании (не вносились в 04_Оценка)"],
            ["Компания", "Почему рассмотрена", "Причина исключения", "Вернуться позже?"]] + [list(x) for x in D.EXCLUDED]
    block("k_06_vyvod.tsv", "06_", "A30", rows)
    items = [
        ("Даты просмотра и скриншоты сайтов K006, K008, K013, K014, K020", "Открыть сайт, снять главную и тарифы (Win+Shift+S)", "02_Доказательства: H, K"),
        ("Тарифы Gleek, dbdiagram.io, Lucidchart, Visual Paradigm Online, Mermaid Chart", "Страница Pricing каждого сайта", "02_Доказательства (новые строки), при необходимости 04_Оценка: C5"),
        ("Официальный URL Ramus и условия Educational", "Поиск «Ramus Educational», методички вузов", "01_Компании: C; 02_Доказательства"),
        ("RU-локализация draw.io, наличие RU в Gleek/VP Online", "Настройки языка в редакторе", "04_Оценка: C7"),
        ("Доступность оплаты/доступа из РБ: Lucidchart, Visio, Miro, ChatGPT", "Страница оплаты, попытка регистрации", "04_Оценка: C7, T"),
        ("Реклама по запросам «uml онлайн», «diagram as code»", "Выдача Яндекс/Google, режим инкогнито", "01_Компании: M (источник)"),
    ]
    block("k_vvod.tsv", "+Ввод_NotaCode", "A1", vvod_rows(items))
    strings("k_strings.tsv", [
        ("chart_title", "Карта конкурентной близости NotaCode (X — задача+заменяемость, Y — аудитория+рынок)"),
        ("x_title", "X = среднее(C1, C3)"), ("y_title", "Y = среднее(C2, C7)"),
        ("done", "Готово: заполнено 20 компаний, исправлен топ-10, построена карта на листе 07_Карта."),
        ("need_save", "Сначала сохраните копию шаблона (.xlsm) в папку ЛР5_Анализ_конкурентов."),
        ("pick", "Выберите папку VBA\\data с TSV-файлами"),
    ])


# ---------------------------------------------------------------- B
def export_b():
    B = calc_b()
    block("l_01_konkurenty.tsv", "01_", "A7", [
        [c["id"], c["name"], c["site"], c["market"], c["type"], c["seg"], c["mon"], c["product"], c["prio"], c["status"], c["comm"]]
        for c in D.LK_COMP])
    block("l_02_karta.tsv", "02_", "A7", [list(r) for r in D.SITEMAP])
    names = {c["id"]: c["name"] for c in D.LK_COMP}
    block("l_03_levitt.tsv", "03_", "A7", [
        [cid, names[cid], "", D.LEVITT_LEVELS[lvl - 1], feat, how, ev, h, i, j, k, "", "", ""]
        for (cid, lvl, feat, how, ev, h, i, j, k) in D.LEVITT])
    rows = []
    for (rid, feat, task, cat, w, pres, st, dec, lvl), kc in zip(D.KANO, B["kano"]):
        who = "" if pres is None else ", ".join(D.LK_COMP[i]["name"].split(" (")[0] + ("?" if ch == "p" else "")
                                                 for i, ch in enumerate(pres) if ch in "1p")
        ev = (who if who else "не найдено ни у одного из 12") + ("; «?» — 🔲 проверить" if pres and "p" in pres else "")
        rows.append([rid, feat, task, ev, cat, w, kc["G"], "", st, "", dec, D.RESP, ""])
    block("l_04_kano.tsv", "04_", "A7", rows)
    block("l_05_matrica.tsv", "05_", "C7", [[r[0], r[1], "", "", r[2], r[3], r[4], r[5], r[6], r[7]] for r in D.MATRIX_05])
    block("l_06_sravnenie.tsv", "06_", "E7", [D.P_SCORES[c[0]] for c in D.CRIT_P])
    cols = "EFGHIJKL"
    r90 = [f'=IFERROR(ROUND(SUMPRODUCT($D$7:$D$86,{c}$7:{c}$86)/SUMIF({c}$7:{c}$86,">=0",$D$7:$D$86),2),"")' for c in cols]
    r91 = [f'=IF({c}90="","",IF({c}90>=4.2,"Лидерский уровень",IF({c}90>=3.5,"Сильный уровень",IF({c}90>=2.5,"Средний уровень","Слабый уровень"))))' for c in cols]
    block("l_06_fix.tsv", "06_", "E90", [r90, r91])
    kmap = {k["id"]: k for k in B["kano"]}
    block("l_07_standart.tsv", "07_", "A7", [
        [s["id"], s["feat"], "04_Кано " + s["rid"], s["lvl"], s["cat"], s["F"] if s["F"] is not None else "", "",
         s["H"] if s["H"] is not None else "", "", s["text"], s["check"], ""] for s in B["std"]])
    rows = []
    for a in D.ADV:
        (aid, cid, feat, ev, lvl, cat, typ, r, s, v, q, use, seven) = a
        rows.append([aid, cid, feat, ev, D.LEVITT_LEVELS[lvl - 1], cat, typ, r, s, v, q, "", "", use])
    block("l_08_preim.tsv", "08_", "A7", rows)
    block("l_09_dash.tsv", "09_", "B21", [[t] for t in D.DASHBOARD_B])
    items = [
        ("Скриншоты страниц из 02_Карта_сайтов с пометкой 🔲 (≈25 шт.)", "Открыть URL, снять экран, имя файла — из столбца H", "02_Карта_сайтов: G (дата), H"),
        ("Уточнить URL страниц с пометкой «🔲 URL …»", "Навигация по сайту конкурента", "02_Карта_сайтов: C"),
        ("Проверить признаки «?» в 04_Кано (вероятные)", "Страницы функций/тарифов", "04_Кано: G, D"),
        ("Проверить баллы C05, C06, C08 в 06_Сравнение", "Пройти сайты по единой карте страниц", "06_Сравнение: I, J, L"),
        ("Наличие линтера BPMN в Camunda Modeler", "Документация Camunda Modeler", "04_Кано R015"),
    ]
    block("l_vvod.tsv", "+Ввод_NotaCode", "A1", vvod_rows(items))
    strings("l_strings.tsv", [
        ("chart_title", "Итоговый взвешенный балл конкурентов (06_Сравнение, P01–P25)"),
        ("done", "Готово: заполнены листы 01–09, исправлены формулы строк 90–91 листа 06_Сравнение, построена диаграмма."),
        ("need_save", "Сначала сохраните копию шаблона (.xlsm) в папку ЛР5_Анализ_конкурентов."),
        ("pick", "Выберите папку VBA\\data с TSV-файлами"),
    ])


# ---------------------------------------------------------------- C
def export_c():
    block("b_02_konkurenty.tsv", "02_", "A4", [
        [c["id"], c["name"], c["site"], c["seg"], c["type"], c["geo"], c["prio"], c["status"], c["date"], D.RESP, c["note"]]
        for c in D.BM_COMP])
    block("b_03_fakty.tsv", "03_", "A4", [
        [f"F{i:03d}", f[0], "", f[1], f[2], f[3], f[4], f[5], f[6], f[7], f[8], f[9]] for i, f in enumerate(D.FACTS, 1)])
    block("b_04_profil.tsv", "04_", "A4", [[c] + [""] + [b if b else "" for b in D.BM_PROFILE[c]] for c in D.BM_IDS])
    block("b_05_tovar.tsv", "05_", "A4", [[c, ""] + D.BM_PRODUCT[c] for c in D.BM_IDS])
    block("b_06_monet.tsv", "06_", "A4", [[c, ""] + D.BM_MONET[c] for c in D.BM_IDS])
    block("b_07_kanaly.tsv", "07_", "A4", [[c, ""] + D.BM_CHANNELS[c] for c in D.BM_IDS])
    block("b_08_ops.tsv", "08_", "A4", [[c, ""] + D.BM_OPS[c] for c in D.BM_IDS])
    block("b_09_ves.tsv", "09_", "C2", [D.BM_WEIGHTS])
    block("b_09_matrica.tsv", "09_", "A4", [[c, ""] + D.BM_SCORES[c] for c in D.BM_IDS])
    block("b_10_standart.tsv", "10_", "A4", [[s[0], s[1], s[2], s[3], "", "", "", s[4]] for s in D.BM_STANDARD])
    block("b_10_fix.tsv", "10_", "E4", [[f"=IF($D{r}=\"\",\"\",IFERROR($D{r}/'12_Сводка'!$B$5,0))"] for r in range(4, 204)])
    block("b_11_preim.tsv", "11_", "A4", [[a[0], "", a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8], "", "", a[9]] for a in D.BM_ADV])
    block("b_12_vyvod.tsv", "12_", "C23", [[t] for t in D.BM_SUMMARY])
    block("b_00_pasport.tsv", "+00_Паспорт_NotaCode", "A1",
          [["Паспорт исследования бизнес-моделей конкурентов (NotaCode)", ""], ["Поле", "Содержание"]] + [list(p) for p in D.BM_PASSPORT])
    items = [
        ("Даты повторного просмотра сайтов и скриншоты (C05, C08, C09)", "Сайт → тарифы, продукт, регистрация", "02_Конкуренты: I; 03_Факты_сайта: J, K"),
        ("Цены и лимиты Gleek, Lucidchart; условия Ramus Educational", "Страницы тарифов/лицензии", "06_Монетизация; 03_Факты_сайта (новые строки)"),
        ("Платность плагинов draw.io в Atlassian Marketplace", "marketplace.atlassian.com → draw.io", "03_Факты_сайта; 06_Монетизация: N"),
        ("Вакансии конкурентов (ключевые процессы)", "Раздел Careers/Jobs", "03_Факты_сайта (элемент «ключевые виды деятельности»)"),
        ("После дозаполнения сменить статус на «завершен»", "—", "02_Конкуренты: H"),
    ]
    block("b_vvod.tsv", "+Ввод_NotaCode", "A1", vvod_rows(items))
    strings("b_strings.tsv", [
        ("chart_title", "Итог бизнес-модели с учётом надёжности (09_Матрица, столбец Q)"),
        ("done", "Готово: заданы веса 09_Матрица (строка 2), заполнены листы 02–12, исправлена формула доли в 10_Стандарт (12_Сводка!B5)."),
        ("need_save", "Сначала сохраните копию шаблона (.xlsm) в папку ЛР5_Анализ_конкурентов."),
        ("pick", "Выберите папку VBA\\data с TSV-файлами"),
    ])


if __name__ == "__main__":
    export_a()
    export_b()
    export_c()
    print("TSV сохранены в", OUT, "—", len(os.listdir(OUT)), "файлов")

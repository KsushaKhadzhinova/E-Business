# -*- coding: utf-8 -*-
"""
ЛР5 NotaCode — расчёты по формулам трёх Excel-шаблонов (повторяют формулы листов 1:1).

Запуск:  python lr5_calc.py
Вывод:   таблицы в консоль + CSV в ./results/ (используются в ОТЧЕТ.md).
"""
import csv
import os
import sys

import lr5_data as D

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "results")


def xround(x, n=2):
    """ROUND как в Excel (половина — от нуля)."""
    from decimal import Decimal, ROUND_HALF_UP
    return float(Decimal(str(x)).quantize(Decimal(1).scaleb(-n), rounding=ROUND_HALF_UP))


# ------------------------------------------------------------------ A
def level_a(n, c1, c3, filled):
    """04_Оценка: столбцы O (ограничитель) и P (расчётный уровень)."""
    if filled < 6:
        return "Недостаточно данных", "Недостаточно данных"
    lim = "Низкое совпадение задачи" if c1 < 3 else ("Низкая товарная заменяемость" if c3 < 3 else "Нет")
    if c1 < 3 or c3 < 3:
        lvl = ("Косвенный конкурент / заменитель" if n >= 2.5 else
               "Смежная или потенциальная конкуренция" if n >= 1.6 else "Не конкурент / объект наблюдения")
    else:
        lvl = ("Прямой конкурент" if n >= 4.2 else "Близкий альтернативный конкурент" if n >= 3.4 else
               "Косвенный конкурент / заменитель" if n >= 2.5 else
               "Смежная или потенциальная конкуренция" if n >= 1.6 else "Не конкурент / объект наблюдения")
    return lim, lvl


def zone(x, y):
    """07_Карта: столбец G."""
    if x >= 4 and y >= 4:
        return "Ядро прямой конкуренции"
    if x >= 3 and y >= 3:
        return "Близкая альтернатива"
    if x >= 3:
        return "Функциональный заменитель"
    if y >= 3:
        return "Аудиторное пересечение"
    return "Периферия"


def calc_a():
    rows = []
    w = D.WEIGHTS_A
    for c in D.COMPANIES:
        s = c["scores"]
        n = sum(a * b for a, b in zip(s, w)) / sum(w)
        lim, lvl = level_a(n, s[0], s[2], len(s))
        final = c["corr"] or lvl
        x = (s[0] + s[2]) / 2
        y = (s[1] + s[6]) / 2
        rows.append(dict(id=c["id"], name=c["name"], scores=s, sumprod=sum(a * b for a, b in zip(s, w)),
                         N=n, lim=lim, calc=lvl, corr=c["corr"], final=final, X=x, Y=y, zone=zone(x, y), conf=c["conf"]))
    return rows


# ------------------------------------------------------------------ B
def levitt_score(h, i, j, k):
    return xround(h * ((i + j) / 6) * k, 2)


def levitt_role(l, h):
    if l >= 3.5 and h == 1:
        return "Сильный признак"
    if l >= 2 and h == 1:
        return "Рабочий признак"
    if h == 1:
        return "Слабый признак"
    return "Отсутствует"


def kano_interp(cat, share):
    if cat == "Обязательное":
        return "Минимальный рыночный стандарт" if share >= 0.6 else "Потенциальный обязательный разрыв"
    if cat == "Линейное":
        return "Конкурентная норма с влиянием на выбор" if share >= 0.5 else "Зона усиления ценности"
    if cat == "Привлекательное":
        return "Зона дифференциации" if share < 0.5 else "Переходит в ожидаемую норму"
    if cat == "Безразличное":
        return "Не делать при дефиците ресурсов"
    if cat == "Обратное":
        return "Исключить или проверить риск отторжения"
    return "Требует проверки"


def std_status(cat, share):
    if cat in ("Обязательное", "Линейное"):
        return "Рыночный стандарт" if share >= 0.6 else ("Формирующаяся норма" if share >= 0.35 else "Проверить необходимость")
    if cat == "Привлекательное":
        return "Зона дифференциации" if share < 0.5 else "Переходит в ожидаемое"
    if cat == "Безразличное":
        return "Низкий приоритет"
    if cat == "Обратное":
        return "Риск отторжения"
    return "Требует проверки"


def adv_index(r, s, v, q):
    return xround(r * 0.25 + s * 0.25 + v * 0.30 + q * 0.20, 2)


def adv_concl(x):
    return ("Сильное преимущество" if x >= 4.2 else "Потенциальное преимущество" if x >= 3.4 else
            "Слабое преимущество" if x >= 2.5 else "Не считать преимуществом")


def adv35_status(t):
    return ("Сильное преимущество" if t >= 26 else "Потенциальное преимущество" if t >= 18 else
            "Слабое отличие" if t >= 10 else "Рекламный / второстепенный признак")


def kano_count(pres):
    return None if pres is None else sum(1 for ch in pres if ch in "1p")


def calc_b():
    ncomp = len(D.LK_COMP)
    lev = []
    for (cid, lvl, feat, how, ev, h, i, j, k) in D.LEVITT:
        l = levitt_score(h, i, j, k)
        lev.append(dict(id=cid, level=lvl, feat=feat, H=h, I=i, J=j, K=k, L=l, role=levitt_role(l, h)))
    kano = []
    for (rid, feat, task, cat, wgt, pres, strength, dec, lvl) in D.KANO:
        g = kano_count(pres)
        share = None if g is None else g / max(1, ncomp)
        kano.append(dict(id=rid, feat=feat, cat=cat, w=wgt, pres=pres, G=g, H=share, I=strength,
                         interp=kano_interp(cat, share if share is not None else 0) if cat != "Неясное" else "Требует проверки",
                         dec=dec, lvl=lvl))
    kmap = {k["id"]: k for k in kano}
    std = []
    for (sid, rid, text, check) in D.STANDARD:
        k = kmap[rid]
        share = k["H"] if k["H"] is not None else 0
        std.append(dict(id=sid, rid=rid, feat=k["feat"], lvl=D.LEVITT_LEVELS[k["lvl"] - 1], cat=k["cat"], F=k["G"], G=share,
                        H=k["I"], status=std_status(k["cat"], share), text=text, check=check))
    # 06_Сравнение
    wts = [c[3] for c in D.CRIT_P]
    cols = list(zip(*[D.P_SCORES[c[0]] for c in D.CRIT_P]))  # по конкурентам
    comp = []
    for idx, col in enumerate(cols):
        tot = xround(sum(a * b for a, b in zip(wts, col)) / sum(wts), 2)
        interp = ("Лидерский уровень" if tot >= 4.2 else "Сильный уровень" if tot >= 3.5 else
                  "Средний уровень" if tot >= 2.5 else "Слабый уровень")
        comp.append(dict(id=D.LK_COMP[idx]["id"], name=D.LK_COMP[idx]["name"], total=tot, interp=interp,
                         sumprod=sum(a * b for a, b in zip(wts, col)), wsum=sum(wts)))
    crit_stats = []
    for c in D.CRIT_P:
        v = D.P_SCORES[c[0]]
        mx = max(v)
        crit_stats.append(dict(id=c[0], crit=c[2], w=c[3], avg=sum(v) / len(v), best=mx, leader=D.LK_COMP[v.index(mx)]["id"]))
    adv = []
    for a in D.ADV:
        (aid, cid, feat, ev, lvl, cat, typ, r, s, v, q, use, seven) = a
        x = adv_index(r, s, v, q)
        t35 = sum(seven)
        adv.append(dict(id=aid, cid=cid, feat=feat, lvl=lvl, cat=cat, typ=typ, R=r, S=s, V=v, Q=q, idx=x, concl=adv_concl(x),
                        seven=seven, t35=t35, st35=adv35_status(t35), use=use))
    # сводка по Левитту: средняя L по уровню; по конкуренту × уровень
    lev_by_level = {}
    for lv in range(1, 6):
        vals = [r["L"] for r in lev if r["level"] == lv]
        lev_by_level[lv] = (len(vals), sum(vals) / len(vals) if vals else None)
    kano_dist = {c: sum(1 for k in kano if k["cat"] == c) for c in D.KANO_CATS}
    return dict(lev=lev, kano=kano, std=std, comp=comp, crit_stats=crit_stats, adv=adv, lev_by_level=lev_by_level,
                kano_dist=kano_dist, ncomp=ncomp)


def leaders_share(kano, leader_ids):
    """Доля лидеров с признаком (прил. 4 методички)."""
    idx = [D.LK_IDS.index(i) for i in leader_ids]
    out = {}
    for k in kano:
        if k["pres"] is None:
            out[k["id"]] = None
        else:
            out[k["id"]] = sum(1 for i in idx if k["pres"][i] in "1p") / len(idx)
    return out


# ------------------------------------------------------------------ C
FACT_IDX = {}


def calc_c():
    # 04: полнота и средняя доказательность
    prof = []
    for c in D.BM_COMP:
        blocks = D.BM_PROFILE[c["id"]]
        o = sum(1 for b in blocks if b) / 12
        ev = [f[6] for f in D.FACTS if f[0] == c["id"]]
        p = sum(ev) / len(ev) if ev else 0
        q = ("недостаточно данных" if o < 0.5 else "низкая доказательность" if p < 3 else
             "готов к сравнению" if o >= 0.85 else "требует дополнения")
        prof.append(dict(id=c["id"], name=c["name"], O=o, P=p, Q=q, nfacts=len(ev)))
    pm = {p["id"]: p for p in prof}
    w = D.BM_WEIGHTS
    mat = []
    for c in D.BM_COMP:
        s = D.BM_SCORES[c["id"]]
        n = sum(a * b for a, b in zip(s, w)) / sum(w)
        o, p = pm[c["id"]]["O"], pm[c["id"]]["P"]
        qv = n * 0.6 + p * 0.25 + o * 5 * 0.15
        mat.append(dict(id=c["id"], name=c["name"], N=n, P=p, O=o, Q=qv, sumprod=sum(a * b for a, b in zip(s, w))))
    ncomp = len(D.BM_COMP)
    std = []
    for (pr, blk, desc, cnt, urls) in D.BM_STANDARD:
        e = cnt / ncomp
        f = ("рыночный стандарт" if e >= 0.7 else "формирующийся стандарт" if e >= 0.4 else
             "зона дифференциации" if e >= 0.2 else "слабый сигнал")
        std.append(dict(pr=pr, blk=blk, cnt=cnt, E=e, F=f))
    adv = []
    for (cid, hyp, typ, blk, pages, g, h, i, j, risk) in D.BM_ADV:
        k = g * 0.3 + h * 0.25 + i * 0.25 + j * 0.2
        l = ("сильное преимущество" if k >= 4 else "умеренное преимущество" if k >= 3 else
             "слабое преимущество" if k >= 2 else "недостаточно подтверждений")
        adv.append(dict(id=cid, hyp=hyp, K=k, L=l))
    done = sum(1 for c in D.BM_COMP if c["status"] == "завершен")
    return dict(prof=prof, mat=mat, std=std, adv=adv, ncomp=ncomp, done=done)


# ------------------------------------------------------------------ вывод
def write_csv(name, header, rows):
    os.makedirs(OUT, exist_ok=True)
    with open(os.path.join(OUT, name), "w", encoding="utf-8-sig", newline="") as f:
        wr = csv.writer(f, delimiter=";")
        wr.writerow(header)
        for r in rows:
            wr.writerow(r)


def fmt(x, n=2):
    return "" if x is None else (f"{x:.{n}f}".replace(".", ",") if isinstance(x, float) else str(x))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    A = calc_a()
    print("=== ЧАСТЬ A: 04_Оценка ===")
    for r in A:
        print(f'{r["id"]} {r["name"][:28]:28} {r["scores"]} N={r["N"]:.2f} ({r["sumprod"]}/100) {r["lim"][:18]:18} '
              f'{r["calc"]:38} -> {r["final"]:38} X={r["X"]:.1f} Y={r["Y"]:.1f} {r["zone"]}')
    ns = [round(r["N"], 6) for r in A]
    dup = sorted({n for n in ns if ns.count(n) > 1})
    print("Одинаковые баллы (проверка топ-10):", dup)
    from collections import Counter
    print("Распределение итоговых уровней:", Counter(r["final"] for r in A))
    print("Распределение по зонам:", Counter(r["zone"] for r in A))
    print("Средний балл:", sum(r["N"] for r in A) / len(A))
    write_csv("A_ocenka.csv", ["ID", "Компания"] + [c[0] for c in D.CRIT] + ["СУММПРОИЗВ", "Взвешенный балл", "Ограничитель", "Расчётный уровень",
                                                                              "Корректировка", "Итоговый уровень", "X", "Y", "Зона", "Уверенность"],
              [[r["id"], r["name"]] + r["scores"] + [r["sumprod"], fmt(r["N"]), r["lim"], r["calc"], r["corr"], r["final"],
                                                      fmt(r["X"], 1), fmt(r["Y"], 1), r["zone"], r["conf"]] for r in A])

    B = calc_b()
    print("\n=== ЧАСТЬ B ===")
    print("Левитт: признаков", len(B["lev"]), "по уровням", B["lev_by_level"])
    from collections import Counter as C2
    print("Роли:", C2(r["role"] for r in B["lev"]))
    print("Кано распределение:", B["kano_dist"], "всего", len(B["kano"]))
    for k in B["kano"]:
        print(f'{k["id"]} {k["feat"][:45]:45} {k["cat"]:15} G={k["G"]} H={fmt(k["H"])} {k["interp"]}')
    print("Стандарт:")
    for s in B["std"]:
        print(f'{s["id"]} {s["feat"][:45]:45} {s["cat"]:15} F={s["F"]} G={fmt(s["G"])} {s["status"]}')
    print("06_Сравнение:")
    for c in sorted(B["comp"], key=lambda z: -z["total"]):
        print(f'{c["id"]} {c["name"]:24} {c["total"]:.2f} ({c["sumprod"]}/{c["wsum"]}) {c["interp"]}')
    for cs in B["crit_stats"]:
        print(f'{cs["id"]} avg={cs["avg"]:.2f} best={cs["best"]} leader={cs["leader"]}')
    print("08_Преимущества:")
    for a in B["adv"]:
        print(f'{a["id"]} {a["cid"]} {a["feat"][:45]:45} idx={a["idx"]:.2f} {a["concl"]:28} 0–35: {a["t35"]} {a["st35"]}')
    top3 = [c["id"] for c in sorted(B["comp"], key=lambda z: -z["total"])[:3]]
    ls = leaders_share(B["kano"], top3)
    print("Лидеры (топ-3 по 06):", top3)
    write_csv("B_levitt.csv", ["ID", "Уровень", "Признак", "Наличие", "Сила", "Понятность", "Влияние", "Оценка", "Роль"],
              [[r["id"], r["level"], r["feat"], r["H"], r["I"], r["J"], r["K"], fmt(r["L"]), r["role"]] for r in B["lev"]])
    write_csv("B_kano.csv", ["ID", "Признак", "Категория", "Вес", "Кол-во", "Доля", "Доля лидеров", "Сила", "Интерпретация", "Решение"],
              [[k["id"], k["feat"], k["cat"], k["w"], k["G"], fmt(k["H"]), fmt(ls[k["id"]]), fmt(k["I"], 1), k["interp"], k["dec"]] for k in B["kano"]])
    write_csv("B_sravnenie.csv", ["ID", "Конкурент", "СУММПРОИЗВ", "Сумма весов", "Итог", "Интерпретация"],
              [[c["id"], c["name"], c["sumprod"], c["wsum"], fmt(c["total"]), c["interp"]] for c in B["comp"]])
    write_csv("B_preimushchestva.csv", ["ID", "Конкурент", "Признак", "Редкость", "Сила", "Ценность", "Доказательство", "Индекс", "Вывод",
                                        "Сумма 0–35", "Статус 0–35"],
              [[a["id"], a["cid"], a["feat"], a["R"], a["S"], a["V"], a["Q"], fmt(a["idx"]), a["concl"], a["t35"], a["st35"]] for a in B["adv"]])

    Cc = calc_c()
    print("\n=== ЧАСТЬ C ===")
    for p in Cc["prof"]:
        print(f'{p["id"]} {p["name"]:22} полнота={p["O"]:.2f} доказ={p["P"]:.2f} фактов={p["nfacts"]} {p["Q"]}')
    for m in sorted(Cc["mat"], key=lambda z: -z["Q"]):
        print(f'{m["id"]} {m["name"]:22} N={m["N"]:.2f} ({m["sumprod"]}/100) Q={m["Q"]:.3f}')
    for s in Cc["std"]:
        print(f'{s["pr"][:50]:50} {s["cnt"]}/8={s["E"]:.3f} {s["F"]}')
    for a in Cc["adv"]:
        print(f'{a["id"]} {a["hyp"][:50]:50} {a["K"]:.2f} {a["L"]}')
    print("Завершённых профилей:", Cc["done"], "из", Cc["ncomp"])
    write_csv("C_matrica.csv", ["ID", "Конкурент", "СУММПРОИЗВ", "Индекс силы", "Доказательность", "Полнота", "Итог"],
              [[m["id"], m["name"], m["sumprod"], fmt(m["N"]), fmt(m["P"]), fmt(m["O"]), fmt(m["Q"], 3)] for m in Cc["mat"]])
    write_csv("C_standart.csv", ["Практика", "Блок", "Кол-во", "Доля", "Статус"],
              [[s["pr"], s["blk"], s["cnt"], fmt(s["E"], 3), s["F"]] for s in Cc["std"]])
    print("\nCSV сохранены в", OUT)


if __name__ == "__main__":
    main()

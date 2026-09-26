# -*- coding: utf-8 -*-
"""
ЛР1 «Анализ интереса аудитории и фиксация границ рынка» — проект NotaCode.

Что делает скрипт:
  1) загружает реальный помесячный ряд Google Trends (PlantUML и text to diagram,
     Worldwide, 5 лет, снято 2026-09-17, 40 месяцев 2021-09 … 2024-12) из архивного
     xlsx старой версии (DiagramCode = прежнее имя NotaCode); если файла нет —
     берёт встроенную копию тех же чисел;
  2) считает показатели тренда, сезонные индексы и циклические коэффициенты
     ТОЧНО по логике Excel-шаблона методички (скользящее среднее 12 мес. «назад»,
     Y/T, сезонный индекс = средний Y/T месяца / среднее 12 месяцев,
     K = Y / (T × S), пороги 0,9 / 1,1);
  3) сохраняет рисунки в ../Рисунки/ и таблицу результатов в ../Рисунки/lr1_results.csv.

Запуск:  python lr1_gt_analysis.py
Зависимости: pandas, numpy, matplotlib, openpyxl.
"""
from pathlib import Path
import numpy as np
import pandas as pd
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt

HERE = Path(__file__).resolve().parent
OUT = HERE.parent / "Рисунки"
OUT.mkdir(exist_ok=True)
ARCHIVE_XLSX = (HERE.parent.parent / "_архив_DiagramCode" / "labs" / "lab1-demand"
                / "ЛР1_Тренд_Сезонность_DiagramCode.xlsx")

plt.rcParams.update({
    "font.family": ["Arial", "DejaVu Sans"],
    "font.size": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.edgecolor": "#8a8984",
    "axes.labelcolor": "#52514e",
    "xtick.color": "#52514e",
    "ytick.color": "#52514e",
    "axes.grid": True,
    "axes.axisbelow": True,
    "grid.color": "#e6e5e1",
    "grid.linewidth": 0.8,
    "figure.facecolor": "#fcfcfb",
    "axes.facecolor": "#fcfcfb",
    "savefig.facecolor": "#fcfcfb",
})
BLUE, ORANGE, AQUA, GRAY, INK = "#2a78d6", "#eb6834", "#1baf7a", "#8a8984", "#0b0b0b"
MONTHS_RU = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]

# Встроенная копия реальных данных (тот же замер, что в архивном xlsx).
EMBEDDED = {
    "month": pd.period_range("2021-09", "2024-12", freq="M").astype(str).tolist(),
    "PlantUML": [16, 20, 21, 20, 22, 28, 25, 25, 27, 22, 26, 22, 24, 25, 28, 26, 25, 33, 31, 30,
                 31, 30, 33, 26, 30, 34, 38, 34, 37, 35, 46, 48, 47, 49, 40, 37, 39, 50, 51, 53],
    "text to diagram": [9, 7, 10, 8, 11, 15, 12, 9, 9, 8, 4, 3, 9, 7, 7, 7, 7, 8, 7, 9,
                        8, 7, 5, 5, 8, 8, 8, 9, 7, 9, 9, 11, 10, 6, 5, 6, 8, 10, 11, 12],
}


def load_data() -> tuple[pd.DataFrame, str]:
    if ARCHIVE_XLSX.exists():
        df = pd.read_excel(ARCHIVE_XLSX, sheet_name="01_Данные_и_расчёт", header=1,
                           usecols="A:C", nrows=40)
        df.columns = ["month", "PlantUML", "text to diagram"]
        src = f"архивный файл {ARCHIVE_XLSX.name}"
    else:
        df = pd.DataFrame(EMBEDDED)
        src = "встроенная копия данных"
    df["date"] = pd.to_datetime(df["month"].astype(str).str[:7] + "-01")
    return df, src


def template_decomposition(y: np.ndarray, months: np.ndarray, centered: bool = False) -> dict:
    """Повторяет формулы листов Расчет/Сезонность/Оценка_тренда/Деловые_циклы.

    centered=False — как в шаблоне (скользящее среднее 12 последних месяцев);
    centered=True  — поправка: центрированное скользящее среднее 2×12
    (не отстаёт от растущего ряда, см. раздел «Замечание к шаблону» в ОТЧЕТ.md).
    """
    s = pd.Series(y, dtype=float)
    if centered:
        trend = s.rolling(12).mean().rolling(2).mean().shift(-6)
    else:
        trend = s.rolling(12).mean()                   # Расчет!H: AVERAGE последних 12
    yt = s / trend                                     # Расчет!I
    c = yt.groupby(months).mean()                      # Сезонность!C (среднее Y/T месяца)
    si = c / c.mean()                                  # Сезонность!D
    si_row = si.reindex(months).to_numpy()
    sa = s / si_row                                    # Расчет!J
    k = s / (trend * si_row)                           # Расчет!K
    phase = np.where(k.isna(), "", np.where(k < 0.9, "ниже тренда",
                                            np.where(k > 1.1, "выше тренда", "норма")))
    n = len(s)
    first12, last12 = s[:12].mean(), s[-12:].mean()
    rel = (last12 - first12) / first12
    slope = np.polyfit(np.arange(1, n + 1), s, 1)[0]
    ampl = si.max() - si.min()
    return dict(series=s, trend=trend, yt=yt, si=si, sa=sa, k=k, phase=phase,
                n=n, first12=first12, last12=last12, abs_change=last12 - first12,
                rel=rel, slope=slope, ampl=ampl,
                trend_class=("растущий тренд" if rel >= 0.15 else
                             "снижающийся тренд" if rel <= -0.15 else "стабильный тренд"),
                season_class=("выраженная сезонность" if ampl >= 0.30 else
                              "умеренная сезонность" if ampl >= 0.15 else "слабая сезонность"),
                max_month=int(si.idxmax()), min_month=int(si.idxmin()))


def classic_seasonal_index(y: np.ndarray, months: np.ndarray) -> pd.Series:
    """Формула Word-методики: среднее месяца по всем годам / среднее всех месяцев × 100."""
    s = pd.Series(y, dtype=float)
    return s.groupby(months).mean() / s.mean() * 100


def main() -> None:
    df, src = load_data()
    months = df["date"].dt.month.to_numpy()
    res = {name: template_decomposition(df[name].to_numpy(), months)
           for name in ["PlantUML", "text to diagram"]}
    p = res["PlantUML"]
    pc = template_decomposition(df["PlantUML"].to_numpy(), months, centered=True)

    # ---------- Рис. 1. Динамика GT + скользящее среднее 12 мес. ----------
    fig, ax = plt.subplots(figsize=(10, 4.8))
    ax.plot(df["date"], df["PlantUML"], color=BLUE, lw=2, label="PlantUML — факт")
    ax.plot(df["date"], p["trend"], color=BLUE, lw=2, ls="--",
            label="PlantUML — скользящее среднее 12 мес.")
    t = res["text to diagram"]
    ax.plot(df["date"], df["text to diagram"], color=ORANGE, lw=2, label="text to diagram — факт")
    ax.plot(df["date"], t["trend"], color=ORANGE, lw=2, ls="--",
            label="text to diagram — скользящее среднее 12 мес.")
    ax.annotate(f"{int(df['PlantUML'].iloc[-1])}", (df["date"].iloc[-1], df["PlantUML"].iloc[-1]),
                xytext=(6, 0), textcoords="offset points", va="center", color=INK)
    ax.annotate(f"{int(df['PlantUML'].iloc[0])}", (df["date"].iloc[0], df["PlantUML"].iloc[0]),
                xytext=(-16, 0), textcoords="offset points", va="center", color=INK)
    ax.set_ylim(0, 60)
    ax.set_ylabel("Индекс Google Trends, 0–100")
    ax.set_title("Интерес к категории «diagram as code» (Google Trends, весь мир, "
                 "2021-09…2024-12)", loc="left", fontsize=11, color=INK)
    ax.legend(loc="upper left", frameon=False, fontsize=9)
    fig.text(0.01, 0.01, "Источник: Google Trends, замер PlantUML vs text to diagram, 5 лет, "
             "Worldwide, снято 2026-09-17; выборка — первая неделя месяца.",
             fontsize=8, color="#52514e")
    fig.tight_layout(rect=(0, 0.03, 1, 1))
    fig.savefig(OUT / "fig01_gt_dinamika_ma12.png", dpi=150)
    plt.close(fig)

    # ---------- Рис. 2. Сезонные индексы по месяцам ----------
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.2), sharey=True)
    for ax, name, col in zip(axes, ["PlantUML", "text to diagram"], [BLUE, ORANGE]):
        si = res[name]["si"].reindex(range(1, 13))
        colors = [col if v >= 1 else "#9ec5f4" if col == BLUE else "#f5b79c" for v in si]
        ax.bar(range(1, 13), si.values, color=colors, width=0.7, edgecolor="#fcfcfb", lw=2)
        ax.axhline(1.0, color=GRAY, lw=1)
        ax.axhline(1.1, color=GRAY, lw=0.8, ls=":")
        ax.axhline(0.9, color=GRAY, lw=0.8, ls=":")
        ax.set_xticks(range(1, 13), MONTHS_RU)
        ax.set_ylim(0.5, 1.4)
        ax.grid(axis="x", visible=False)
        ax.set_title(f"{name}: амплитуда {res[name]['ampl']:.3f} — {res[name]['season_class']}",
                     fontsize=10, color=INK, loc="left")
        for m in (res[name]["max_month"], res[name]["min_month"]):
            ax.annotate(f"{si[m]:.2f}", (m, si[m]), xytext=(0, 3), textcoords="offset points",
                        ha="center", fontsize=8, color=INK)
    axes[0].set_ylabel("Сезонный индекс (1,0 = средний уровень)")
    fig.suptitle("Сезонные индексы по месяцам (метод Excel-шаблона: средний Y/T месяца "
                 "/ среднее), пунктир — пороги 0,9 и 1,1", x=0.01, ha="left", fontsize=11, color=INK)
    fig.tight_layout()
    fig.savefig(OUT / "fig02_sezonnye_indeksy.png", dpi=150)
    plt.close(fig)

    # ---------- Рис. 3. Декомпозиция PlantUML ----------
    fig, axes = plt.subplots(4, 1, figsize=(10, 9), sharex=True)
    axes[0].plot(df["date"], p["series"], color=BLUE, lw=2)
    axes[0].set_title("Наблюдаемый ряд Y (индекс GT)", loc="left", fontsize=10)
    axes[1].plot(df["date"], p["trend"], color=BLUE, lw=2)
    axes[1].set_title("Тренд T — скользящее среднее 12 мес.", loc="left", fontsize=10)
    axes[2].plot(df["date"], p["si"].reindex(months).to_numpy(), color=AQUA, lw=2, marker="o", ms=3)
    axes[2].axhline(1, color=GRAY, lw=1)
    axes[2].set_title("Сезонная составляющая S (индекс месяца)", loc="left", fontsize=10)
    k = p["k"]
    axes[3].plot(df["date"], k, color=ORANGE, lw=2, marker="o", ms=3, label="K по шаблону (MA-12 «назад»)")
    axes[3].plot(df["date"], pc["k"], color=GRAY, lw=2, ls="--", label="K при центрированной MA 2×12")
    axes[3].legend(frameon=False, fontsize=8, loc="lower left")
    axes[3].axhline(1, color=GRAY, lw=1)
    axes[3].axhspan(0.9, 1.1, color="#f0efec", zorder=0)
    axes[3].set_title("Циклический коэффициент K = Y / (T × S); серая зона 0,9–1,1 = «норма»",
                      loc="left", fontsize=10)
    for ax in axes:
        ax.tick_params(labelsize=9)
    fig.suptitle("Мультипликативная декомпозиция ряда PlantUML: тренд × сезонность × цикл",
                 x=0.01, ha="left", fontsize=11, color=INK)
    fig.tight_layout()
    fig.savefig(OUT / "fig03_dekompoziciya.png", dpi=150)
    plt.close(fig)

    # ---------- Рис. 4. Региональный срез ----------
    fig, axes = plt.subplots(1, 2, figsize=(11, 4.2))
    # PlantUML: зафиксирован только порядок топ-5 (значения индекса не записаны)
    reg_p = ["Боливия", "Китай", "Япония", "Россия", "Беларусь"]
    heights = [5, 4, 3, 2, 1]
    cols = [BLUE if r in ("Россия", "Беларусь") else "#b7d3f6" for r in reg_p]
    axes[0].barh(reg_p[::-1], heights[::-1], color=cols[::-1], height=0.6)
    axes[0].set_xticks([])
    for i, r in enumerate(reg_p[::-1]):
        axes[0].text(0.1, i, f"№{5 - i}", va="center", color="white" if r in ("Россия", "Беларусь")
                     else INK, fontsize=9)
    axes[0].set_title("PlantUML: топ-5 стран из 71 (зафиксирован порядок,\nзначения индекса не "
                      "сохранены)", loc="left", fontsize=10, color=INK)
    axes[0].grid(False)
    reg_t = {"Филиппины": 100, "Шри-Ланка": 25, "США": 25, "Кения": 22, "Непал": 20}
    axes[1].barh(list(reg_t)[::-1], list(reg_t.values())[::-1], color=ORANGE, height=0.6)
    for i, v in enumerate(list(reg_t.values())[::-1]):
        axes[1].text(v + 1.5, i, str(v), va="center", fontsize=9, color=INK)
    axes[1].set_xlim(0, 115)
    axes[1].set_title("text to diagram: топ-5 стран, индекс 0–100", loc="left", fontsize=10,
                      color=INK)
    fig.suptitle("Региональный срез Google Trends (весь мир, 5 лет, снято 2026-09-17)",
                 x=0.01, ha="left", fontsize=11, color=INK)
    fig.text(0.01, 0.01, "Беларусь (№5) и Россия (№4) входят в топ-5 мира по PlantUML; "
             "AI-запрос «text to diagram» сосредоточен в англоязычных IT-хабах.",
             fontsize=8, color="#52514e")
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    fig.savefig(OUT / "fig04_regionalnyy_srez.png", dpi=150)
    plt.close(fig)

    # ---------- Рис. 5. Средние по годам ----------
    ann = df.groupby(df["date"].dt.year)[["PlantUML", "text to diagram"]].mean()
    fig, ax = plt.subplots(figsize=(8, 4))
    x = np.arange(len(ann))
    ax.bar(x - 0.18, ann["PlantUML"], 0.34, color=BLUE, label="PlantUML")
    ax.bar(x + 0.18, ann["text to diagram"], 0.34, color=ORANGE, label="text to diagram")
    for i, (a, b) in enumerate(zip(ann["PlantUML"], ann["text to diagram"])):
        ax.text(i - 0.18, a + 0.8, f"{a:.1f}", ha="center", fontsize=8)
        ax.text(i + 0.18, b + 0.8, f"{b:.1f}", ha="center", fontsize=8)
    ax.set_xticks(x, [f"{y}{' (4 мес.)' if y == 2021 else ''}" for y in ann.index])
    ax.set_ylabel("Средний индекс GT")
    ax.legend(frameon=False)
    ax.set_title("Средний индекс интереса по годам", loc="left", fontsize=11, color=INK)
    fig.tight_layout()
    fig.savefig(OUT / "fig05_srednie_po_godam.png", dpi=150)
    plt.close(fig)

    # ---------- Таблица результатов ----------
    rows = []
    for i in range(len(df)):
        rows.append({
            "Период": df["month"].iloc[i],
            "PlantUML_Y": p["series"].iloc[i],
            "PlantUML_T_MA12": round(p["trend"].iloc[i], 3) if pd.notna(p["trend"].iloc[i]) else "",
            "PlantUML_Y/T": round(p["yt"].iloc[i], 4) if pd.notna(p["yt"].iloc[i]) else "",
            "PlantUML_S": round(p["si"][months[i]], 4),
            "PlantUML_K": round(p["k"].iloc[i], 4) if pd.notna(p["k"].iloc[i]) else "",
            "PlantUML_Фаза": p["phase"][i],
            "t2d_Y": t["series"].iloc[i],
            "t2d_T_MA12": round(t["trend"].iloc[i], 3) if pd.notna(t["trend"].iloc[i]) else "",
            "t2d_K": round(t["k"].iloc[i], 4) if pd.notna(t["k"].iloc[i]) else "",
        })
    pd.DataFrame(rows).to_csv(OUT / "lr1_results.csv", index=False, encoding="utf-8-sig", sep=";")

    print(f"Данные: {src}; наблюдений {len(df)} ({df['month'].iloc[0]} … {df['month'].iloc[-1]})")
    for name, r in res.items():
        s = r["series"]
        print(f"\n=== {name} ===")
        print(f"среднее {s.mean():.2f}; max {s.max():.0f} ({df['month'][s.idxmax()]}); "
              f"min {s.min():.0f} ({df['month'][s.idxmin()]})")
        print(f"первые 12 мес. {r['first12']:.2f}; последние 12 мес. {r['last12']:.2f}; "
              f"абс. {r['abs_change']:.2f}; отн. {r['rel']*100:.1f}% -> {r['trend_class']}")
        print(f"наклон {r['slope']:.3f} п./мес.; посл. 3 мес. {s[-3:].mean():.2f} к пред. 3 "
              f"{s[-6:-3].mean():.2f} ({(s[-3:].mean()/s[-6:-3].mean()-1)*100:.1f}%)")
        print("сезонные индексы:", {MONTHS_RU[m - 1]: round(v, 3) for m, v in r["si"].items()})
        print(f"амплитуда {r['ampl']:.3f} -> {r['season_class']}; max {MONTHS_RU[r['max_month']-1]},"
              f" min {MONTHS_RU[r['min_month']-1]}")
        print("классический индекс (x100):",
              {MONTHS_RU[m - 1]: round(v, 1) for m, v in classic_seasonal_index(s, months).items()})
        kk = r["k"]
        print(f"K: последний {kk.iloc[-1]:.3f} ({r['phase'][-1]}); min {kk.min():.3f} "
              f"({df['month'][kk.idxmin()]}); max {kk.max():.3f} ({df['month'][kk.idxmax()]})")
        ph = pd.Series(r["phase"][11:]).value_counts().to_dict()
        print("фазы:", ph)
        print("выше/ниже тренда:",
              [(df['month'][i], r['phase'][i], round(kk.iloc[i], 2)) for i in range(len(df))
               if r['phase'][i] in ('выше тренда', 'ниже тренда')])
        print("годовые средние:", s.groupby(df["date"].dt.year.values).mean().round(2).to_dict())
        up = (np.diff(s) > 0).sum()
        print(f"доля месяцев роста {up}/{len(s)-1} = {up/(len(s)-1)*100:.1f}%")
    print("\n=== PlantUML, поправка: центрированная MA 2x12 ===")
    print("сезонные индексы:", {MONTHS_RU[m - 1]: round(v, 3) for m, v in pc["si"].items()})
    print(f"амплитуда {pc['ampl']:.3f} -> {pc['season_class']}; max {MONTHS_RU[pc['max_month']-1]},"
          f" min {MONTHS_RU[pc['min_month']-1]}")
    kc = pc["k"]
    print(f"K: min {kc.min():.3f} ({df['month'][kc.idxmin()]}); max {kc.max():.3f} "
          f"({df['month'][kc.idxmax()]}); последний рассчитанный {kc.dropna().iloc[-1]:.3f} "
          f"({df['month'][kc.dropna().index[-1]]})")
    print("фазы:", pd.Series(pc["phase"]).replace("", np.nan).dropna().value_counts().to_dict())
    print("выше/ниже:", [(df['month'][i], str(pc['phase'][i]), round(float(kc.iloc[i]), 2))
                          for i in range(len(df)) if pc['phase'][i] in ('выше тренда', 'ниже тренда')])
    print(f"\nРисунки сохранены в {OUT}")


if __name__ == "__main__":
    main()

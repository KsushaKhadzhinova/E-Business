from docx import Document
import re, os

base = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
for lr in ["ЛР1", "ЛР2-3", "ЛР4", "ЛР5", "ЛР6", "ЛР7"]:
    d = Document(os.path.join(base, lr, "ОТЧЕТ_%s_NotaCode.docx" % lr))
    paras = [p.text for p in d.paragraphs]
    tabs = [int(m.group(1)) for t in paras for m in [re.match(r"^Таблица (\d+) [–-]", t)] if m]
    figs = [int(m.group(1)) for t in paras for m in [re.match(r"^Рисунок (\d+) [–-]", t)] if m]

    def gaps(seq):
        if not seq:
            return "нет"
        exp = list(range(1, max(seq) + 1))
        missing = [x for x in exp if x not in seq]
        dup = sorted({x for x in seq if seq.count(x) > 1})
        order = seq == sorted(seq)
        return "всего %d, пропуски %s, дубли %s, по порядку %s" % (len(seq), missing[:8], dup[:8], order)

    sec = d.sections[0]
    mm = [round(x / 36000) for x in (sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin)]
    bad_chars = sum(1 for t in paras if any(c in t for c in ("\u0451", "\u2014", "\U0001F532")))
    dbl = sum(1 for t in paras if "  " in t)
    print(lr, "| таблицы:", gaps(tabs), "| рисунки:", gaps(figs), "| поля мм:", mm, "| bad chars paras:", bad_chars, "| двойные пробелы:", dbl)

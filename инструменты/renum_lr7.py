from docx import Document
import re

p = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР7\ОТЧЕТ_ЛР7_NotaCode.docx"
d = Document(p)
n = 0
for par in d.paragraphs:
    m = re.match(r"^Таблица (\d+) ([–-])", par.text)
    if not m:
        continue
    new = int(m.group(1)) - 1
    done = False
    for r in par.runs:
        if re.search(r"Таблица \d+", r.text):
            r.text = re.sub(r"Таблица \d+", "Таблица %d" % new, r.text, count=1)
            done = True
            break
    if not done:
        full = re.sub(r"^Таблица \d+", "Таблица %d" % new, par.text)
        par.runs[0].text = full
        for r in par.runs[1:]:
            r._r.getparent().remove(r._r)
    n += 1
d.save(p)
print("captions renumbered", n)

from docx import Document
from docx.oxml.ns import qn
import glob, os

base = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
for p in glob.glob(base + r"\**\*.docx", recursive=True):
    if os.path.basename(p).startswith("~"):
        continue
    d = Document(p)
    st = d.settings.element
    x = st.find(qn("w:updateFields"))
    if x is not None:
        st.remove(x)
        d.save(p)
    toc = sum(1 for q in d.paragraphs if q.style.name.lower().startswith(("toc", "оглавление")))
    print(os.path.relpath(p, base), "flag removed" if x is not None else "no flag", "| toc paragraphs:", toc)

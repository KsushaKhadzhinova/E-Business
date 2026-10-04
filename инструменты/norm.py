from docx import Document
from docx.oxml.ns import qn
import os
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
M={"\u2014":"\u2013","\u0451":"\u0435","\u0401":"\u0415"}
def fix(root):
    n=0
    for t in root.iter(qn('w:t')):
        if t.text and any(k in t.text for k in M):
            s=t.text
            for k,v in M.items(): s=s.replace(k,v)
            t.text=s; n+=1
    return n
for lr in ["ЛР1","ЛР2-3","ЛР4","ЛР5","ЛР6","ЛР7"]:
    p=os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"); d=Document(p)
    n=fix(d.element)
    for s in d.sections:
        for part in (s.header,s.footer,s.first_page_header,s.first_page_footer):
            try: n+=fix(part._element)
            except Exception: pass
    d.save(p); print(lr,"runs changed:",n)

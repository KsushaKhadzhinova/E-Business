from docx import Document
from docx.oxml.ns import qn
import os
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"; box=chr(0x1F532)
def fix(root):
    n=0
    for t in root.iter(qn('w:t')):
        if t.text and box in t.text:
            s=t.text.replace(box+" ДОСНЯТЬ","не собрано").replace(box+" ","").replace(box,""); t.text=s; n+=1
    return n
for lr in ["ЛР1","ЛР2-3","ЛР4","ЛР5","ЛР6","ЛР7"]:
    p=os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"); d=Document(p); n=fix(d.element); d.save(p); print(lr,n)

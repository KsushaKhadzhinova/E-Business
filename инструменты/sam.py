from docx import Document
from collections import Counter
d=Document(r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР1\ОТЧЕТ_ЛР1_NotaCode.docx")
on=False;c=Counter();longs=[]
for p in d.paragraphs:
    if p.style.name=="Heading 2":
        on=p.text.startswith("2.3 ")
        if p.text.startswith("2.4"): break
    if on and p.text.strip():
        c[p.style.name]+=1; longs.append((len(p.text),p.style.name,p.text[:110]))
print(c); 
for x in sorted(longs,reverse=True)[:6]: print(x)
print(len([1 for x in longs if x[0]<80]),"short paras")
for x in longs[:25]: print(x[1],"|",x[2][:90])

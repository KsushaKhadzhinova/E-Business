from docx import Document
from docx.oxml.ns import qn
d=Document(r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР7\ОТЧЕТ_ЛР7_NotaCode.docx")
body=list(d.element.body.iterchildren()); shown=0
for i,k in enumerate(body):
    if k.tag==qn('w:tbl'):
        prev=body[i-1]; txt=''.join(t.text or '' for t in prev.iter(qn('w:t')))[:90]
        nxt=body[i+1]; nt=''.join(t.text or '' for t in nxt.iter(qn('w:t')))[:60]
        print(i,"TABLE; prev:",repr(txt),"| next:",repr(nt)); shown+=1
        if shown>=4: break
for p in d.paragraphs:
    if "Таблица 1" in p.text or "табл. 1" in p.text or "таблице 1" in p.text: print("ref:",p.text[:120])

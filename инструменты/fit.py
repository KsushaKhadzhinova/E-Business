from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os,re
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
TOTAL=9638
def fit(tb):
    t=tb._tbl; rows=tb.rows; n=len(tb.columns)
    if n<2: return
    sc=[]
    for j in range(n):
        lens=[];words=[]
        for r in rows:
            try: c=r.cells[j]
            except IndexError: continue
            s=c.text.strip(); lens.append(len(s)); words+= [len(w) for w in s.split()] or [0]
        avg=sum(lens)/max(1,len(lens)); lw=max(words) if words else 3
        sc.append(max(avg**0.8*1.6, lw*1.0, 5))
    tot=sum(sc); w=[max(650,int(TOTAL*x/tot)) for x in sc]
    diff=TOTAL-sum(w); w[w.index(max(w))]+=diff
    grid=t.find(qn('w:tblGrid'))
    for g,x in zip(grid.findall(qn('w:gridCol')),w): g.set(qn('w:w'),str(x))
    tw=t.tblPr.find(qn('w:tblW')); tw.set(qn('w:w'),str(TOTAL)); tw.set(qn('w:type'),'dxa')
    for r in rows:
        for j,c in enumerate(r._tr.findall(qn('w:tc'))):
            if j<n:
                pr=c.get_or_add_tcPr(); cw=pr.find(qn('w:tcW'))
                if cw is None: cw=OxmlElement('w:tcW'); pr.insert(0,cw)
                cw.set(qn('w:w'),str(w[j])); cw.set(qn('w:type'),'dxa')
    # header row repeat
    trPr=rows[0]._tr.get_or_add_trPr()
    if trPr.find(qn('w:tblHeader')) is None: trPr.append(OxmlElement('w:tblHeader'))
    for r in rows:
        tp=r._tr.get_or_add_trPr()
        if tp.find(qn('w:cantSplit')) is None: tp.insert(0,OxmlElement('w:cantSplit'))
for lr in ["ЛР1","ЛР2-3","ЛР4","ЛР5","ЛР6","ЛР7"]:
    p=os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"); d=Document(p)
    k=0
    for tb in d.tables:
        try: fit(tb); k+=1
        except Exception as e: print(lr,"skip",e)
    d.save(p); print(lr,"tables fitted",k)

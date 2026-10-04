from docx import Document
from docx.oxml.ns import qn
import os,re,copy,sys
from docx.text.paragraph import Paragraph
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
NOTE_APP="Подробные данные приложения приведены в полной версии отчета и в рабочих книгах Excel лабораторной работы."
NOTE_TBL="Полные данные таблицы приведены в рабочих книгах Excel лабораторной работы и в полной версии отчета."
def is_h1(el):
    st=el.xpath('./w:pPr/w:pStyle/@w:val'); return el.tag==qn('w:p') and st and st[0].lower() in('heading1','1')
def txt(el): return ''.join(el.itertext()).strip()
def body_par(d,text,like):
    new=copy.deepcopy(like); 
    for r in new.xpath('./w:r')[1:]: new.remove(r)
    p=Paragraph(new,None); p.runs[0].text=text; return new
def cut_appendix(d,name):
    body=d.element.body; els=list(body.iterchildren()); n=0
    for i,el in enumerate(els):
        if is_h1(el) and txt(el).startswith(name):
            j=i+1; keep=0; like=None
            while j<len(els) and not is_h1(els[j]) and els[j].tag!=qn('w:sectPr'):
                e=els[j]
                if e.tag==qn('w:p') and txt(e) and keep<2: keep+=1; j+=1; continue
                if like is None and e.tag==qn('w:p') and txt(e): like=e
                if e.tag==qn('w:p') and not txt(e) and e.xpath('.//w:sectPr'): j+=1; continue
                body.remove(e); n+=1; j+=1
            last=els[i+1+ (1 if keep else 0)] if False else None
            # add note after kept paragraphs
            anchor=els[i]
            k=i+1; cnt=0
            while k<len(els) and cnt<keep:
                if els[k].getparent() is not None and els[k].tag==qn('w:p') and txt(els[k]): cnt+=1
                anchor=els[k]; k+=1
            ref=[e for e in els if e.getparent() is not None and e.tag==qn('w:p') and e.xpath('./w:pPr/w:pStyle/@w:val') and e.xpath('./w:pPr/w:pStyle/@w:val')[0] in('BodyText','FirstParagraph')]
            src=ref[0] if ref else anchor
            anchor.addnext(body_par(d,NOTE_APP,src)); return n
    return 0
def replace_tables(d,min_rows,skip_after_h1=("ОТВЕТЫ","ПРОВЕРКА","ВЫВОДЫ","СПИСОК","ПРИЛОЖЕНИЕ")):
    body=d.element.body; els=list(body.iterchildren()); zone=True; n=0
    ref=[e for e in els if e.tag==qn('w:p') and e.xpath('./w:pPr/w:pStyle/@w:val') and e.xpath('./w:pPr/w:pStyle/@w:val')[0] in('BodyText','FirstParagraph')][0]
    for e in els:
        if is_h1(e):
            t=txt(e); zone = not any(t[2:].startswith(s) or t.startswith(s) or re.sub(r'^\d+ ','',t).startswith(s) for s in skip_after_h1)
        elif zone and e.tag==qn('w:tbl') and len(e.xpath('./w:tr'))>=min_rows:
            e.addprevious(body_par(d,NOTE_TBL,ref)); body.remove(e); n+=1
    return n
def shrink_images(d,maxw_cm=11.0):
    from docx.shared import Cm
    for s in d.inline_shapes:
        if s.width.cm>maxw_cm:
            r=maxw_cm/s.width.cm; s.height=int(s.height*r); s.width=Cm(maxw_cm)
cfg={"ЛР2-3":dict(rows=6,img=10.0),"ЛР1":dict(rows=7,img=10.0)}
for lr,c in cfg.items():
    p=os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"); d=Document(p)
    t=replace_tables(d,c["rows"]); shrink_images(d,c["img"]); d.save(p); print(lr,"tables replaced",t)

from docx import Document
from docx.oxml.ns import qn
import os
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
for lr in ["ЛР1","ЛР2-3","ЛР4"]:
    d=Document(os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"))
    secs=[];cur=None
    for el in d.element.body.iterchildren():
        if el.tag==qn('w:p'):
            st=el.xpath('./w:pPr/w:pStyle/@w:val'); st=st[0] if st else ''
            txt=''.join(el.itertext())
            if st in('Heading1','Heading2','1','2') or st.lower().startswith('heading1') or st.lower().startswith('heading2'):
                cur=[txt[:40],0,0,0,0]; secs.append(cur); continue
            if cur is None: continue
            cur[1]+=len(txt)
            if el.xpath('.//w:drawing'): cur[3]+=1
        elif el.tag==qn('w:tbl') and cur is not None:
            cur[2]+=1; cur[4]+=len(el.xpath('./w:tr'))
    print("==",lr)
    for s in secs: print("  %-42s chars=%6d tbl=%2d rows=%3d img=%d"%(s[0],s[1],s[2],s[4],s[3]))

from docx import Document
import re
for lr in ["ЛР1","ЛР7"]:
    d=Document(rf"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\{lr}\ОТЧЕТ_{lr}_NotaCode.docx")
    print(lr,[ (s.style_id,s.name) for s in d.styles if s.type==1 and ("eading" in s.name or "toc" in s.name.lower() or "Оглавление" in s.name)])
    body=d.element.body; kids=list(body.iterchildren())
    for i,k in enumerate(kids[:40]):
        tag=k.tag.split('}')[1]; tx=''.join(k.itertext())[:50]
        st=k.xpath('./w:pPr/w:pStyle/@w:val'); print(i,tag,st[0] if st else '',repr(tx), 'FLD' if k.xpath('.//w:fldChar|.//w:instrText') else '', 'PB' if k.xpath('.//w:br[@w:type="page"]|./w:pPr/w:pageBreakBefore') else '')
    print(len(kids))

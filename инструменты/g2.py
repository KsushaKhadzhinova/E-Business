from docx import Document
d=Document(r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР5\ОТЧЕТ_ЛР5_NotaCode.docx")
ps=d.paragraphs
i=[k for k,p in enumerate(ps) if p.style.name=="Heading 1" and p.text.startswith("1 ЦЕЛЬ")][0]
for q in ps[i:i+6]: print(q.style.name,"|",q.text[:400],"|",q.paragraph_format.first_line_indent, q.alignment)
print(ps[i+1]._p.xml[:1500])

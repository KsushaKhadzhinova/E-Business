from docx import Document
import re
d=Document(r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР2-3\ОТЧЕТ_ЛР2-3_NotaCode.docx")
for n in ["Normal","Body Text","First Paragraph","Compact","Heading 1","Heading 2","Heading 3","Table","Caption"]:
    try:
        s=d.styles[n]; x=s.element.xml
        sp=re.findall(r'<w:spacing[^>]*/>',x); ind=re.findall(r'<w:ind[^>]*/>',x); sz=re.findall(r'<w:sz w:val="\d+"',x)
        print(n,"based:",s.base_style.name if s.base_style else None,sp,ind,sz[:1])
    except KeyError: print(n,"missing")
print(re.findall(r'<w:docDefaults>.*?</w:docDefaults>',d.styles.element.xml,re.S)[0][:600])
for im in d.inline_shapes: print(round(im.width.cm,1),round(im.height.cm,1),end=" | ")

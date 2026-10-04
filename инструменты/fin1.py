from docx import Document
import re
d=Document(r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР1\ОТЧЕТ_ЛР1_NotaCode.docx")
k=list(d.element.body.iterchildren())[35]
x=k.xml
x=re.sub(r' xmlns:\w+="[^"]+"','',x)
print(x[:2500])

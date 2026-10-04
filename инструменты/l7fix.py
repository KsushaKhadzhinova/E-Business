from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os,glob
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
n_ref=0
for p in glob.glob(base+r"\**\*.docx",recursive=True):
    d=Document(p); ch=0
    for t in d.element.iter(qn("w:t")):
        if t.text and "ЛР4, приложение Д" in t.text:
            t.text=t.text.replace("ЛР4, приложение Д","ЛР4, приложение Б"); ch+=1
    st=d.settings.element
    if st.find(qn("w:updateFields")) is None:
        u=OxmlElement("w:updateFields"); u.set(qn("w:val"),"true"); st.append(u)
    d.save(p); n_ref+=ch; print(os.path.relpath(p,base),"refs fixed:",ch)
print("total refs",n_ref)

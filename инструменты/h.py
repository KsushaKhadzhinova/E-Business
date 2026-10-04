from docx import Document
import os,re
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
for lr in ["ЛР1","ЛР2-3","ЛР4","ЛР5","ЛР6","ЛР7"]:
    p=os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx")
    d=Document(p)
    print("==",lr, "sections:",len(d.sections), "margins mm:", [round(x.emu/36000) for x in (d.sections[0].left_margin,d.sections[0].right_margin,d.sections[0].top_margin,d.sections[0].bottom_margin)])
    n=0
    for para in d.paragraphs:
        if para.style.name.startswith("Heading 1"):
            has=para._p.pPr is not None and para._p.pPr.numPr is not None
            print("  H1",repr(para.text[:50]),"numPr" if has else "")
        elif para.style.name.startswith("Heading 2") and n<4:
            n+=1; print("  H2",repr(para.text[:50]))
    st=d.styles['Heading 1']
    print("  H1 numPr in style:", st.element.pPr is not None and st.element.pPr.numPr is not None)

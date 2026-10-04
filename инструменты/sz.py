from docx import Document
import os
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
for lr in ["ЛР1","ЛР2-3","ЛР4"]:
    d=Document(os.path.join(base,lr,f"ОТЧЕТ_{lr}_NotaCode.docx"))
    rows=[len(t.rows) for t in d.tables]
    big=[(i,r) for i,r in enumerate(rows) if r>=18]
    print(lr,"tables",len(rows),"total rows",sum(rows),"big(>=18):",len(big),"rows in big:",sum(r for _,r in big))
    print("  ",[(r) for _,r in big][:40])

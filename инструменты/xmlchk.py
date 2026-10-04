import zipfile,lxml.etree as E
z=zipfile.ZipFile(r"СДАЧА/ЛР1/ОТЧЕТ_ЛР1_NotaCode.docx")
for n in z.namelist():
    if n.endswith((".xml",".rels")):
        try: E.fromstring(z.read(n))
        except Exception as e: print("BAD XML",n,e)
print("xml check done", [ (i.filename,i.file_size) for i in z.infolist()][:30])

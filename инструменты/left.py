import openpyxl,os,re,warnings
warnings.filterwarnings("ignore")
idp=re.compile(r"^[A-Za-zА-Яа-я]{0,4}[-_]?\d{1,4}$")
for r,_,fs in os.walk("СДАЧА"):
    for f in sorted(fs):
        if not f.endswith(".xlsx") or f.startswith(("~","XLT")): continue
        wb=openpyxl.load_workbook(os.path.join(r,f),data_only=True); out=[]
        for ws in wb.worksheets:
            n=0
            for row in ws.iter_rows(min_row=3):
                a=row[0].value
                if a is None or not (isinstance(a,(int,float)) or idp.match(str(a).strip())): continue
                rest=[c.value for c in row[1:]]
                if all(v in (None,"",0,"Не начато") for v in rest): n+=1
            if n>=2: out.append((ws.title,n))
        if out: print(f[:45],out)
print("scan done")

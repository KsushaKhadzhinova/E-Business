import openpyxl,os,warnings
warnings.filterwarnings("ignore")
for r,_,fs in os.walk("СДАЧА"):
    for f in sorted(fs):
        if f.endswith(".xlsx") and not f.startswith("~"):
            n=0
            for ws in openpyxl.load_workbook(os.path.join(r,f)).worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value,str) and ("ДОСНЯТЬ" in c.value or "TODO" in c.value): n+=1
            if n: print(f[:50],n)
print("done")


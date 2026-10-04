import openpyxl,os,re,collections,warnings
warnings.filterwarnings("ignore")
fn=collections.Counter(); forms=collections.Counter(); files=collections.defaultdict(int)
for r,_,fs in os.walk("СДАЧА"):
    for f in fs:
        if not f.endswith(".xlsx") or f.startswith("~"): continue
        wb=openpyxl.load_workbook(os.path.join(r,f))
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    v=c.value; v=getattr(v,"text",v)
                    if isinstance(v,str) and v.startswith("=") and "_xlfn" in v:
                        for m in re.findall(r"_xlfn\.([A-Z.]+)",v): fn[m]+=1
                        files[f]+=1
                        k=re.sub(r"\d+","#",v)
                        forms[k[:230]]+=1
print(dict(fn)); print(dict(files))
for k,v in forms.most_common(12): print(v,k)

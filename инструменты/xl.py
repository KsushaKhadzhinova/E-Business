import zipfile,subprocess,os,io,openpyxl
base="СДАЧА"
def parts(z): 
    n=z.namelist(); return {"charts":sum(1 for x in n if x.startswith("xl/charts/chart")),"draw":sum(1 for x in n if x.startswith("xl/drawings/drawing")),"media":sum(1 for x in n if "media/" in x),"cv":0}
res=[]
for root,_,fs in os.walk(base):
    for f in sorted(fs):
        if not f.endswith(".xlsx"): continue
        p=os.path.join(root,f); rel=p.replace("\\","/")
        z=zipfile.ZipFile(p); cur=parts(z)
        # first commit version
        old=None
        r=subprocess.run(["git","show",f"9735d9f:{rel}"],capture_output=True)
        if r.returncode==0: old=parts(zipfile.ZipFile(io.BytesIO(r.stdout)))
        wb=openpyxl.load_workbook(p,data_only=True); err=0;nonecached=0
        wf=openpyxl.load_workbook(p,data_only=False)
        for ws in wf.worksheets:
            wv=wb[ws.title]
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value,str) and c.value.startswith("="):
                        v=wv[c.coordinate].value
                        if v is None: nonecached+=1
                        elif isinstance(v,str) and v.startswith("#"): err+=1
        print(f"{rel[6:60]:55s} charts {cur['charts']}/{old['charts'] if old else '-'} draw {cur['draw']}/{old['draw'] if old else '-'} errs={err} nocache={nonecached}")

import openpyxl,os,re,warnings
from openpyxl.formula import Tokenizer
from openpyxl.utils import range_boundaries
warnings.filterwarnings("ignore")
bad=0
for r,_,fs in os.walk("СДАЧА"):
    for f in sorted(fs):
        if not f.endswith(".xlsx") or f.startswith("~"): continue
        wb=openpyxl.load_workbook(os.path.join(r,f)); n=0; ex=[]
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    v=getattr(c.value,"text",c.value)
                    if not (isinstance(v,str) and v.startswith("=")): continue
                    try: toks=Tokenizer(v).items
                    except Exception: continue
                    for t in toks:
                        if t.type=="OPERAND" and t.subtype=="RANGE":
                            ref=t.value; sh=None
                            if "!" in ref: sh,ref=ref.rsplit("!",1); sh=sh.strip("'")
                            if sh not in (None,ws.title): continue
                            try: c1,r1,c2,r2=range_boundaries(ref.replace("$",""))
                            except Exception: continue
                            if None in (c1,r1,c2,r2): continue
                            if c1<=c.column<=c2 and r1<=c.row<=r2:
                                n+=1
                                if len(ex)<3: ex.append((ws.title,c.coordinate,v[:80]))
        if n: print(f,n,ex); bad+=n
print("self-refs total",bad)

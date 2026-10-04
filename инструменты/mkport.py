import json,openpyxl,os,warnings,re
warnings.filterwarnings("ignore")
T=r"C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
base=r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"
d=json.load(open(T+r"\fills.json",encoding="utf-8"))
ABBR={"у.е","т.д","т.п","т.е","мес","тыс","млн","руб","шт","стр","рис","табл","др"}
def clean(v):
    if not isinstance(v,str): return v
    v=v.replace("\u2014","\u2013").replace("\u0451","\u0435").replace("\u0401","\u0415").replace("Хаджынова К.А","К. А. Хаджинова")
    v=v.replace(chr(0x1F532)+" ДОСНЯТЬ","не собрано").replace(chr(0x1F532)+" ","")
    if v.endswith(".") and not v.endswith("..") and len(v)>3:
        m=re.search(r"([A-Za-zА-Яа-я]+)$",v[:-1]); last=m.group(1).lower() if m else ""
        if not (last in ABBR and len(last)<=4) and not v[:-1].endswith("у.е"): v=v[:-1]
    return v
out={}; tot=0
for root,_,fs in os.walk(base):
    for f in fs:
        if f in d:
            wb=openpyxl.load_workbook(os.path.join(root,f)); lst=[]
            for s,c,x in d[f]:
                if s not in wb.sheetnames: continue
                ws=wb[s]; row=int(re.sub(r"\D","",c))
                if row>ws.max_row: continue
                if isinstance(x,str) and x.strip() in ("□",""): continue
                if isinstance(x,str) and re.match(r"^\d{4}-\d\d-\d\d",x): continue
                lst.append({"p":os.path.join(root,f),"s":s,"c":c,"v":clean(x)})
            out[f]=lst; tot+=len(lst); print(f[:40],len(lst))
json.dump(out,open(T+r"\port.json","w",encoding="utf-8"),ensure_ascii=False)
print("total",tot)

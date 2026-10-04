import json,collections
T=r"C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
d=json.load(open(T+r"\fills.json",encoding="utf-8"))
for f,v in d.items():
    by=collections.defaultdict(list)
    for s,c,x in v: by[s].append((c,str(x)[:55]))
    print("==",f[:40])
    for s,l in by.items(): print("  ",s,len(l),l[:3])

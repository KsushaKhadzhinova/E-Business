import json
T=r"C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
o=json.load(open(T+r"\other.json",encoding="utf-8"))
seen=set()
for f,s,c,a,b in o:
    if a.startswith("="): continue
    if "Хадж" in a or "Хадж" in b or ("руб" in b and "BYN" in a) or a.strip() in ("3",):
        k=(a[:40],b[:40])
        if k in seen: continue
        seen.add(k); print(f[:22],s,c,"| NEW:",a[:70],"| OLD:",b[:70])
        if len(seen)>14: break

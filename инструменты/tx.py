import json,re,collections
T=r"C:\Users\zheny\AppData\Local\Temp\claude\E-----------------------\89f89955-8ca8-4485-9364-f108500ccfc2\scratchpad"
o=json.load(open(T+r"\other.json",encoding="utf-8"))
tx=[x for x in o if not x[3].startswith("=")]
print(len(tx),"non-formula 'other'")
import difflib
kinds=collections.Counter()
for f,s,c,a,b in tx:
    sm=difflib.SequenceMatcher(None,b,a)
    d=[(b[i1:i2],a[j1:j2]) for t,i1,i2,j1,j2 in sm.get_opcodes() if t!="equal"]
    for x in d[:3]: kinds[x]+=1
for k,v in kinds.most_common(40): print(v,k)

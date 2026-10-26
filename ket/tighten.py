import re
WS="C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"
txt=open(f"{WS}/extract_OFFICIAL.txt",encoding="utf-8").read()
lines=txt.splitlines()
start=next(i for i,l in enumerate(lines) if re.match(r'^A$',l.strip()))
end=len(lines)
for i in range(start,len(lines)):
    if re.match(r'^Appendix\s',lines[i].strip(),re.I): end=i; break
head_re=re.compile(r"^([A-Za-z][A-Za-z \-/'.]*?[A-Za-z'])\s*\(([a-z ,&]*[a-z])\)\s*$")
official=[(m.group(1).strip(),m.group(2).strip().lower()) for i in range(start,end) if (m:=head_re.match(lines[i].strip()))]
def norm(w): return re.sub(r'\s+',' ',w.lower().strip())
user_norm=set(norm(u) for u in open(f"{WS}/user_headwords.txt",encoding="utf-8") if u.strip())
GRAM={'det','pron','conj','prep','modal','mv','av','auxiliary','exclam'}
CONT={'n','v','adj'}
def keep(pos):
    toks=re.findall(r'[a-z]+',pos)
    if any(t in GRAM for t in toks): return False
    return any(t in CONT for t in toks)
miss=[(w,p) for (w,p) in official if norm(w) not in user_norm and keep(p)]
print("tightened content missing (n/v/adj, excl grammar):",len(miss))
with open(f"{WS}/missing_tight.txt","w",encoding="utf-8") as f:
    for w,p in miss: f.write(f"{norm(w)}\t{p}\t{w}\n")
ap2=next(i for i,l in enumerate(lines) if re.match(r'^Appendix\s*2',l.strip(),re.I))
print("=== Appendix 2 sample (from line",ap2,") ===")
for i in range(ap2,min(ap2+130,len(lines))):
    print(f"{i:4d}|{lines[i]}")

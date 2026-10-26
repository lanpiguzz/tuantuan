import re

WS="C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"
txt=open(f"{WS}/extract_OFFICIAL.txt",encoding="utf-8").read()
lines=txt.splitlines()

start=None
for i,l in enumerate(lines):
    if re.match(r'^A$', l.strip()):
        start=i; break
end=len(lines)
for i in range(start, len(lines)):
    if re.match(r'^Appendix\s', lines[i].strip(), re.I):
        end=i; break
print("main list region:",start,"-",end)

head_re=re.compile(r"^([A-Za-z][A-Za-z \-/'.]*?[A-Za-z'])\s*\(([a-z ,&]*[a-z])\)\s*$")
official=[]
for i in range(start, end):
    m=head_re.match(lines[i].strip())
    if m:
        official.append((m.group(1).strip(), m.group(2).strip().lower()))

def norm(w): return re.sub(r'\s+',' ',w.lower().strip())
user=[u.strip() for u in open(f"{WS}/user_headwords.txt",encoding="utf-8") if u.strip()]
user_norm=set(norm(u) for u in user)

content={'n','v','adj','adv','phr','phrv'}
def keep(pos):
    toks=re.findall(r'[a-z]+', pos)
    return any(t in content for t in toks)

miss=[(w,pos) for (w,pos) in official if norm(w) not in user_norm and keep(pos)]
print("official parsed:",len(official))
print("content-word missing:",len(miss))
with open(f"{WS}/missing_words.txt","w",encoding="utf-8") as f:
    for w,p in miss: f.write(f"{w}\t{p}\n")
for w,p in miss[:200]:
    print(f"{w}  ({p})")

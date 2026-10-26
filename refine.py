import zipfile, re

WS="C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"
DOCX1="C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表.docx"
DOCX2="C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表2.docx"

def docx_paras(path):
    z=zipfile.ZipFile(path); data=z.read('word/document.xml')
    out=[]
    for p in re.findall(rb'<w:p[ >].*?</w:p>', data, re.S):
        out.append(''.join(t.decode('utf-8','ignore') for t in re.findall(rb'<w:t[^>]*>(.*?)</w:t>', p, re.S)))
    return out

def headwords(path):
    ps=docx_paras(path); hw=[]
    for i,p in enumerate(ps):
        p=p.strip()
        if not p: continue
        nxt=ps[i+1].strip() if i+1<len(ps) else ''
        if re.match(r'^/.+/$', nxt):  # next line is IPA
            hw.append(p)
    return hw

h1=headwords(DOCX1); h2=headwords(DOCX2)
print("DOCX1 headwords:",len(h1),"sample:",h1[:12])
print("DOCX2 headwords:",len(h2),"sample:",h2[:12])
with open(f"{WS}/user_headwords.txt","w",encoding="utf-8") as f:
    for w in h1+h2: f.write(w.lower()+"\n")

# official pdf middle section
txt=open(f"{WS}/extract_OFFICIAL.txt",encoding="utf-8").read()
lines=txt.splitlines()
print("\n--- OFFICIAL lines 90..230 ---")
for i in range(90,min(231,len(lines))):
    print(f"{i:3d}|{lines[i]}")

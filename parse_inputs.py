import zipfile, re

WS="C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"
DOCX1="C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表.docx"
DOCX2="C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表2.docx"
PDF=WS+"/KET材料/备考资料/a2-key-2020-vocabulary-list.pdf"

def docx_rows(path):
    z=zipfile.ZipFile(path); data=z.read('word/document.xml')
    rows=[]
    for tbl in re.findall(rb'<w:tbl>.*?</w:tbl>', data, re.S):
        for tr in re.findall(rb'<w:tr[ >].*?</w:tr>', tbl, re.S):
            cells=[]
            for tc in re.findall(rb'<w:tc>.*?</w:tc>', tr, re.S):
                txt=[t.decode('utf-8','ignore') for t in re.findall(rb'<w:t[^>]*>(.*?)</w:t>', tc, re.S)]
                cells.append(''.join(txt))
            rows.append(cells)
    return rows

def headwords(path):
    rows=docx_rows(path); hw=[]
    for r in rows:
        if not r: continue
        first=r[0].strip()
        # skip header row
        if first in ('单词','音标','词义','例句'): continue
        # headword: primarily english, single token or phrase
        if re.match(r'^[A-Za-z][A-Za-z \-\'/]+$', first) and len(first)<=40:
            hw.append(first)
    return hw

h1=headwords(DOCX1); h2=headwords(DOCX2)
print("DOCX1 headwords:",len(h1))
print("DOCX2 headwords:",len(h2))
print("sample DOCX1:", h1[:15])
print("sample DOCX2:", h2[:15])

with open(f"{WS}/user_headwords.txt","w",encoding="utf-8") as f:
    for w in h1+h2: f.write(w+"\n")

# PDF
from pypdf import PdfReader
r=PdfReader(PDF)
txt="\n".join((p.extract_text() or "") for p in r.pages)
open(f"{WS}/extract_OFFICIAL.txt","w",encoding="utf-8").write(txt)
lines=txt.splitlines()
print("\nOFFICIAL PDF lines:",len(lines))
for i,ln in enumerate(lines[:90]):
    print(f"{i:3d}|{ln}")

import zipfile, re, os

WS = "C:/Users/Administrator/WorkBuddy/2026-10-26-10-40-27"
DOCX1 = "C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表.docx"
DOCX2 = "C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表2.docx"
PDF   = WS + "/KET材料/备考资料/a2-key-2020-vocabulary-list.pdf"

def docx_text(path):
    z = zipfile.ZipFile(path)
    data = z.read('word/document.xml')
    paras = re.findall(rb'<w:p[ >].*?</w:p>', data, re.S)
    out=[]
    for p in paras:
        texts = re.findall(rb'<w:t[^>]*>(.*?)</w:t>', p, re.S)
        s=''.join(t.decode('utf-8','ignore') for t in texts)
        out.append(s)
    # also grab table cell texts in order (already included via p iteration, but ensure)
    return '\n'.join(out)

for label,path in [('DOCX1',DOCX1),('DOCX2',DOCX2)]:
    t=docx_text(path)
    open(f"{WS}/extract_{label}.txt","w",encoding="utf-8").write(t)
    lines=t.splitlines()
    print(f"===== {label} (chars={len(t)}, lines={len(lines)}) =====")
    for ln in lines[:60]:
        print(ln)
    print("...\n")

def pdf_text(path):
    try:
        from pypdf import PdfReader
        r=PdfReader(path); return '\n'.join((p.extract_text() or '') for p in r.pages)
    except Exception as e:
        try:
            from PyPDF2 import PdfReader
            r=PdfReader(path); return '\n'.join((p.extract_text() or '') for p in r.pages)
        except Exception as e2:
            try:
                import pdfplumber
                with pdfplumber.open(path) as pdf:
                    return '\n'.join((pg.extract_text() or '') for pg in pdf.pages)
            except Exception as e3:
                return f"PDF_EXTRACT_FAIL::{e}||{e2}||{e3}"

pt = pdf_text(PDF)
open(f"{WS}/extract_OFFICIAL.txt","w",encoding="utf-8").write(pt)
plines=pt.splitlines()
print(f"===== OFFICIAL PDF (chars={len(pt)}, lines={len(plines)}) =====")
for ln in plines[:80]:
    print(ln)
print("...")

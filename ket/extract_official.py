# -*- coding: utf-8 -*-
"""从剑桥官方 A2 Key / Key for Schools 词汇表 PDF (Aug 2025) 抽取完整词条集合。

PDF 结构：
  p1      封面
  p2-3    Introduction
  p4-23   字母序主词表
  p24     Appendix 1: Word sets
  p25-32  Appendix 2: Topic Lists

主词表条目行有多种形态，必须全部正确处理：
  a/an (det)
  centre/center (n)                       并列写法
  grey (adj & n) (Br Eng) (Am Eng: gray)   英美变体（后置注释）
  mobile (phone) (n)                      词的一部分含括号
  photo(graph) (n)                       括号在词中间
  rubbish (unc n)
  TV / television (n)
所以：**词名 = 行首到最后一个「词性括号」之前**；括号内的注释需剥离。
"""
import re
import pypdf

PDF = "KET材料/备考资料/a2-key-2020-vocabulary-list.pdf"


def load_pages():
    r = pypdf.PdfReader(PDF)
    return [p.extract_text() or "" for p in r.pages]


SKIP_HEADER = re.compile(
    r"^(© UCLES|Vocabulary List|A2 Key and A2 Key for|Schools|Page \d+ of \d+)")

# 一整行都像条目：至少含一个括号
LINE_HAS_PAREN = re.compile(r"\(")
PAREN = re.compile(r"\s*\([^)]*\)")
WORD = re.compile(r"^[A-Za-z][A-Za-z'’\-]*$")
NOTE_HINT = re.compile(
    r"(etc\.|one thousand|to thirty-first|candidates|expected to know|"
    r"the words|in the|of the|are also|include)", re.I)

# 词性标注词表（用于识别「最后一个括号是词性」）
POSWORDS = {
    "det", "adj", "adv", "prep", "pron", "n", "v", "phr", "conj", "av",
    "auxiliary", "exclam", "num", "modal", "unc", "pl", "sing", "abbr",
}


def is_pos_group(s):
    """判断括号内容是否是词性标注组合。

    需要处理复合形态：'phr v'、'prep phr'、'unc n'、'adj – for colours'。
    策略：按 & / , 和空格全部切碎，只要**每一块**都是已知词性标记才算。
    'for colours' 这类说明性尾巴允许出现（视为词性标注的补充说明）。
    """
    s = s.strip().lower()
    if not s:
        return False
    parts = re.split(r"[&/,]| and ", s)
    flat = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        # 去掉 '– for colours' 之类的说明尾巴
        p = re.split(r"[–—-]\s", p)[0].strip()
        # 复合词性再按空格切
        for q in p.split():
            q = q.strip()
            if q:
                flat.append(q)
    if not flat:
        return False
    return all(q in POSWORDS for q in flat)


def parse_entry_line(line):
    """解析一行主词表条目，返回 (head, pos) 或 None。"""
    if not LINE_HAS_PAREN.search(line):
        return None
    # 找所有括号段（group(1) 是不含外层括号的内部内容）
    spans = [(m.start(), m.end(), m.group(1)) for m in re.finditer(r"\(([^)]*)\)", line)]
    if not spans:
        return None
    # 从后往前找第一个「词性括号」
    for st, en, txt in reversed(spans):
        if is_pos_group(txt):
            head = line[:st].strip()
            pos = txt.strip()
            if not head or not re.match(r"^[A-Za-z]", head):
                return None
            return head, pos
    return None


def expand_head(head):
    """把词名展开成多个拼写变体。"""
    out = []
    # 先处理并列写法 'a/an' 'centre/center' 'television / TV'
    for w in re.split(r"\s*/\s*", head):
        w = w.strip().strip(",")
        if not w:
            continue
        # 去掉词内的括号注释：'photo(graph)' -> 'photograph'
        w2 = re.sub(r"\(([^)]*)\)", "", w).strip()
        # 'a/an' -> ['a','an']
        if "/" in w2:
            for p in w2.split("/"):
                p = p.strip()
                if p:
                    out.append(p)
        else:
            out.append(w2)
    return [w for w in out if w]


def split_words(line):
    """从多列内容行切出词条。"""
    cleaned = PAREN.sub("", line).strip()
    out = []
    for w in re.split(r"[,\s]+", cleaned):
        w = w.strip().rstrip(".")
        if not w:
            continue
        if "/" in w:
            for p in w.split("/"):
                p = p.strip()
                if WORD.match(p):
                    out.append(p)
        elif WORD.match(w):
            out.append(w)
    return out


def extract():
    pages = load_pages()

    app1 = app2 = None
    for i, t in enumerate(pages):
        if app1 is None and re.search(r"^Appendix 1\s*$", t, re.M):
            app1 = i
        if app2 is None and re.search(r"^Appendix 2\s*$", t, re.M):
            app2 = i
    assert app1 is not None and app2 is not None, (app1, app2)

    main = {}
    variants = {}   # word -> 变体说明（Am Eng 等）

    # ---------- 主词表 ----------
    for page in pages[:app1]:
        for l in [x.strip() for x in page.split("\n")]:
            if not l or SKIP_HEADER.match(l) or l.startswith("•"):
                continue
            r = parse_entry_line(l)
            if not r:
                continue
            head, pos = r
            for w in expand_head(head):
                main.setdefault(w, pos)
            # 记录英美变体：(Br Eng) (Am Eng: gray)
            am = re.search(r"\(Am Eng:\s*([^)]+)\)", l)
            if am:
                for w in expand_head(head):
                    for v in re.split(r"[,/]", am.group(1)):
                        v = v.strip()
                        if v:
                            variants.setdefault(v, "Am Eng")

    # ---------- Appendix 1 ----------
    lines1 = []
    for page in pages[app1:app2]:
        lines1 += page.split("\n")
    wordsets = {}
    cur = None
    for l in lines1:
        s = l.strip()
        if not s or SKIP_HEADER.match(s) or s.startswith("•"):
            continue
        if s in ("Appendix 1", "Word sets"):
            continue
        if NOTE_HINT.search(s):
            continue
        if len(s) < 60 and re.search(r"[a-z]", s) and not s.endswith(".") and "(" not in s:
            cur = s
            wordsets.setdefault(cur, [])
        elif cur:
            wordsets[cur].extend(split_words(s))

    # ---------- Appendix 2 ----------
    lines2 = []
    for page in pages[app2:]:
        if lines2:
            lines2.append("")
        lines2 += page.split("\n")
    topics = {}
    cur = None
    i = 0
    while i < len(lines2):
        s = lines2[i].strip()
        if not s or SKIP_HEADER.match(s) or s.startswith("•"):
            i += 1
            continue
        if s in ("Appendix 2", "Topic Lists"):
            i += 1
            continue
        if NOTE_HINT.search(s):
            i += 1
            continue
        nxt = lines2[i + 1].strip() if i + 1 < len(lines2) else ""
        is_title = (nxt == "" and len(s) < 45 and re.search(r"[a-z]", s)
                    and not s.endswith((".", ",", ";")))
        if is_title:
            cur = s
            topics.setdefault(cur, [])
            i += 1
            continue
        if cur:
            topics[cur].extend(split_words(s))
        i += 1

    KNOWN = re.compile(r"^[A-Z][A-Za-z:,&'’\- ]*$")
    topics = {k: v for k, v in topics.items()
              if KNOWN.match(k) and not k.startswith(("Appendix", "Topic"))}

    return main, wordsets, topics, variants


def normalize(w):
    return re.sub(r"\s+", " ", w.strip().lower()).rstrip(".")


if __name__ == "__main__":
    main, ws, tp, var = extract()
    print("主词表:", len(main), "词条 | 附录1:", len(ws), "组 | 附录2:", len(tp), "主题 | 变体:", len(var))
    print()
    print("=== 附录2 主题 ===")
    for k, v in tp.items():
        print(f"  {k}  ({len(v)})")
    print()
    probes = ["mug", "bottle", "grandparent", "surfing", "poor", "ski", "mobile",
              "grey", "gray", "TV", "television", "photograph", "photo", "petrol",
              "rubbish", "accommodation", "aeroplane", "taxi", "centre", "center",
              "a", "an", "OK", "okay"]
    for p in probes:
        print(f"  {p:16s} -> {main.get(p, '【缺】')}")
    print()
    print("=== 变体样例 ===")
    for k, v in list(var.items())[:12]:
        print(f"  {k} ({v})")

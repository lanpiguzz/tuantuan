# -*- coding: utf-8 -*-
"""超纲筛查：把用户两份词表 + 我的补遗，逐词对照剑桥官方 A2 Key / Key for Schools
词汇表（2025-08 版），找出「不在官方词表内」的词。

判定规则（保守，避免误伤）：
  IN      = 词（或其英美变体/并列写法/词形变化）出现在官方主词表
  VARIANT = 命中官方表的 Am Eng / Br Eng 变体记录（gray、center 等）
  MORPH   = 命中官方表中某词的派生形式（如 official -> officially）
  TOPIC   = 只出现在附录2 主题列表里（官方认可，但不在字母主词表）
  OUT     = 完全查不到 -> 超纲
"""
import re
import json
from extract_official import extract, normalize


# 官方表 Introduction 的「举例」行明确列出这些国家/国籍，属于考试认可词汇，
# 但不在字母主词表（仅出现在 note 里）。筛查时视为 IN。
NOTE_COUNTRIES = {
    "china", "chinese", "france", "french", "ireland", "irish", "india", "indian",
    "italy", "italian", "spain", "spanish", "america", "american", "australia",
    "australian", "britain", "british", "canada", "canadian", "germany", "german",
    "japan", "japanese", "england", "english", "europe", "european",
}

# ---------- 读取用户两份词表 ----------
def read_docx_words(path):
    from docx import Document
    d = Document(path)
    out = []
    for t in d.tables:
        for r in t.rows[1:]:
            w = r.cells[0].text.strip()
            if w:
                out.append(w)
    return out


def read_my_docx(path):
    from docx import Document
    d = Document(path)
    out = []
    for t in d.tables:
        for r in t.rows[1:]:
            w = r.cells[0].text.strip()
            if w:
                out.append(w)
    return out


def build_lemma_index(main):
    """建立「词干 -> 原词」索引，用于派生形式匹配。"""
    idx = {}
    for w in main:
        base = w.lower()
        idx.setdefault(base, w)
        for suf in ("s", "es", "ed", "ing", "er", "est", "ly", "d", "ment",
                    "ness", "tion", "sion", "ful", "less", "able", "ive"):
            if base.endswith(suf) and len(base) - len(suf) >= 3:
                stem = base[: -len(suf)]
                idx.setdefault(stem, w)
                # 双写辅音还原：running -> run
                if len(stem) >= 3 and stem[-1] == stem[-2] and stem[-1] not in "aeiou":
                    idx.setdefault(stem[:-1], w)
                # y->i 还原：happily -> happy
                if stem.endswith("i") and len(stem) >= 2:
                    idx.setdefault(stem[:-1] + "y", w)
    return idx


def main():
    main_words, ws, topics, variants = extract()
    main_norm = {normalize(w): w for w in main_words}
    variant_norm = {normalize(v): v for v in variants}
    topic_norm = set()
    for t, ws_ in topics.items():
        for w in ws_:
            topic_norm.add(normalize(w))
    lemma = build_lemma_index(main_words)

    # 官方表里的多词短语（空格分隔）
    multi = {w for w in main_norm if " " in w}

    def classify(word):
        w = normalize(word)
        if not w:
            return "EMPTY", ""
        if w in main_norm:
            return "IN", main_norm[w]
        if w in variant_norm:
            return "VARIANT", variant_norm[w]
        if w in topic_norm:
            return "TOPIC", ""
        # 国家/国籍：官方表 Introduction 明确举例，视为认可
        if w in NOTE_COUNTRIES:
            return "NOTE", "Introduction 举例"
        # 多词短语：任一实义词命中主词表即视为认可
        if " " in w or "-" in w:
            toks = re.split(r"[\s\-]+", w)
            for tok in toks:
                t = tok.strip()
                if len(t) >= 3 and (t in main_norm or t in variant_norm
                                    or t in NOTE_COUNTRIES):
                    return "IN", main_norm.get(t, t)
        # 空格变体：'t-shirt' vs 'T shirt'
        alt = w.replace("-", " ")
        if alt in main_norm:
            return "IN", main_norm[alt]
        alt2 = w.replace(" ", "-")
        if alt2 in main_norm:
            return "IN", main_norm[alt2]
        # 派生形式
        if w in lemma:
            return "MORPH", lemma[w]
        # 复数/三单
        if w.endswith("ies") and w[:-3] + "y" in lemma:
            return "MORPH", lemma[w[:-3] + "y"]
        if w.endswith("es") and w[:-2] in lemma:
            return "MORPH", lemma[w[:-2]]
        if w.endswith("s") and w[:-1] in lemma:
            return "MORPH", lemma[w[:-1]]
        return "OUT", ""

    # ---------- 三份来源 ----------
    srcs = {
        "原表1_精选324": read_docx_words(
            "C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表.docx"),
        "原表2_全量1080": read_docx_words(
            "C:/Users/Administrator/Documents/Trae-CN/temp_20261024_01/KET主题词汇表2.docx"),
        "我的补遗256": read_my_docx("KET词汇补遗（紧凑版）.docx"),
    }

    report = {}
    for name, words in srcs.items():
        buckets = {"IN": [], "VARIANT": [], "MORPH": [], "TOPIC": [], "NOTE": [], "OUT": []}
        for w in words:
            c, base = classify(w)
            buckets[c].append((w, base))
        report[name] = buckets
        total = len(words)
        print(f"=== {name}（{total} 词）===")
        for k in ("IN", "VARIANT", "MORPH", "TOPIC", "NOTE", "OUT"):
            print(f"  {k:8s} {len(buckets[k]):4d}  ({len(buckets[k])*100//total}%)")
        print()

    # ---------- 超纲明细 ----------
    all_out = {}
    for name, buckets in report.items():
        for w, base in buckets["OUT"]:
            all_out.setdefault(w, []).append(name)

    print("=" * 70)
    print(f"超纲词总表（任一来源出现即收录）: {len(all_out)} 个")
    print("=" * 70)
    for w in sorted(all_out, key=lambda x: x.lower()):
        print(f"  {w:28s} 出现于: {', '.join(all_out[w])}")

    json.dump(
        {"report": {k: {kk: vv for kk, vv in v.items()} for k, v in report.items()},
         "out": {k: v for k, v in all_out.items()}},
        open("out_of_scope.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    with open("out_of_scope.txt", "w", encoding="utf-8") as f:
        for w in sorted(all_out, key=lambda x: x.lower()):
            f.write(f"{w}\t{', '.join(all_out[w])}\n")
    print("\n已写出 out_of_scope.json / out_of_scope.txt")

    # 官方表统计
    print(f"\n官方主词表 {len(main_words)} 条 | 主题列表 {len(topic_norm)} 词 | "
          f"英美变体 {len(variants)} 个 | 多词短语 {len(multi)} 条")


if __name__ == "__main__":
    main()

"""把待審的品牌字串分群，讓人一次決定一群，而不是一列一列決定。

例如「台灣角川」「台灣角川股份有限公司」「角川」會被分到同一群 —— 看一次就能三筆一起處理。

分群方式刻意保守：寧可分太細（多看幾群）也不要把不同品牌黏在一起（會誤判）。
用共同的英文 token 或中文片段當 blocking key，再用 union-find 合併。
"""
from __future__ import annotations

import re
from collections import defaultdict

from rapidfuzz import fuzz

from .text import CJK_RE, STOP_TOKENS, clean_raw, normalize, split_parts

# 這些中文字太常見，單獨當分群依據會把不相干的品牌黏在一起
CJK_STOP = {"有限", "公司", "股份", "企業", "國際", "實業", "事業", "科技", "生技", "文化", "出版", "台灣",
            "中國", "工業", "貿易", "商行", "食品", "生活", "用品", "精品", "嚴選", "專賣", "旗艦"}


def _blocks(s: str) -> set:
    """一個品牌字串的 blocking key：拿來跟誰比較的依據。"""
    full, lat, cjk = split_parts(s)
    keys = set()
    for t in re.findall(r"[a-z0-9]+", lat or full):
        if len(t) >= 4 and t not in STOP_TOKENS:
            keys.add("L:" + t)
    # 中文取前 2 字與後 2 字（品牌名常以這兩端為主體）
    if len(cjk) >= 2:
        for piece in {cjk[:2], cjk[-2:]}:
            if piece not in CJK_STOP:
                keys.add("C:" + piece)
    return keys


class _UF:
    def __init__(self):
        self.p = {}

    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]
            x = self.p[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.p[rb] = ra


def group_keys(rows, min_ratio=70) -> dict:
    """rows: [(key, 顯示用的品牌字串)]。回傳 {key: 群組編號}，單獨一個的不給編號（空字串）。"""
    texts = {k: clean_raw(v) for k, v in rows if v}
    by_block = defaultdict(list)
    for k, v in texts.items():
        for b in _blocks(v):
            by_block[b].append(k)

    uf = _UF()
    for k in texts:
        uf.find(k)
    for b, members in by_block.items():
        if len(members) < 2 or len(members) > 60:      # 太大的 block 多半是泛用詞，跳過
            continue
        base = members[0]
        for m in members[1:]:
            # blocking 只是候選，還要字串真的夠像才合併
            if fuzz.token_set_ratio(normalize(texts[base]), normalize(texts[m])) >= min_ratio:
                uf.union(base, m)

    members = defaultdict(list)
    for k in texts:
        members[uf.find(k)].append(k)
    out, n = {}, 0
    for root, ms in sorted(members.items(), key=lambda kv: -len(kv[1])):
        if len(ms) < 2:
            continue
        n += 1
        for m in ms:
            out[m] = f"G{n:04d}"
    return out

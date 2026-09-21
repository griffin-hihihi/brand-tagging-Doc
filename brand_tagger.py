#!/usr/bin/env python3
"""
競網品牌 → shp 品牌庫 對應工具 v0.4

用法（在 brand_tagging 資料夾執行）：
    py brand_tagger.py --items momo.csv --site momo --category Beauty

每次執行：
  ① 自動讀回 output/ 裡這個站點＋品類的 Excel，把【人工】欄位的判斷寫進規則庫 brand_rules.db
  ② 套用規則庫＋自動比對
  ③ 產出一個新的 Excel：output/<站點>_<品類>_<日期_時間>.xlsx，舊檔自動移到 output/歷史/
"""
import argparse
import datetime as dt
import re
import shutil
import sqlite3
import sys
import time
import unicodedata
from collections import defaultdict
from pathlib import Path

try:
    import numpy as np
    import openpyxl  # noqa: F401  讀回 Excel 用
    import pandas as pd
    import xlsxwriter  # noqa: F401  輸出 Excel 用
    from rapidfuzz import fuzz, process
except ImportError as e:
    sys.exit(f"✖ 缺少套件 {e.name}，請先執行：\n"
             f"  py -m pip install pandas openpyxl xlsxwriter rapidfuzz opencc-python-reimplemented")
try:
    from opencc import OpenCC
    to_trad = OpenCC("s2t").convert
except Exception:
    to_trad = lambda s: s

# ================================================================================ 設定（可自行修改）
# 三種建議結果
TYPE_POOL = "品牌庫品牌"
TYPE_NEW = "建議品牌庫新增品牌"     # 有品牌，只是品牌庫沒有
TYPE_NB = "No brand"                # 白牌，沒有品牌
NB_ID = 0
SPECIAL_IDS = {0, 99999999}         # 品牌庫中不當候選的 id（NoBrand / nodata）

# No brand 原因（Excel 下拉選單也用這份）
R_EMPTY = "品牌欄與商品名稱皆無品牌"
R_DESC = "商品名稱【】為描述非品牌"
R_COMPAT = "相容/副廠配件"
R_WORD = "品牌欄標示無品牌"
R_STORE = "店家/賣場名稱非品牌"
R_OTHER = "其他(請備註)"
REASONS = [R_EMPTY, R_DESC, R_COMPAT, R_WORD, R_STORE, R_OTHER]

# 判斷路徑：代碼 → (名稱, 信心, 條件, 例子)
PATHS = {
    "A1": ("品牌庫完全相符", "高", "品牌名與品牌庫某品牌完全相同", "DHC、Simba 小獅王辛巴"),
    "A2": ("同名多筆→選業績最高", "高", "品牌庫有多個同名品牌（大小寫、空格、標點不同）", "LUX / Lux、GATSBY ×2"),
    "A3": ("相近名稱多筆→選業績最高", "中", "品牌庫有多個名稱很相近的品牌", "REACH 麗奇 / REACH 麗奇 II"),
    "A4": ("中英文名擇一相符", "高", "「英文 中文」雙語名稱，其中一半與品牌庫相符、另一半不衝突", "MUJI 無印良品 → Muji"),
    "A5": ("最小顆粒度→選最具體的", "中", "多個候選都出現在競品品牌中，選最具體（最長）的那個", "Apple iPhone → iPhone"),
    "A6": ("多個不同品牌同分", "低", "有兩個以上名稱不同的品牌同樣符合，無法自動決定", ""),
    "A7": ("自動學習的別名", "中", "之前在「英文 中文」品牌中高信心確認過的中文或英文名", "莉婕 → Liese（學自 Liese 莉婕）"),
    "B1": ("由商品名稱【】判斷", "中", "品牌欄空白或比對不到，改用標題【】比對", "品牌欄空白、標題【Simba】"),
    "B2": ("商品名稱內含品牌名", "低", "品牌欄空白，只在商品名稱中找到品牌名", "…Innisfree 綠茶精華 同款收納包"),
    "C1": ("部分相符", "低", "品牌名包含品牌庫的某個品牌名，但不完全相同", ""),
    "N1": ("建議新增（無相似品牌）", "中", "品牌欄有明確品牌名，品牌庫找不到相符或相似的品牌", "Addme 愛戴美"),
    "N2": ("建議新增（有相似候選）", "低", "只有相似但不夠像的候選，依最小顆粒度不採用", "L’OREAL 巴黎萊雅PRO（候選 L'Oreal Paris）"),
    "N3": ("疑似新品牌（只有【】）", "低", "品牌欄空白，【】看起來像品牌但品牌庫沒有", "【MEIDIAN】"),
    "Z1": ("No brand：品牌欄標示無品牌", "高", "品牌欄寫無品牌、其他、副廠、OEM…", ""),
    "Z2": ("No brand：相容/副廠配件", "中", "出現「適用iPhone、相容Switch、副廠」等字眼", "【適用iPhone】充電線"),
    "Z3": ("No brand：無任何品牌資訊", "中", "品牌欄空白，商品名稱也找不到品牌", ""),
    "Z4": ("No brand：【】為商品描述", "中", "品牌欄空白，【】是商品描述而非品牌", "【瓶罐收納】"),
    "R1": ("規則庫（人工判定過）", "高", "這個品牌字串之前人工審過", ""),
    "R2": ("單品人工判定", "高", "這個品號在「總結果」頁被人工指定", ""),
}
DOWN = {"高": "中", "中": "低", "低": "低"}

NOISE_WORDS = ["官方直營", "官方旗艦店", "旗艦店", "官方授權", "官方", "台灣總代理", "總代理", "即期品",
               "台灣公司貨", "公司貨", "原廠", "正品", "現貨", "免運"]
ORIGIN_PREFIX = ["紐西蘭", "日本", "韓國", "美國", "法國", "德國", "英國", "澳洲", "義大利", "台灣", "泰國"]
NOBRAND_WORDS = {"無品牌", "其他", "其它", "副廠", "oem", "nobrand", "none", "無", "通用", "自有品牌", "mit",
                 "白牌", "0"}
GENERIC_WORDS = ["收納", "居家", "生活", "買一送一", "限時", "特價", "熱銷", "新品", "任選", "組合", "入組",
                 "超值", "現貨", "免運", "團購", "禮盒", "必備", "推薦", "精選", "款"]
COMPAT_RE = re.compile(r"(副廠|通用款|(適用|相容|兼容|for)\s*(iphone|ipad|apple|samsung|switch|airpods|galaxy|"
                       r"macbook|pixel|蘋果|三星|小米))", re.I)
STOP_TOKENS = {"the", "body", "shop", "house", "home", "life", "care", "baby", "beauty", "paris", "pro",
               "professional", "plus", "health", "healthcare", "lab", "labs", "official", "store", "company",
               "international", "taiwan", "japan", "korea", "group", "entertainment", "choice", "goals",
               "signature", "coffee", "tech", "village", "face", "north", "balance", "good", "smile", "mobile",
               "national", "geographic", "digital", "western", "space", "power", "cook", "dream", "trend",
               "future", "biotech", "educational", "nutrition", "optimum", "nature", "natural", "organic"}

AUTO_TH = 95        # 視為「相符」的分數
SUGGEST_TH = 75     # 低於此分數的候選只列出、不建議（最小顆粒度：不往相似或上層品牌歸）
TIER_GAP = 5        # 決策樹第 ① 步：與最高分差距在此之內，才算「同樣相關」
NEAR_RATIO = 85     # 兩個候選名稱相似度 ≥ 此值，算「名稱相近」
TITLE_FACTOR = 0.95
FUZZY_CUTOFF = 88
SAMPLE_N = 30       # 高 / 中信心各抽查幾個品牌字串

# 使用者看到的欄位
H_PICK, H_REASON, H_NOTE = "【人工】判斷", "【人工】No brand原因", "【人工】備註"
HUMAN_COLS = [H_PICK, H_REASON, H_NOTE]
PICK_OPTIONS = ["v", "1", "2", "3", TYPE_NB, TYPE_NEW]
RESULT_COLS = ["shp brand name1", "shp brand name2", "shp brand name3", "suggest brand name", "suggest brand id",
               "新增品牌名稱", "No brand原因", "信心程度", "是否待人工判斷", "判斷路徑", "判斷說明"]
CJK_RE = re.compile(r"[\u3400-\u9fff\uf900-\ufaff]")
ACCEPT_WORDS = {"v", "ok", "y", "yes", "o", "✓", "✔", "對", "同意", "正確"}


# ================================================================================ 文字處理
def blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "None", "<NA>")


def clean_raw(s) -> str:
    s = "" if blank(s) else unicodedata.normalize("NFKC", str(s)).strip()
    for w in NOISE_WORDS:
        s = s.replace(w, " ")
    return re.sub(r"\s+", " ", s).strip()


def normalize(s) -> str:
    s = to_trad(unicodedata.normalize("NFKC", "" if blank(s) else str(s))).lower()
    s = s.replace("’", "").replace("'", "").replace("`", "")
    return "".join(ch for ch in s if ch.isalnum() or CJK_RE.match(ch))


def split_parts(s):
    n = normalize(s)
    return n, "".join(re.findall(r"[a-z0-9]+", n)), "".join(CJK_RE.findall(n))


def latin_tokens(s) -> set:
    s = to_trad(unicodedata.normalize("NFKC", str(s or ""))).lower().replace("’", "").replace("'", "")
    return {t for t in re.findall(r"[a-z0-9]+", s) if len(t) >= 4 and t not in STOP_TOKENS}


def is_bilingual(s: str) -> bool:
    """「英文 中文」或「中文 英文」中間有空格 → 正常雙語品牌名（MUK 潮嘜）；LG生活健康 則不是"""
    return bool(re.fullmatch(r"[^\u3400-\u9fff]*[A-Za-z0-9][^\u3400-\u9fff]*\s+[\u3400-\u9fff].*", s)
                or re.fullmatch(r"[\u3400-\u9fff][^A-Za-z]*\s+[A-Za-z].*", s))


def extract_bracket(title) -> str:
    m = re.search(r"【(.*?)】", "" if blank(title) else str(title))
    if not m:
        return ""
    b = clean_raw(m.group(1))
    for p in ORIGIN_PREFIX:
        if b.startswith(p) and len(b) > len(p):
            b = b[len(p):]
    return b.strip()


def is_generic(b: str) -> bool:
    return (not b) or any(w in b for w in GENERIC_WORDS) or len(normalize(b)) > 14


def long_enough(a: str) -> bool:
    cjk = len(CJK_RE.findall(a))
    return (cjk == 0 and len(a) >= 5) or cjk >= 3


def names_close(g1: str, g2: str) -> bool:
    """兩個品牌庫名稱是否「相同或很相近」"""
    if g1 == g2:
        return True
    if min(len(g1), len(g2)) >= 3 and (g1 in g2 or g2 in g1):
        return True
    return fuzz.ratio(g1, g2) >= NEAR_RATIO


def id_str(x):
    return "" if blank(x) else str(int(float(x)))


# ================================================================================ 品牌庫索引
class BrandIndex:
    def __init__(self, pool: pd.DataFrame, aliases: list, cat_col: str):
        self.name = dict(zip(pool.brand_id, pool.brand_name))
        self.adg = dict(zip(pool.brand_id, pool["adg"])) if "adg" in pool else {}
        self.cat = dict(zip(pool.brand_id, pool[cat_col])) if cat_col in pool else {}
        self.parts = {b: split_parts(n) for b, n in self.name.items()}
        self.group = {b: p[0] for b, p in self.parts.items()}
        self.by_norm = defaultdict(set)
        self.idx = defaultdict(set)
        self.natural = set()
        for b, (full, lat, cjk) in self.parts.items():
            self.by_norm[full].add(b)
            if b in SPECIAL_IDS:
                continue
            for a in (full, lat, cjk):
                if len(a) >= 2 or (a and a == full):
                    self.idx[a].add(b)
                    self.natural.add((a, b))
        self.human, self.auto = set(), set()
        for alias, b, src in aliases:
            if b in self.name and b not in SPECIAL_IDS:
                self.idx[alias].add(b)
                if src == "human":
                    self.human.add((alias, b))
                elif (alias, b) not in self.natural:
                    self.auto.add((alias, b))
        self.tok = defaultdict(set)
        for b, n in self.name.items():
            if b not in SPECIAL_IDS:
                for t in latin_tokens(n):
                    self.tok[t].add(b)
        self.long_keys = {a for a in self.idx if long_enough(a)}
        self._fz = {}
        self.fuzzy_by_len = defaultdict(list)
        for a in self.idx:
            if len(a) >= 5:
                self.fuzzy_by_len[(len(a), a[0])].append(a)

    def label(self, b):
        return f"{self.name.get(b, '?')} [{b}]"

    def adg_of(self, b):
        try:
            return float(self.adg.get(b, 0) or 0)
        except ValueError:
            return 0.0

    def fuzzy_choices(self, q):
        L, c = len(q), q[0]
        if (L, c) not in self._fz:
            lo, hi = int(L * 0.78), int(L / 0.78) + 1
            self._fz[(L, c)] = [a for n in range(lo, hi + 1) for a in self.fuzzy_by_len.get((n, c), [])]
        return self._fz[(L, c)]

    def match(self, s: str) -> dict:
        """回傳 {brand_id: [分數, 說明, 命中長度, 人工別名?, 命中方式]}"""
        out = {}
        full, lat, cjk = split_parts(s)
        if not full:
            return out
        bil = is_bilingual(clean_raw(s))

        def add(bids, score, desc, alias, via):
            for b in bids:
                h = (alias, b) in self.human
                v = via
                if h and score >= AUTO_TH:
                    desc2 = "人工確認過的別名"
                elif (alias, b) in self.auto and score >= AUTO_TH:
                    desc2, v = f"自動學習的別名「{alias}」", "auto"
                else:
                    desc2 = desc
                c = [score, desc2, len(alias), h, v]
                if b not in out or (c[0], c[3], c[2]) > (out[b][0], out[b][3], out[b][2]):
                    out[b] = c

        if full in self.idx:
            exact = {b for b in self.idx[full] if self.group[b] == full or (full, b) in self.human}
            add(exact, 100, "品牌名完全相符", full, "full")
            add(self.idx[full] - exact, 95, "與品牌庫品牌的" + ("中文名" if CJK_RE.search(full) else "英文名") + "相符",
                full, "cjk" if CJK_RE.search(full) else "lat")
        hl = self.idx.get(lat, set()) if len(lat) >= 2 and lat != full else set()
        hc = self.idx.get(cjk, set()) if len(cjk) >= 2 and cjk != full else set()
        both = hl & hc
        add(both, 100, "中英文名皆相符", full, "full")
        weak_l = len(lat) <= 3 and bool(cjk) and not bil
        weak_c = len(cjk) <= 2 and len(lat) >= 4 and not bil
        add(hl - both, 70 if weak_l else 95,
            "英文縮寫過短且黏著中文（可能是集團/公司名）" if weak_l else "英文名相符", lat, "lat")
        add(hc - both, 70 if weak_c else 95, "中文名過短（可能是泛用詞）" if weak_c else "中文名相符", cjk, "cjk")
        for i in range(len(full)):
            for j in range(i + 3, min(len(full), i + 25) + 1):
                sub = full[i:j]
                if sub != full and sub in self.long_keys:
                    add(self.idx[sub], 75, f"品牌名包含「{sub}」", sub, "sub")
        # 空格隔開的單字剛好等於某品牌完整名稱（例：Kanebo KATE 的 KATE、ASUS ROG 的 ROG）
        words = [normalize(w) for w in re.split(r"[\s/|｜]+", clean_raw(s))]
        if len(words) > 1:
            for w_ in words:
                if w_ and w_ != full and w_ not in STOP_TOKENS and (len(w_) >= 3 or CJK_RE.search(w_) and len(w_) >= 2):
                    add({b for b in self.idx.get(w_, set()) if self.group[b] == w_}, 85,
                        f"品牌名中的「{w_}」與品牌庫品牌相同", w_, "word")
        for t in latin_tokens(s):
            add(self.tok.get(t, set()), 65, f"英文單字「{t}」相同", t, "tok")
        if not any(v[0] >= AUTO_TH for v in out.values()):
            for q in {full, lat}:
                if len(q) >= 5:
                    for a, sc, _ in process.extract(q, self.fuzzy_choices(q), scorer=fuzz.ratio, limit=3,
                                                    score_cutoff=FUZZY_CUTOFF):
                        if a != q:
                            add(self.idx[a], round(60 + (sc - FUZZY_CUTOFF) / (100 - FUZZY_CUTOFF) * 14),
                                f"拼字相似({sc:.0f}%)", a, "fuzzy")
        # 中英文互相矛盾 → 大幅降分（Dove 多芬 ≠ Dove 德芙；CLEAR 淨 ≠ CLEAR 可麗兒）
        for b, v in out.items():
            if v[4] in ("full", "auto") or v[3]:
                continue
            _, pl, pc = self.parts[b]
            cjk_conf = bool(cjk) and bool(pc) and cjk not in pc and pc not in cjk
            lat_conf = len(lat) >= 2 and len(pl) >= 2 and lat not in pl and pl not in lat
            if cjk_conf and v[4] != "cjk":
                v[0], v[1] = min(v[0], 60), f"{v[1]}，但中文名不同（{pc}≠{cjk}）"
            elif lat_conf and v[4] == "cjk":
                v[0], v[1] = min(v[0], 60), f"{v[1]}，但英文名不同（{pl}≠{lat}）"
        return out

    def scan_title(self, title) -> dict:
        out = {}
        t = normalize(title)[:80]
        for i in range(len(t)):
            for j in range(i + 3, min(len(t), i + 20) + 1):
                sub = t[i:j]
                if sub in self.long_keys:
                    for b in self.idx[sub]:
                        if b not in out or out[b][2] < len(sub):
                            out[b] = [75, f"商品名稱內含「{sub}」", len(sub), (sub, b) in self.human, "scan"]
        return out

    def resolve(self, text):
        t = str(text).strip()
        m = re.search(r"\[(\d+)\]\s*$", t)
        if m:
            t = m.group(1)
        if re.fullmatch(r"\d+(\.0)?", t):
            b = int(float(t))
            return (b, None) if b in self.name else (None, f"brand_id {b} 不在品牌庫")
        ids = self.by_norm.get(normalize(t), set())
        if len(ids) == 1:
            return next(iter(ids)), None
        if len(ids) > 1:
            return None, f"「{t}」在品牌庫有多筆 ({', '.join(map(str, sorted(ids)))})，請改填 brand_id"
        return None, f"品牌庫找不到「{t}」（品牌庫沒有的品牌請選：{TYPE_NEW}）"


# ================================================================================ 決策樹
def choose(strong, comp_full, bi):
    """
    多個候選時：① 相關度 → ② 最小顆粒度 → ③ 業績(adg)
    回傳 (brand_id, 路徑代碼或 None, 補充說明)
    """
    top_id, top = strong[0]
    if top[3]:                                                        # 人工確認過的別名直接採用
        return top_id, None, ""
    # ① 相關度：只留下和最高分同一層的候選
    tier = [(b, v) for b, v in strong if v[0] >= top[0] - TIER_GAP and not v[3]]
    tg = bi.group[top_id]
    near = [(b, v) for b, v in tier if names_close(bi.group[b], tg)]
    far_tie = [b for b, v in tier if v[0] == top[0] and (b, v) not in near]
    if far_tie:
        # ② 最小顆粒度：同分的不同品牌如果都出現在競品品牌中（例：Apple iPhone、Kanebo KATE），
        #    子品牌通常寫在母品牌後面 → 選位置最後的；互相包含時選較長的
        tied = [top_id] + far_tie
        pos = {b: comp_full.find(bi.group[b]) for b in tied}
        if all(p >= 0 for p in pos.values()):
            pick = max(tied, key=lambda b: (pos[b] + len(bi.group[b]), len(bi.group[b]), bi.adg_of(b)))
            others = "、".join(dict.fromkeys(bi.name[b] for b in tied if bi.name[b] != bi.name[pick]))
            return pick, "A5", f"；{others} 也出現在競品品牌中，依最小顆粒度選較具體的 {bi.name[pick]}"
        names = "、".join(bi.name[b] for b in tied[:3])
        return top_id, "A6", f"；{names} 同分但名稱不同"
    if len(near) == 1:
        return top_id, None, ""
    # ② 最小顆粒度：名稱完整出現在競品品牌中的候選，取最具體（最長）的
    supported = [(b, v) for b, v in near if bi.group[b] and bi.group[b] in comp_full]
    cands, gran = near, False
    if supported:
        L = max(len(bi.group[b]) for b, _ in supported)
        cands = [(b, v) for b, v in supported if len(bi.group[b]) == L]
        gran = any(len(bi.group[b]) < L for b, _ in supported)
    # ③ 業績：名稱相同或相近 → adg 最高
    pick = max(cands, key=lambda kv: bi.adg_of(kv[0]))[0]
    if len(cands) == 1:
        if gran:
            return pick, "A5", "；多個候選都出現在競品品牌中，選最具體的"
        return pick, None, "；其他相近候選名稱多了競品沒寫的字，不採用"
    same = len({bi.group[b] for b, _ in cands}) == 1
    return pick, "A2" if same else "A3", f"；品牌庫有 {len(cands)} 個{'同名' if same else '相近'}品牌，選業績最高"


def rank(cands, bi):
    return sorted(cands.items(), key=lambda kv: (-kv[1][0], -kv[1][3], -kv[1][2], -bi.adg_of(kv[0])))


def result(bi, typ, path, explain, ranked=(), bid=None, new_name="", reason="", conf=None):
    labels = [bi.label(b) for b, _ in list(ranked)[:3]] + ["", "", ""]
    conf = conf or PATHS[path][1]
    name = bi.name[bid] if typ == TYPE_POOL else typ
    return {"shp brand name1": labels[0], "shp brand name2": labels[1], "shp brand name3": labels[2],
            "suggest brand name": name,
            "suggest brand id": str(bid) if typ == TYPE_POOL else ("0" if typ == TYPE_NB else ""),
            "新增品牌名稱": new_name if typ == TYPE_NEW else "", "No brand原因": reason if typ == TYPE_NB else "",
            "信心程度": conf, "是否待人工判斷": "是" if conf == "低" or typ == TYPE_NEW else "否",
            "判斷路徑": f"{path} {PATHS[path][0]}", "判斷說明": explain, "_type": typ, "_bid": bid}


def tag(raw_brand, title, cat, bi: BrandIndex, cat_check: bool) -> dict:
    rb = clean_raw(raw_brand)
    t = "" if blank(title) else str(title)
    bracket = extract_bracket(t)

    if rb and normalize(rb) in NOBRAND_WORDS:
        return result(bi, TYPE_NB, "Z1", "品牌欄寫明無品牌", reason=R_WORD)
    if COMPAT_RE.search(rb) or (not rb and COMPAT_RE.search(t)):
        return result(bi, TYPE_NB, "Z2", "偵測到相容/副廠字眼", reason=R_COMPAT)

    via, flags, cands, comp = "brand", [], {}, normalize(rb)
    if rb:
        cands = bi.match(rb)
        if bracket and normalize(bracket) != comp:
            tb = bi.match(bracket)
            sb = {b for b, v in cands.items() if v[0] >= SUGGEST_TH}
            st = {b for b, v in tb.items() if v[0] >= AUTO_TH}
            if sb and st and not (sb & st):
                flags.append("品牌欄與商品名稱【】指向不同品牌")
            elif st and not sb:
                cands, via, comp = tb, "bracket", normalize(bracket)
    else:
        if bracket and not is_generic(bracket):
            cands, via, comp = bi.match(bracket), "bracket", normalize(bracket)
        if not any(v[0] >= SUGGEST_TH for v in cands.values()):
            sc = bi.scan_title(t)
            if sc:
                cands, via, comp = sc, "scan", normalize(t)
    if via == "bracket":
        for v in cands.values():
            v[0], v[1] = round(v[0] * TITLE_FACTOR), "商品名稱【】" + v[1]

    ranked = rank(cands, bi)
    strong = [kv for kv in ranked if kv[1][0] >= SUGGEST_TH]

    if not strong:              # ---- 沒有夠像的候選 → 建議新增 或 No brand
        weak = f"只有相似候選（{ranked[0][1][1]}），依最小顆粒度不採用" if ranked else ""
        if rb:
            return result(bi, TYPE_NEW, "N2" if ranked else "N1",
                          weak or "品牌欄有品牌名，品牌庫找不到相符或相似品牌", ranked, new_name=rb)
        if bracket and not is_generic(bracket):
            return result(bi, TYPE_NEW, "N3", weak or "品牌欄空白，【】像品牌但品牌庫沒有", ranked, new_name=bracket)
        if bracket:
            return result(bi, TYPE_NB, "Z4", "品牌欄空白，【】內容為商品描述", ranked, reason=R_DESC)
        return result(bi, TYPE_NB, "Z3", "品牌欄空白，商品名稱也找不到品牌", ranked, reason=R_EMPTY)

    # ---- 有候選 → 決策樹
    best_id, tree_path, extra = choose(strong, comp, bi)
    best = dict(strong)[best_id]
    if via == "scan":
        path = "B2"
    elif via == "bracket":
        path = "B1"
    elif tree_path:
        path = tree_path
    elif best[4] == "auto":
        path = "A7"
    elif best[0] >= AUTO_TH:
        path = "A1" if best[4] == "full" or best[3] else "A4"
    else:
        path = "C1"
    explain = best[1] + extra
    conf = PATHS[path][1]
    if flags:
        conf = "低"
        explain += "；" + "；".join(flags)
    if cat_check and cat and bi.cat.get(best_id) and bi.cat[best_id] != cat:
        conf = DOWN[conf]
        explain += f"；品牌類目({bi.cat[best_id]})與商品類目({cat})不同"
    # 候選排序：選中的放第一個
    ranked = [(best_id, best)] + [kv for kv in ranked if kv[0] != best_id]
    return result(bi, TYPE_POOL, path, explain, ranked, bid=best_id, conf=conf)


# ================================================================================ 規則庫
SCHEMA = """
CREATE TABLE IF NOT EXISTS brand_rules(
  key TEXT PRIMARY KEY, decision_type TEXT, brand_id INTEGER, brand_name TEXT, nobrand_reason TEXT,
  note TEXT, raw_example TEXT, site TEXT, category TEXT, source TEXT, updated_at TEXT);
CREATE TABLE IF NOT EXISTS item_overrides(
  site TEXT, item_id TEXT, decision_type TEXT, brand_id INTEGER, brand_name TEXT, nobrand_reason TEXT,
  note TEXT, source TEXT, updated_at TEXT, PRIMARY KEY(site, item_id));
CREATE TABLE IF NOT EXISTS brand_aliases(
  alias TEXT, brand_id INTEGER, source TEXT, created_at TEXT, PRIMARY KEY(alias, brand_id));
CREATE TABLE IF NOT EXISTS review_log(
  reviewed_at TEXT, source TEXT, site TEXT, category TEXT, key TEXT, raw_brand TEXT, path TEXT, conf TEXT,
  sampled INTEGER, sku INTEGER, sugg_type TEXT, sugg_id INTEGER, sugg_name TEXT,
  human_type TEXT, human_id INTEGER, human_name TEXT, agree INTEGER);
CREATE TABLE IF NOT EXISTS imported_files(source TEXT PRIMARY KEY, imported_at TEXT);
"""
OLD_TYPES = {"品牌池品牌": TYPE_POOL, "品牌池無此品牌": TYPE_NEW, "NoBrand": TYPE_NB}


def now():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def open_db(path):
    con = sqlite3.connect(path)
    cols = [r[1] for r in con.execute("PRAGMA table_info(brand_rules)")]
    if cols and "decision_type" not in cols:            # v0.2 → 新格式
        con.execute("ALTER TABLE brand_rules RENAME TO brand_rules_v02")
        if con.execute("SELECT name FROM sqlite_master WHERE name='item_overrides'").fetchone():
            con.execute("ALTER TABLE item_overrides RENAME TO item_overrides_v02")
        con.executescript(SCHEMA)
        con.execute(f"""INSERT INTO brand_rules SELECT key, CASE WHEN brand_id=0 THEN '{TYPE_NB}' ELSE '{TYPE_POOL}' END,
                        brand_id, brand_name, nobrand_reason, note, raw_example, site, '', source_file, updated_at
                        FROM brand_rules_v02""")
    con.executescript(SCHEMA)
    for old, new in OLD_TYPES.items():                  # v0.3 → v0.4 類型名稱
        con.execute("UPDATE brand_rules SET decision_type=? WHERE decision_type=?", (new, old))
        con.execute("UPDATE item_overrides SET decision_type=? WHERE decision_type=?", (new, old))
        con.execute("UPDATE brand_rules SET brand_name=? WHERE brand_name='NoBrand'", (TYPE_NB,))
    con.commit()
    return con


def row_type(row):
    name = "" if blank(row.get("suggest brand name")) else str(row["suggest brand name"])
    return name if name in (TYPE_NB, TYPE_NEW) else TYPE_POOL


def parse_human(val, reason, note, row, bi):
    """人工輸入 → (類型, brand_id, 名稱, 原因, 備註) 或 錯誤訊息；沒填 → None"""
    if blank(val) and blank(reason):
        return None
    note = "" if blank(note) else str(note).strip()
    reason = "" if blank(reason) else str(reason).strip()
    v = "" if blank(val) else str(val).strip()
    if v.lower() in ACCEPT_WORDS:                                  # v = 同意建議
        typ = row_type(row)
        if typ == TYPE_POOL:
            return TYPE_POOL, int(float(row["suggest brand id"])), row["suggest brand name"], "", note
        if typ == TYPE_NEW:
            return TYPE_NEW, None, str(row.get("新增品牌名稱") or ""), "", note
        return TYPE_NB, NB_ID, TYPE_NB, reason or str(row.get("No brand原因") or ""), note
    if re.fullmatch(r"[123](\.0)?", v):                            # 1/2/3 = 選第幾個候選
        cand = row.get(f"shp brand name{int(float(v))}")
        if blank(cand):
            return f"沒有第 {int(float(v))} 個候選"
        v = str(cand)
    if not v or normalize(v) in NOBRAND_WORDS:
        return TYPE_NB, NB_ID, TYPE_NB, reason or "人工判定無品牌", note
    if v.startswith(("建議品牌庫新增", "新增", "新品牌")):
        name = re.split(r"[:：]", v, 1)[1].strip() if re.search(r"[:：]", v) else ""
        if not name:
            for c in ("新增品牌名稱", "品牌欄", "商品名稱【】"):
                if not blank(row.get(c)):
                    name = clean_raw(row[c])
                    break
        return (TYPE_NEW, None, name, "", note) if name else "請寫成「新增:品牌名」"
    bid, err = bi.resolve(v)
    return err if err else (TYPE_POOL, bid, bi.name[bid], "", note)


def learn(files, con, bi, site, category, col_id, col_brand):
    """讀回 Excel 的【人工】欄位 → 規則庫。回傳 (筆數, 同意數, 錯誤)"""
    total = agree = 0
    errors = []
    for f in files:
        src = f"{f.name}@{dt.datetime.fromtimestamp(f.stat().st_mtime):%Y%m%d%H%M%S}"
        if con.execute("SELECT 1 FROM imported_files WHERE source=?", (src,)).fetchone():
            continue
        try:
            rv = pd.read_excel(f, sheet_name="審核", dtype=str)
            it = pd.read_excel(f, sheet_name="總結果", dtype=str,
                               usecols=lambda c: c in {col_id, col_brand, *HUMAN_COLS, "suggest brand name",
                                                       "suggest brand id", "新增品牌名稱", "No brand原因"})
        except PermissionError:
            sys.exit(f"✖ 讀不到 {f.name}，請先關閉 Excel 再執行")
        except ValueError:
            continue                                               # 舊版格式的檔案，略過
        it = it.rename(columns={col_brand: "品牌欄"})
        if H_PICK not in it:
            it[H_PICK], it[H_REASON], it[H_NOTE] = "", "", ""
        for _, row in rv.iterrows():
            d = parse_human(row.get(H_PICK), row.get(H_REASON), row.get(H_NOTE), row, bi)
            if d is None:
                continue
            label = row.get("品牌欄") if not blank(row.get("品牌欄")) else row.get("範例商品名稱")
            if isinstance(d, str):
                errors.append(f"審核頁｜{label}：{d}")
                continue
            key = row["key"]
            raw = "" if blank(row.get("品牌欄")) else str(row["品牌欄"])
            raw = raw or ("" if blank(row.get("商品名稱【】")) else str(row["商品名稱【】"]))
            if key.startswith("I:"):
                con.execute("INSERT OR REPLACE INTO item_overrides VALUES(?,?,?,?,?,?,?,?,?)",
                            (site, key[2:], *d, src, now()))
            else:
                con.execute("INSERT OR REPLACE INTO brand_rules VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                            (key, *d, raw, site, category, src, now()))
                if d[0] == TYPE_POOL and raw:        # 人工確認的中/英文名存成別名
                    full, lat, cjk = split_parts(clean_raw(raw))
                    for a in {full, lat if len(lat) >= 4 else "", cjk if len(cjk) >= 2 else ""} - {""}:
                        con.execute("INSERT OR REPLACE INTO brand_aliases VALUES(?,?,?,?)", (a, d[1], "human", now()))
            styp = row_type(row)
            sid = None if blank(row.get("suggest brand id")) else int(float(row["suggest brand id"]))
            ok = int(d[0] == styp and (d[0] != TYPE_POOL or d[1] == sid))
            con.execute("INSERT INTO review_log VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (now(), src, site, category, key, raw, str(row["判斷路徑"])[:2], row["信心程度"],
                         int(row.get("抽查") == "★"), int(float(row["商品數"])), styp, sid,
                         row["suggest brand name"], d[0], d[1], d[2], ok))
            total += 1
            agree += ok
        for _, row in it[~(it[H_PICK].map(blank) & it[H_REASON].map(blank))].iterrows():
            d = parse_human(row.get(H_PICK), row.get(H_REASON), row.get(H_NOTE), row, bi)
            if isinstance(d, str):
                errors.append(f"總結果頁｜品號 {row.get(col_id)}：{d}")
            elif d:
                con.execute("INSERT OR REPLACE INTO item_overrides VALUES(?,?,?,?,?,?,?,?,?)",
                            (site, str(row[col_id]).strip(), *d, src, now()))
                total += 1
        con.execute("INSERT INTO imported_files VALUES(?,?)", (src, now()))
        con.commit()
    return total, agree, errors


def learn_auto_aliases(kres, kinfo, bi, con):
    """「英文 中文」高信心相符 → 另一半自動存成別名（例：Liese 莉婕 → 之後只寫「莉婕」也能對到 Liese）"""
    n = 0
    ok = kres["判斷路徑"].str[:2].isin(["A1", "A2", "A4"]) & kres["信心程度"].eq("高") & kres["_type"].eq(TYPE_POOL)
    for k in kres.index[ok]:
        raw = clean_raw(kinfo.at[k, "raw"])
        if not is_bilingual(raw):
            continue
        bid = int(kres.at[k, "_bid"])
        _, lat, cjk = split_parts(raw)
        for a in (lat if len(lat) >= 4 else "", cjk if len(cjk) >= 2 else ""):
            if a and a not in bi.idx:                                  # 品牌庫還沒有這個名稱才學
                if con.execute("INSERT OR IGNORE INTO brand_aliases VALUES(?,?,?,?)",
                               (a, bid, "auto", now())).rowcount:
                    n += 1
    con.commit()
    return n


# ================================================================================ 讀檔
def read_table(path, must_have, table=None):
    p = Path(path)
    if not p.exists():
        sys.exit(f"✖ 找不到檔案 {p}，請確認檔案放在 {Path.cwd()}")
    if p.suffix.lower() in (".db", ".sqlite", ".sqlite3"):
        if not table:
            sys.exit(f"✖ {p.name} 是 SQLite，請用 --table / --pool-table 指定資料表")
        df = pd.read_sql(f'SELECT * FROM "{table}"', sqlite3.connect(p))
        df = df.astype(object).where(df.notna(), "").astype(str)
        miss = [c for c in must_have if c not in df]
        if miss:
            sys.exit(f"✖ 資料表 {table} 缺少欄位 {miss}")
        return df
    if p.suffix.lower() in (".xlsx", ".xls"):
        raw = pd.read_excel(p, header=None, dtype=str)
    else:
        raw = None
        for enc in ("utf-8-sig", "cp950"):
            try:
                raw = pd.read_csv(p, header=None, dtype=str, encoding=enc, keep_default_na=False,
                                  sep="\t" if p.suffix.lower() == ".tsv" else ",")
                break
            except UnicodeDecodeError:
                continue
    for i in range(min(10, len(raw))):
        row = [str(x).strip() for x in raw.iloc[i].tolist()]
        if all(c in row for c in must_have):
            df = raw.iloc[i + 1:].copy()
            df.columns = row
            return df.reset_index(drop=True).fillna("")
    sys.exit(f"✖ {p.name} 前 10 列找不到欄位 {must_have}")


def load_pool(a):
    pool = read_table(a.pool, ["brand_id", "brand_name"], a.pool_table)
    pool = pool[pd.to_numeric(pool["brand_id"], errors="coerce").notna()].copy()
    pool["brand_id"] = pool["brand_id"].astype(float).astype(int)
    pool["brand_name"] = pool["brand_name"].astype(str).str.strip()
    if "adg" in pool:
        pool["adg"] = pd.to_numeric(pool["adg"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
    return pool.drop_duplicates("brand_id")


# ================================================================================ 主流程
def run(a):
    t0 = time.time()
    base = Path.cwd()
    outdir, histdir = base / "output", base / "output" / "歷史"
    histdir.mkdir(parents=True, exist_ok=True)
    site = a.site or Path(a.items).stem
    cat_label = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", a.category).strip("_") if a.category else "全品類"
    prefix = f"{site}_{cat_label}"
    pat = re.compile(rf"^{re.escape(prefix)}_\d{{8}}_\d{{4,6}}\.xlsx$")
    print(f"━━━ 品牌 Tagging：{site} / {a.category or '全品類'} ━━━")

    # ① 讀回人工標記
    olds = sorted([f for f in outdir.glob("*.xlsx") if pat.match(f.name)], key=lambda f: f.stat().st_mtime)
    locked = [f.name for f in olds if (outdir / f"~${f.name}").exists()]
    if locked:
        sys.exit(f"✖ {locked[0]} 還開著。請在 Excel 存檔並關閉後再執行。")
    con = open_db(base / a.db)
    pool = load_pool(a)
    aliases = lambda: con.execute("SELECT alias, brand_id, source FROM brand_aliases").fetchall()
    n, agree, errors = learn(olds, con, BrandIndex(pool, aliases(), a.pool_col_cat), site, a.category or "", a.col_id, a.col_brand)
    if n:
        print(f"① 寫入規則庫：{n} 筆人工判斷（同意建議 {agree}、修改 {n - agree}）")
    else:
        print("① 沒有新的人工判斷" + ("" if olds else "（第一次執行）"))
    for e in errors[:30]:
        print(f"   ⚠ {e}")
    if errors:
        print(f"   ⚠ 以上 {len(errors)} 筆沒寫入，請在新檔案重新填")
    if a.learn_only:
        return

    # ② 讀資料
    bi = BrandIndex(pool, aliases(), a.pool_col_cat)
    items = read_table(a.items, [a.col_brand, a.col_title, a.col_id], a.table)
    if a.category:
        if a.col_cat not in items:
            sys.exit(f"✖ 商品資料沒有「{a.col_cat}」欄位，無法用 --category 篩選（可用 --col-cat 指定其他欄位）")
        vals = items[a.col_cat].astype(str).str.strip()
        items = items[vals == a.category]
        if items.empty:
            sys.exit(f"✖ {a.col_cat} 沒有「{a.category}」這個值。現有的值：{', '.join(sorted(vals.unique())[:15])}")
    items = items[~items[a.col_id].map(blank)].reset_index(drop=True)
    has_cat = a.col_cat in items
    cat_check = has_cat and bool(bi.cat) and bool(set(items[a.col_cat].unique()) & set(bi.cat.values()))
    print(f"② 品牌庫 {len(pool):,} 個品牌｜商品 {len(items):,} 筆"
          + ("" if cat_check or not has_cat else "｜（商品類目與品牌庫類目名稱對不上，略過類目檢查）"))

    # ③ 判斷
    brand_norm = items[a.col_brand].map({b: normalize(clean_raw(b)) for b in items[a.col_brand].unique()})
    brackets = items[a.col_title].map(extract_bracket)
    br_norm = brackets.map(lambda b: normalize(b) if b and not is_generic(b) else "")
    items["key"] = ("B:" + brand_norm).where(brand_norm != "",
                                             ("T:" + br_norm).where(br_norm != "", "I:" + items[a.col_id].astype(str)))
    g = items.groupby("key", sort=False)
    kinfo = g.agg(raw=(a.col_brand, "first"), title=(a.col_title, "first"), sku=(a.col_id, "size"))
    kinfo["bracket"] = brackets.groupby(items["key"], sort=False).first()
    kinfo["group"] = g[a.col_group].first() if a.col_group in items else ""
    if has_cat:
        cc = items.groupby(["key", a.col_cat]).size().reset_index(name="n").sort_values("n", ascending=False)
        kinfo["cat"] = cc.drop_duplicates("key").set_index("key")[a.col_cat]
    else:
        kinfo["cat"] = ""
    rules = {r[0]: r[1:] for r in con.execute(
        "SELECT key, decision_type, brand_id, brand_name, nobrand_reason FROM brand_rules")}
    rows, nk = [], len(kinfo)
    for i, (k, r) in enumerate(kinfo.iterrows(), 1):
        res = tag(r["raw"], r["title"], r["cat"] if has_cat else "", bi, cat_check)
        if k in rules:
            apply_rule(res, *rules[k], "R1", bi)
        rows.append(res)
        if i % 10000 == 0:
            print(f"   …{i:,}/{nk:,}")
    kres = pd.DataFrame(rows, index=kinfo.index)
    n_alias = learn_auto_aliases(kres, kinfo, bi, con)
    print(f"③ 判斷完成：{nk:,} 個品牌字串" + (f"｜自動學到 {n_alias} 個新別名" if n_alias else ""))

    # 展開到每個商品 + 單品人工判定
    full = items.drop(columns="key").join(kres, on=items["key"])
    full["key"] = items["key"]
    ovr = {r[0]: r[1:] for r in con.execute(
        "SELECT item_id, decision_type, brand_id, brand_name, nobrand_reason FROM item_overrides WHERE site=?", (site,))}
    hit = full[a.col_id].astype(str).str.strip().isin(ovr)
    for i in full.index[hit]:
        res = full.loc[i, RESULT_COLS + ["_type", "_bid"]].to_dict()
        apply_rule(res, *ovr[str(full.at[i, a.col_id]).strip()], "R2", bi)
        for c, v in res.items():
            full.at[i, c] = v

    # ④ 輸出
    review = build_review(kinfo, kres, a)
    stats = build_stats(full, review, a, con)
    newb = review[review["suggest brand name"] == TYPE_NEW][
        ["新增品牌名稱", "品牌欄", "商品數", "信心程度", "判斷路徑", "shp brand name1", "範例商品名稱"]
    ].rename(columns={"shp brand name1": "最相似的品牌庫品牌"}).sort_values("商品數", ascending=False)
    out = outdir / f"{prefix}_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
    if out.exists():
        out = outdir / f"{prefix}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    full_out = full.drop(columns=["_type", "_bid", "key"])
    for c in HUMAN_COLS:
        full_out[c] = ""
    full_out["key"] = full["key"]
    write_excel(out, review, full_out, newb, stats, a)
    moved = 0
    for f in olds:
        try:
            shutil.move(str(f), str(histdir / f.name))
            moved += 1
        except (PermissionError, OSError):
            pass
    if not a.no_db_results:
        save_results(con, full, site, a)

    pend = review["是否待人工判斷"].eq("是")
    print(f"④ 輸出完成（{time.time() - t0:.0f} 秒）" + (f"，舊檔 {moved} 個已移到 output/歷史" if moved else ""))
    print("━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    print(f"檔案：output/{out.name}")
    vc = full["suggest brand name"].map(lambda x: x if x in (TYPE_NEW, TYPE_NB) else TYPE_POOL).value_counts()
    print(f"結果：{TYPE_POOL} {vc.get(TYPE_POOL, 0):,}｜{TYPE_NEW} {vc.get(TYPE_NEW, 0):,}｜{TYPE_NB} "
          f"{vc.get(TYPE_NB, 0):,}（商品數）")
    print(f"待你審核：{pend.sum():,} 個品牌字串（涵蓋 {review.loc[pend, '商品數'].sum():,} 筆商品）"
          f"＋ 抽查★ {review['抽查'].eq('★').sum()} 個")
    print("下一步：打開檔案 →「審核」頁 → 填【人工】判斷 → 存檔、關閉 → 再執行一次同樣的指令")


def apply_rule(res, typ, bid, name, reason, code, bi):
    typ = OLD_TYPES.get(typ, typ)
    if typ == TYPE_NEW and res["_type"] == TYPE_POOL and res["判斷路徑"][:2] in ("A1", "A2", "A4") and code == "R1":
        res["判斷說明"] += "；（之前判為建議新增，但品牌庫已有此品牌）"
        res["信心程度"], res["是否待人工判斷"] = "中", "否"
        return
    res.update({"suggest brand name": bi.name.get(int(bid), name) if typ == TYPE_POOL else typ,
                "suggest brand id": id_str(bid) if typ == TYPE_POOL else ("0" if typ == TYPE_NB else ""),
                "新增品牌名稱": name if typ == TYPE_NEW else "", "No brand原因": (reason or "") if typ == TYPE_NB else "",
                "信心程度": "高", "是否待人工判斷": "否", "判斷路徑": f"{code} {PATHS[code][0]}",
                "判斷說明": PATHS[code][0], "_type": typ, "_bid": bid})


def build_review(kinfo, kres, a):
    rv = kinfo.join(kres).reset_index().rename(columns={
        "raw": "品牌欄", "bracket": "商品名稱【】", "group": "館別/集團(參考)", "cat": "主要類目",
        "sku": "商品數", "title": "範例商品名稱", "index": "key"})
    rv.insert(1, "抽查", "")
    cand = rv[rv["信心程度"].isin(["高", "中"]) & rv["是否待人工判斷"].eq("否") & ~rv["判斷路徑"].str.startswith("R")]
    for c in ("高", "中"):
        sub = cand[cand["信心程度"] == c]
        if len(sub) <= SAMPLE_N:
            pick = sub.index
        else:   # 依商品數加權抽樣（開根號避免大品牌壟斷樣本）
            w = sub["商品數"].astype(float) ** 0.5
            pick = np.random.default_rng(42).choice(sub.index, size=SAMPLE_N, replace=False, p=(w / w.sum()).values)
        rv.loc[pick, "抽查"] = "★"
    rv["_o"] = np.where(rv["是否待人工判斷"].eq("是"), 0, np.where(rv["抽查"].eq("★"), 1, 2))
    rv = rv.sort_values(["_o", "商品數"], ascending=[True, False])
    cols = ["key", "抽查", "品牌欄", "商品名稱【】", "館別/集團(參考)", "主要類目", "商品數", "範例商品名稱"] + RESULT_COLS
    rv = rv[cols].copy()
    for c in HUMAN_COLS:
        rv[c] = ""
    return rv


def build_stats(full, review, a, con):
    tot = len(full)
    typ = full["suggest brand name"].map(lambda x: x if x in (TYPE_NEW, TYPE_NB) else TYPE_POOL)
    s1 = full.assign(建議類型=typ).groupby(["建議類型", "信心程度"]).size().rename("商品數").reset_index()
    s1["商品占比"] = (s1["商品數"] / tot).map("{:.1%}".format)
    s2 = review.groupby("判斷路徑").agg(品牌字串數=("key", "size"), 商品數=("商品數", "sum")).reset_index()
    s2["信心"] = s2["判斷路徑"].str[:2].map(lambda c: PATHS.get(c, ("", ""))[1])
    s2["商品占比"] = (s2["商品數"] / tot).map("{:.1%}".format)
    s2["條件"] = s2["判斷路徑"].str[:2].map(lambda c: PATHS.get(c, ("", "", ""))[2])
    s3 = []
    for col in [c.strip() for c in a.breakdown.split(",") if c.strip() in full]:
        for val, d in full.assign(_t=typ).groupby(col):
            s3.append({"拆解欄位": col, "值": val, "商品數": len(d),
                       TYPE_POOL: f"{(d._t == TYPE_POOL).mean():.0%}", TYPE_NEW: f"{(d._t == TYPE_NEW).mean():.0%}",
                       TYPE_NB: f"{(d._t == TYPE_NB).mean():.0%}", "高信心": f"{(d['信心程度'] == '高').mean():.0%}",
                       "待人工": f"{(d['是否待人工判斷'] == '是').mean():.0%}"})
    log = pd.read_sql("SELECT * FROM review_log", con)
    acc = []
    if len(log):
        for dim, col in (("判斷路徑", "path"), ("信心程度", "conf")):
            for v, d in log.groupby(col):
                acc.append({"維度": dim, "值": f"{v} {PATHS[v][0]}" if v in PATHS else v, "審核數": len(d),
                            "系統建議正確": int(d.agree.sum()), "準確率": f"{d.agree.mean():.0%}",
                            "商品加權準確率": f"{(d.agree * d.sku).sum() / max(d.sku.sum(), 1):.0%}"})
        s = log[log.sampled == 1]
        if len(s):
            acc.append({"維度": "抽查★", "值": "高/中信心抽查", "審核數": len(s), "系統建議正確": int(s.agree.sum()),
                        "準確率": f"{s.agree.mean():.0%}",
                        "商品加權準確率": f"{(s.agree * s.sku).sum() / max(s.sku.sum(), 1):.0%}"})
    acc = pd.DataFrame(acc) if acc else pd.DataFrame({"說明": ["還沒有審核記錄，審核後再執行一次就會出現"]})
    return [("① 建議結果 × 信心", s1), ("② 判斷路徑（哪種情況佔多少）", s2),
            ("③ 類目拆解", pd.DataFrame(s3)), ("④ 系統準確率（累積所有審核記錄）", acc)]


def save_results(con, full, site, a):
    cols = ["site", "category", "item_id", "raw_brand"] + RESULT_COLS + ["tagged_at"]
    old = [r[1] for r in con.execute("PRAGMA table_info(item_results)")]
    if old and old != cols:
        con.execute("DROP TABLE item_results")
    con.execute("CREATE TABLE IF NOT EXISTS item_results(" + ", ".join(f'"{c}" TEXT' for c in cols) + ")")
    con.execute("DELETE FROM item_results WHERE site=? AND category=?", (site, a.category or ""))
    db = full[[a.col_id, a.col_brand] + RESULT_COLS].astype(str).copy()
    db.columns = ["item_id", "raw_brand"] + RESULT_COLS
    db.insert(0, "category", a.category or "")
    db.insert(0, "site", site)
    db["tagged_at"] = now()
    db.to_sql("item_results", con, if_exists="append", index=False, chunksize=20000)
    con.commit()


# ================================================================================ Excel
GUIDE = [
    "【每一輪怎麼做】",
    "1. 到「審核」頁。已經排好順序：待人工判斷 → 抽查★ → 其他（同順序內商品數多的在前）",
    "2. 在黃色的【人工】判斷 填（有下拉選單，也可以直接打字）：",
    "      v　　　　　　　　　 同意系統建議",
    "      1 / 2 / 3　　　　　 改選 shp brand name1 / 2 / 3",
    "      brand_id 或品牌名　 改成品牌庫的其他品牌",
    f"      {TYPE_NB}　　　　　 白牌、沒有品牌（再選【人工】No brand原因）",
    f"      {TYPE_NEW}　 有品牌但品牌庫沒有（名稱預設用品牌欄，要改名寫「新增:品牌名」）",
    "   不確定的留空就好。",
    "3. 存檔 → 關閉 Excel → 再執行一次同樣的指令",
    "   → 你的判斷會自動寫進規則庫，以後所有站點、品類遇到同樣的品牌字串都直接套用",
    "   → 產出新檔案（檔名有日期時間），舊檔自動移到 output/歷史",
    "",
    "【想針對單一商品修改】到「總結果」頁該品號的【人工】欄位填，只會影響這一個品號",
    "",
    "【多個候選時怎麼選：決策樹】",
    "   ① 相關度：先留下跟競品品牌最相關的一群（例：羅技 vs 雷技、羅剛 → 羅技）",
    "   ② 最小顆粒度：名稱完整出現在競品品牌中、而且最具體的優先（例：Apple iPhone → iPhone）",
    "   ③ 業績：剩下名稱相同或很相近的，選 adg 最高（例：LUX / Lux → 業績高的那個）",
    "   ④ 名稱不同卻同分 → 低信心，交給人判斷",
    "",
    "【信心與待人工】高 = 可直接用｜中 = 大致可信，靠抽查★量準確率｜低 = 一定要看",
    f"   是否待人工判斷 = 低信心 或 {TYPE_NEW}（因為可能只是品牌庫名稱語言不同，例：莉婕 = Liese）",
    "",
    "【判斷路徑一覽】",
]


def write_excel(path, review, full, newb, stats, a):
    with pd.ExcelWriter(path, engine="xlsxwriter") as w:
        wb = w.book
        F = {k: wb.add_format(v) for k, v in {
            "h_blue": {"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": True, "valign": "vcenter"},
            "h_yel": {"bold": True, "bg_color": "#FFD966", "border": 1, "text_wrap": True, "valign": "vcenter"},
            "h": {"bold": True, "bg_color": "#F2F2F2", "border": 1, "text_wrap": True, "valign": "vcenter"},
            "yel": {"bg_color": "#FFF2CC"}, "hi": {"bg_color": "#C6EFCE"}, "mid": {"bg_color": "#FFEB9C"},
            "lo": {"bg_color": "#FFC7CE"}, "title": {"bold": True, "font_size": 12}, "wrap": {"text_wrap": True}}.items()}

        # 說明
        ws = wb.add_worksheet("說明")
        w.sheets["說明"] = ws
        for i, line in enumerate(GUIDE):
            ws.write(i, 0, line, F["title"] if line.startswith("【") else None)
        r0 = len(GUIDE)
        for j, h in enumerate(["代碼", "名稱", "信心", "條件", "例子"]):
            ws.write(r0, j, h, F["h"])
        for i, (k, v) in enumerate(PATHS.items(), r0 + 1):
            ws.write_row(i, 0, [k, v[0], v[1], v[2], v[3]])
        ws.set_column(0, 0, 14)
        ws.set_column(1, 1, 28)
        ws.set_column(2, 2, 6)
        ws.set_column(3, 4, 60)

        def sheet(df, name, widths, freeze_col):
            df.to_excel(w, sheet_name=name, index=False)
            ws = w.sheets[name]
            cols, n = list(df.columns), len(df)
            for j, c in enumerate(cols):
                ws.write(0, j, c, F["h_yel"] if c in HUMAN_COLS else F["h_blue"] if c in RESULT_COLS else F["h"])
                ws.set_column(j, j, widths.get(c, 22 if c in RESULT_COLS + HUMAN_COLS else 12))
            ws.set_row(0, 32)
            ws.freeze_panes(1, freeze_col)
            ws.autofilter(0, 0, max(n, 1), len(cols) - 1)
            if n:
                j = cols.index("信心程度")
                for val, f in (("高", "hi"), ("中", "mid"), ("低", "lo")):
                    ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "==", "value": f'"{val}"',
                                                       "format": F[f]})
                for c in HUMAN_COLS:
                    j = cols.index(c)
                    ws.conditional_format(1, j, n, j, {"type": "no_errors", "format": F["yel"]})
                j = cols.index(H_PICK)
                ws.data_validation(1, j, n, j, {"validate": "list", "source": PICK_OPTIONS, "show_error": False,
                                                "input_title": "填寫方式",
                                                "input_message": "v=同意｜1/2/3=選候選｜也可打 brand_id 或品牌名"})
                j = cols.index(H_REASON)
                ws.data_validation(1, j, n, j, {"validate": "list", "source": REASONS, "show_error": False})
            ws.write_comment(0, cols.index(H_PICK),
                             f"v = 同意建議\n1/2/3 = 改選第幾個候選\nbrand_id 或品牌名 = 改成該品牌\n"
                             f"{TYPE_NB} = 白牌\n{TYPE_NEW} = 品牌庫沒有", {"x_scale": 1.6, "y_scale": 1.6})
            ws.set_column(cols.index("key"), cols.index("key"), None, None, {"hidden": True})

        sheet(review, "審核", {"抽查": 5, "品牌欄": 20, "範例商品名稱": 40, "判斷說明": 45, "判斷路徑": 24,
                              "商品數": 7}, 3)
        if len(full) > 1_000_000:
            sys.exit("✖ 單一品類超過 100 萬筆，Excel 放不下，請再細分品類")
        sheet(full, "總結果", {a.col_title: 40, "判斷說明": 40, "判斷路徑": 24}, 0)

        newb.to_excel(w, sheet_name="新增品牌建議", index=False)
        ws = w.sheets["新增品牌建議"]
        for j, c in enumerate(newb.columns):
            ws.write(0, j, c, F["h"])
            ws.set_column(j, j, 40 if c == "範例商品名稱" else 22)

        ws = wb.add_worksheet("統計")
        w.sheets["統計"] = ws
        r = 0
        for title, df in stats:
            ws.write(r, 0, title, F["title"])
            if len(df):
                df.to_excel(w, sheet_name="統計", startrow=r + 1, index=False)
                for j, c in enumerate(df.columns):
                    ws.write(r + 1, j, c, F["h"])
            r += len(df) + 4
        ws.set_column(0, 0, 30)
        ws.set_column(1, 8, 16)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="競網品牌 → shp 品牌庫 對應工具 v0.4")
    ap.add_argument("--items", required=True, help="競網商品資料：csv / xlsx / SQLite(.db)")
    ap.add_argument("--site", help="站點名稱，例如 momo")
    ap.add_argument("--category", help="只跑這個品類（對應 --col-cat 欄位的值），例如 Beauty")
    ap.add_argument("--col-cat", default="level1_category", help="品類欄位（預設 level1_category）")
    ap.add_argument("--pool", default="brand_pool.csv", help="shp 品牌庫（預設 brand_pool.csv）")
    ap.add_argument("--table", help="--items 是 SQLite 時的資料表名稱")
    ap.add_argument("--pool-table", help="--pool 是 SQLite 時的資料表名稱")
    ap.add_argument("--pool-col-cat", default="level1_category", help="品牌庫的類目欄位")
    ap.add_argument("--db", default="brand_rules.db", help="規則庫（預設 brand_rules.db）")
    ap.add_argument("--breakdown", default="level2_category,level3_category", help="統計拆解欄位，逗號分隔")
    ap.add_argument("--learn-only", action="store_true", help="只把人工判斷寫進規則庫，不重新產出")
    ap.add_argument("--no-db-results", action="store_true", help="不把結果寫進規則庫的 item_results 表")
    ap.add_argument("--col-brand", default="品牌")
    ap.add_argument("--col-title", default="商品名稱")
    ap.add_argument("--col-id", default="品號")
    ap.add_argument("--col-group", default="L2")
    run(ap.parse_args())
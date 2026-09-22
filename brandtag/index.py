"""品牌庫索引：把 7.6 萬個品牌建成可以快速比對的多重索引。

一個品牌會被拆成三把鑰匙：完整正規化名、只留英文、只留中文，
加上人工確認與自動學習的別名，讓「莉婕」也能對到「Liese 莉婕」。
"""
from __future__ import annotations

import re
import sqlite3
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz, process

from .const import SPECIAL_IDS
from .text import (CJK_RE, blank, clean_raw, is_bilingual, latin_tokens, long_enough, normalize, split_parts,
                   surface_key, STOP_TOKENS)

# 公司型態後綴：競網常省略，品牌庫常登記完整。多出來的只有這些字 → 視為同一家品牌
CORP_SUFFIX = {"文化", "出版", "出版社", "文創", "事業", "生技", "科技", "國際", "實業", "企業",
               "有限公司", "股份有限公司", "公司", "圖書", "文庫", "社", "行", "廠", "館",
               "文化事業", "出版事業", "生物科技", "文教", "傳媒", "媒體", "工作室"}

# 這些字單獨出現不足以當品牌證據（「前衛出版社」不該因為都有「出版社」就配到「DK出版社」）
GENERIC_KEYS = CORP_SUFFIX | {"書店", "書局", "書版", "書坊", "大學", "學院", "研究所", "基金會",
                              "百貨", "超市", "藥局", "藥妝", "診所", "醫院", "農場", "牧場",
                              "食品", "生活", "用品", "精品", "嚴選", "專賣", "旗艦", "選物",
                              "國際企業", "文化出版", "出版集團", "股份",
                              # 英文側同樣的問題：Penguin Books Ltd 不該因為都有 Ltd 就配到叫「LTD」的品牌
                              "ltd", "limited", "inc", "incorporated", "corp", "corporation", "llc", "plc",
                              "gmbh", "sas", "bv", "pte", "pty", "kk", "ag", "sa", "nv", "srl",
                              "publishing", "publishers", "publication", "publications", "press", "books",
                              "editions", "imprint", "enterprise", "enterprises", "industries", "holdings",
                              "ation", "ations", "cation", "sion", "ment", "ness"}


# ---------------------------------------------------------------- 品牌庫讀檔
def read_table(path: Path, must_have, table=None) -> pd.DataFrame:
    p = Path(path)
    if not p.exists():
        sys.exit(f"✖ 找不到檔案 {p}")
    if p.suffix.lower() in (".db", ".sqlite", ".sqlite3"):
        if not table:
            sys.exit(f"✖ {p.name} 是 SQLite，請在 config.toml 的 [pool] 指定 table")
        with sqlite3.connect(p) as con:
            df = pd.read_sql(f'SELECT * FROM "{table}"', con)
        return df.astype(object).where(df.notna(), "").astype(str)
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
        if raw is None:
            sys.exit(f"✖ {p.name} 編碼無法辨識（試過 utf-8 與 cp950）")
    for i in range(min(10, len(raw))):
        row = [str(x).strip() for x in raw.iloc[i].tolist()]
        if all(c in row for c in must_have):
            df = raw.iloc[i + 1:].copy()
            df.columns = row
            return df.reset_index(drop=True).fillna("")
    sys.exit(f"✖ {p.name} 前 10 列找不到欄位 {must_have}")


def load_pool(cfg) -> pd.DataFrame:
    pool = read_table(cfg.pool_path, ["brand_id", "brand_name"], cfg.pool_table)
    pool = pool[pd.to_numeric(pool["brand_id"], errors="coerce").notna()].copy()
    pool["brand_id"] = pool["brand_id"].astype(float).astype(int)
    pool["brand_name"] = pool["brand_name"].astype(str).str.strip()
    if "adg" in pool:
        pool["adg"] = pd.to_numeric(pool["adg"].astype(str).str.replace(",", ""), errors="coerce").fillna(0)
    return pool.drop_duplicates("brand_id")


# ---------------------------------------------------------------- 索引
class BrandIndex:
    def __init__(self, pool: pd.DataFrame, aliases: list, cat_col: str, th: dict,
                 entity_links: list = ()):
        self.th = th
        self.name = dict(zip(pool.brand_id, pool.brand_name))
        self.adg = dict(zip(pool.brand_id, pool["adg"])) if "adg" in pool else {}
        self.cat = dict(zip(pool.brand_id, pool[cat_col])) if cat_col in pool else {}
        self.parts = {b: split_parts(n) for b, n in self.name.items()}
        # 使用者確認：iPhone 是 Apple 的產品線，不是獨立品牌。
        # 只接受明確 Apple 登錄；Apple Sidra / Apple House 等不是 Apple Inc.
        self.apple_ids = {b for b, (full, _, _) in self.parts.items()
                          if full in {"apple", "apple蘋果", "蘋果apple", "iphoneapple", "appleiphone"}}
        self.apple_product_ids = {b for b, (full, _, _) in self.parts.items()
                                  if re.fullmatch(r"iphone(?:\d+[a-z0-9]*)?", full)}
        self.group = {b: p[0] for b, p in self.parts.items()}
        self.surface = defaultdict(set)
        for b, n in self.name.items():
            self.surface[surface_key(n)].add(b)
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
        # 別名保留來源層級。curated 是人工查證過的對照，不能和系統自己猜出的 auto
        # 混成同一種信任度；否則錯誤 auto 會在下一輪變成「中英文都相符」的假證據。
        self.human, self.curated, self.auto = set(), set(), set()
        # 同實體關係與文字別名分開保存，避免 H&H 正規化為 hh 後誤把其他 HH 品牌
        # 一起視為同實體。關係只確認實體；最後保留哪個 ID 仍交給 adg。
        self.verified_same = {
            frozenset((int(b1), int(b2))) for b1, b2, _src in entity_links
            if int(b1) in self.name and int(b2) in self.name and int(b1) != int(b2)
        }
        self.entity_peers = defaultdict(set)
        for pair in self.verified_same:
            b1, b2 = tuple(pair)
            self.entity_peers[b1].add(b2)
            self.entity_peers[b2].add(b1)
        for alias, b, src in aliases:
            if b in self.name and b not in SPECIAL_IDS:
                self.idx[alias].add(b)
                if src == "human":
                    self.human.add((alias, b))
                elif src == "curated":
                    self.curated.add((alias, b))
                elif (alias, b) not in self.natural:
                    self.auto.add((alias, b))
        self.tok = defaultdict(set)
        for b, n in self.name.items():
            if b not in SPECIAL_IDS:
                for t in latin_tokens(n):
                    self.tok[t].add(b)
        # 「出版社」「文化」這種公司型態字眼不能單獨當比對依據，否則「前衛出版社」會命中
        # 「DK出版社」—— 兩者只是都叫出版社而已。
        self.long_keys = {a for a in self.idx
                          if long_enough(a) and a not in GENERIC_KEYS and a not in STOP_TOKENS}
        # 中文前綴索引：競網寫「葡萄王」，品牌庫登記成「GRAPE KING BIO 葡萄王生技」。
        # 原本的子字串掃描只找得到「品牌庫的名字包在競網字串裡」，反過來的找不到，
        # 這一段就是補那個方向。用前 2 個字分桶，避免掃全表。
        self.cjk_bucket = defaultdict(list)
        for b, (full, lat, cjk) in self.parts.items():
            if b not in SPECIAL_IDS and len(cjk) >= 3:
                self.cjk_bucket[cjk[:2]].append((cjk, b))
        self._fz = {}
        self.fuzzy_by_len = defaultdict(list)
        for a in self.idx:
            if len(a) >= 5:
                self.fuzzy_by_len[(len(a), a[0])].append(a)

    def label(self, b) -> str:
        return f"{self.name.get(b, '?')} [{b}]"

    def adg_of(self, b) -> float:
        try:
            return float(self.adg.get(b, 0) or 0)
        except (ValueError, TypeError):
            return 0.0

    def fuzzy_choices(self, q):
        L, c = len(q), q[0]
        if (L, c) not in self._fz:
            lo, hi = int(L * 0.78), int(L / 0.78) + 1
            self._fz[(L, c)] = [a for n in range(lo, hi + 1) for a in self.fuzzy_by_len.get((n, c), [])]
        return self._fz[(L, c)]

    def match(self, s: str) -> dict:
        """回傳候選。

        value 欄位依序為：分數、說明、命中長度、是否人工別名、命中方式、
        證據來源（natural/human/curated/auto）、是否為「短英數與中文未被獨立證實」風險。
        """
        AUTO, FUZZY = self.th["auto"], self.th["fuzzy_cutoff"]
        out = {}
        full, lat, cjk = split_parts(s)
        if not full:
            return out
        bil = is_bilingual(clean_raw(s))

        def add(bids, score, desc, alias, via, source_override=None):
            for b in bids:
                h = (alias, b) in self.human
                if source_override:
                    src = source_override
                elif h:
                    src = "human"
                elif (alias, b) in self.curated:
                    src = "curated"
                elif (alias, b) in self.auto:
                    src = "auto"
                else:
                    src = "natural"
                if src == "human" and score >= AUTO:
                    desc2 = "人工確認過的別名"
                elif src == "curated" and score >= AUTO:
                    desc2 = f"人工整理並查證的品牌對照「{alias}」"
                elif src == "auto" and score >= AUTO:
                    desc2 = f"自動學習的別名「{alias}」"
                else:
                    desc2 = desc
                _, _pl, pc = self.parts[b]
                weak_short = (bool(cjk) and 1 <= len(lat) <= 3 and pc != cjk
                              and src not in ("human", "curated") and via in ("lat", "cjk", "full"))
                c = [score, desc2, len(alias), h, via, src, weak_short]
                trust = {"human": 3, "curated": 2, "natural": 1, "auto": 0}
                if b not in out or (c[0], trust[c[5]], c[2]) > (out[b][0], trust[out[b][5]], out[b][2]):
                    out[b] = c

        # 保留標點的精確命中優先，避免 H&H 與 HH 在去標點後混成一群。
        surface_exact = self.surface.get(surface_key(s), set())
        if surface_exact:
            add(surface_exact, 100, "品牌原始名稱（含標點）完全相符", full, "surface")
            peers = {peer for b in surface_exact for peer in self.entity_peers.get(b, set())}
            add(peers, 100, "人工查證為同一品牌實體", full, "entity", "curated")
        if full in self.idx and not surface_exact:
            exact = {b for b in self.idx[full]
                     if self.group[b] == full or (full, b) in self.human or (full, b) in self.curated}
            add(exact, 100, "品牌名完全相符", full, "full")
            add(self.idx[full] - exact, 95,
                "與品牌庫品牌的" + ("中文名" if CJK_RE.search(full) else "英文名") + "相符",
                full, "cjk" if CJK_RE.search(full) else "lat")
        hl = self.idx.get(lat, set()) if len(lat) >= 2 and lat != full else set()
        hc = self.idx.get(cjk, set()) if len(cjk) >= 2 and cjk != full else set()
        both = hl & hc
        # auto 別名不能和同一列的短英文互相佐證。GB 綠鐘曾先猜成 gb，再學到
        # 「綠鐘→gb」，下一輪就被誤升成「中英文皆相符」100 分。
        corroborated = {b for b in both if (lat, b) not in self.auto and (cjk, b) not in self.auto}
        add(corroborated, 100, "中英文名皆相符", full, "full")
        weak_l = len(lat) <= 3 and bool(cjk) and not bil
        weak_c = len(cjk) <= 2 and len(lat) >= 4 and not bil
        # 兩個字母的英文縮寫是很弱的證據：「FJ 豐傑生醫」的 FJ 可能對到任何叫 FJ 的品牌。
        # 壓到 85 分，讓「中文名完全相符」（88 分）能夠勝出。
        short_abbr = len(lat) <= 2 and bool(cjk)
        add(hl - corroborated, 70 if weak_l else 85 if short_abbr else 95,
            "英文縮寫過短且黏著中文（可能是集團/公司名）" if weak_l else
            "英文縮寫相符（僅兩個字母，證據較弱）" if short_abbr else "英文名相符", lat, "lat")
        add(hc - corroborated, 70 if weak_c else 95,
            "中文名過短（可能是泛用詞）" if weak_c else "中文名相符", cjk, "cjk")
        for i in range(len(full)):
            for j in range(i + 3, min(len(full), i + 25) + 1):
                sub = full[i:j]
                if sub != full and sub in self.long_keys:
                    add(self.idx[sub], 75, f"品牌名包含「{sub}」", sub, "sub")
        # 空格隔開的單字剛好等於某品牌完整名稱（例：Kanebo KATE 的 KATE、ASUS ROG 的 ROG）
        words = [normalize(w) for w in re.split(r"[\s/|｜]+", clean_raw(s))]
        if len(words) > 1:
            for w_ in words:
                if w_ and w_ != full and w_ not in STOP_TOKENS and w_ not in GENERIC_KEYS and (len(w_) >= 3 or CJK_RE.search(w_) and len(w_) >= 2):
                    add({b for b in self.idx.get(w_, set()) if self.group[b] == w_ or self.parts[b][1] == w_ or self.parts[b][2] == w_}, 85,
                        f"品牌名中的「{w_}」與品牌庫品牌相同", w_, "word")
        # 競網的中文名是品牌庫中文名的前綴（葡萄王 → 葡萄王生技、小熊出版 → 小熊出版社）。
        # 只在候選夠少時採用，否則像「台灣」這種前綴會命中一大堆不相干的品牌。
        # 兩個字的中文品牌在台灣非常多（商周、聯經、五南…），所以門檻放到 2 字，
        # 但兩字太容易誤中（國家、幸福、生活），因此要求多出來的字必須是公司型態後綴。
        if len(cjk) >= 2 and not self.idx.get(cjk):
            pre = [(pc, b) for pc, b in self.cjk_bucket.get(cjk[:2], []) if pc != cjk and pc.startswith(cjk)]
            short = len(cjk) == 2
            cap = 2 if short else 3
            if 0 < len(pre) <= cap:
                for pc, b in pre:
                    tail = pc[len(cjk):]
                    suffix_ok = tail in CORP_SUFFIX
                    if short and not suffix_ok:
                        continue                       # 兩字的一律要求後綴是公司型態，否則不採用
                    sure = len(pre) == 1 and (suffix_ok or len(tail) <= 2)
                    add({b}, 95 if sure else 85 if suffix_ok else 82,
                        f"品牌庫登記為「{self.name[b]}」（競網少了「{tail}」）", cjk, "cjkpre")

        for t in latin_tokens(s):
            add(self.tok.get(t, set()), 65, f"英文單字「{t}」相同", t, "tok")
        if not any(v[0] >= AUTO for v in out.values()):
            for q in {full, lat}:
                if len(q) >= 5:
                    for a, sc, _ in process.extract(q, self.fuzzy_choices(q), scorer=fuzz.ratio, limit=3,
                                                    score_cutoff=FUZZY):
                        if a != q:
                            add(self.idx[a], round(60 + (sc - FUZZY) / (100 - FUZZY) * 14),
                                f"拼字相似({sc:.0f}%)", a, "fuzzy")
        # 中文名完全相符時，兩個字母的英文縮寫就不該再競爭。
        # 「FJ 豐傑生醫」的 FJ 可以對到任何叫 FJ 的品牌，但「豐傑生醫」四個字完全相符
        # 幾乎確定就是那一家。不做這件事會讓 adg 只有 113 的 FJ 打敗 adg 17 萬的正主。
        if len(cjk) >= 3 and len(lat) <= 2:
            cjk_exact = {b for b, v in out.items() if v[4] == "cjk" and self.parts[b][2] == cjk}
            if cjk_exact:
                for b, v in out.items():
                    if b not in cjk_exact and v[4] == "lat":
                        v[0] = min(v[0], 65)
                        v[1] = f"{v[1]}；但另有品牌的中文名「{cjk}」完全相符，此英文縮寫證據較弱"

        # 中英文互相矛盾 → 大幅降分（Dove 多芬 ≠ Dove 德芙；CLEAR 淨 ≠ CLEAR 可麗兒）
        for b, v in out.items():
            if (v[4] == "full" and v[5] == "natural") or v[5] in ("human", "curated"):
                continue
            _, pl, pc = self.parts[b]
            # 競網字串帶有中文名，但只靠「其中一個英文詞」或子字串命中品牌庫的純英文品牌
            # → 那個中文名是區辨資訊，兩者多半不是同一個品牌。
            #   例：「Beauty Century 美世紀」不該配到品牌庫的「Century」。
            if cjk and not pc and v[4] in ("word", "sub", "tok", "fuzzy") and lat != pl:
                v[0] = min(v[0], 55)
                v[1] = f"{v[1]}，但競網品牌另有中文名「{cjk}」而品牌庫此品牌沒有"
                continue
            cjk_conf = bool(cjk) and bool(pc) and cjk not in pc and pc not in cjk
            lat_conf = len(lat) >= 2 and len(pl) >= 2 and lat not in pl and pl not in lat
            if cjk_conf and v[4] != "cjk":
                v[0], v[1] = min(v[0], 60), f"{v[1]}，但中文名不同（{pc}≠{cjk}）"
            elif lat_conf and v[4] == "cjk" and len(cjk) <= 2:
                # 只在中文名很短（可能是巧合）時，才讓英文名不同推翻中文名相符。
                # 三個字以上的中文品牌名完全相同是強證據，不該被英文拼法差異否決
                # （WeOrganic 唯有機 與 OUI ORGANIC 唯有機 是同一個品牌）。
                v[0], v[1] = min(v[0], 60), f"{v[1]}，但英文名不同（{pl}≠{lat}）"
            elif lat_conf and v[4] == "cjk":
                v[0], v[1] = min(v[0], 88), f"{v[1]}（英文拼法不同：{pl} / {lat}）"
            elif not lat and len(cjk) <= 2 and len(pl) >= 4 and v[4] == "cjk":
                # 競網僅 ≤2 字中文且無英文名，而品牌庫此品牌另有 ≥4 字英文名（例如「和平」vs「WAHEI FREIZ 和平」）
                v[0] = min(v[0], 65)
                v[1] = f"{v[1]}；但競網僅二字中文且無英文名，而品牌庫此品牌另有英文名「{pl}」，證據較弱"
        return self.canonical_products(out)

    def canonical_products(self, candidates):
        """品牌實體已確認後才比較 adg；產品線 ID 不作為最終品牌。"""
        related = (self.apple_ids | self.apple_product_ids) & candidates.keys()
        if not related:
            return candidates
        out = {b: v for b, v in candidates.items() if b not in related}
        if self.apple_ids:
            pick = max(self.apple_ids, key=lambda b: (self.adg_of(b), -b))
            evidence = list(max((candidates[b] for b in related), key=lambda v: v[0]))
            evidence[1] += f"；iPhone 為產品線，Apple 同實體依 adg 取 {self.name[pick]} [{pick}]"
            out[pick] = evidence
        return out

    def scan_title(self, title) -> dict:
        out = {}
        t = normalize(title)
        for i in range(len(t)):
            for j in range(i + 3, min(len(t), i + 20) + 1):
                sub = t[i:j]
                if sub in self.long_keys:
                    for b in self.idx[sub]:
                        if b not in out or out[b][2] < len(sub):
                            src = "human" if (sub, b) in self.human else "curated" if (sub, b) in self.curated else \
                                "auto" if (sub, b) in self.auto else "natural"
                            out[b] = [75, f"商品名稱內含「{sub}」", len(sub), src == "human", "scan", src, False]
        return self.canonical_products(out)

    def same_entity(self, b1, b2) -> bool:
        """兩個 brand_id 是否其實是同一個品牌在品牌庫中的重複登錄。

        這是品牌庫清理的核心判斷：名稱相同（僅大小寫/標點差異）、或一方的名稱
        完整包含另一方且差異部分只是公司型態後綴，視為同一實體。
        判定為同一實體時，系統取 adg 最高者，並把整組記錄下來供品牌庫合併。
        """
        if b1 == b2:
            return True
        if frozenset((b1, b2)) in self.verified_same:
            return True
        g1, g2 = self.group[b1], self.group[b2]
        if g1 == g2:
            return True
        short, long_ = (g1, g2) if len(g1) <= len(g2) else (g2, g1)
        if len(short) >= 2 and long_.startswith(short) and long_[len(short):] in CORP_SUFFIX:
            return True
        # 中英文其中一半完全相同，另一半不衝突（Liese 莉婕 vs Liese）
        _, l1, c1 = self.parts[b1]
        _, l2, c2 = self.parts[b2]
        # 短縮寫不能單獨證明實體相同：H&R 安室家不是汽車懸吊 H&R；GB 綠鐘也不是母嬰 gb。
        if l1 and l1 == l2 and len(l1) >= 4 and (not c1 or not c2 or c1 == c2):
            return True
        if c1 and c1 == c2 and (not l1 or not l2 or l1 == l2):
            return True
        return fuzz.ratio(g1, g2) >= 95

    def duplicate_groups(self) -> list:
        """品牌庫中疑似重複登錄的品牌分組，供品牌庫清理使用。"""
        from collections import defaultdict
        buckets = defaultdict(list)
        for b, (full, lat, cjk) in self.parts.items():
            if b in SPECIAL_IDS or not full:
                continue
            buckets[lat[:6] if lat else cjk[:2]].append(b)
        seen, groups = set(), []
        for members in buckets.values():
            if len(members) < 2 or len(members) > 200:
                continue
            for i, b1 in enumerate(members):
                if b1 in seen:
                    continue
                grp = [b1]
                for b2 in members[i + 1:]:
                    if b2 not in seen and self.same_entity(b1, b2):
                        grp.append(b2)
                if len(grp) > 1:
                    seen.update(grp)
                    groups.append(sorted(grp, key=lambda b: -self.adg_of(b)))
        return groups

    def names_close(self, g1: str, g2: str) -> bool:
        if g1 == g2:
            return True
        if min(len(g1), len(g2)) >= 3 and (g1 in g2 or g2 in g1):
            return True
        return fuzz.ratio(g1, g2) >= self.th["near_ratio"]

    def resolve(self, text):
        """人工填的字串 → brand_id。回傳 (brand_id, None) 或 (None, 錯誤訊息)。"""
        from .const import TYPE_NEW
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

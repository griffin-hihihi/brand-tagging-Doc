"""判斷引擎：MECE 決策樹 + 0–100 信心指數。

決策樹的第一個問題永遠是「品牌欄有沒有值」，兩條分支互斥且窮盡：

    1. 品牌欄有值
       1.1 明示無品牌                      → No brand
       1.2 是相容/副廠描述                  → No brand
       1.3 對到品牌庫，且品牌庫中唯一         → 品牌庫品牌
       1.4 對到品牌庫，但品牌庫有重複 → 取 adg 最高 → 品牌庫品牌
       1.5 對到多個「名稱不同」的品牌         → 品牌庫品牌（低信心）
       1.6 沒對到，但看起來是品牌            → 建議新增
       1.7 沒對到，而且不像品牌              → No brand
    2. 品牌欄空白
       2.1 商品是配件/相容品                → No brand
       2.2 【】對到品牌庫                   → 品牌庫品牌
       2.3 【】像品牌但品牌庫沒有            → 建議新增
       2.4 標題內文嚴格比對到品牌            → 品牌庫品牌
       2.5 都找不到                       → No brand

信心指數是單一的 0–100 數字，表示「系統對這個結論有多確定」，與結論類型無關：
判定為 No brand 也可以是 97 分。原本的「高/中/低」與「是否待人工判斷」兩欄由它取代。
"""
from __future__ import annotations

from .const import (CONF_COL, NB_ID, PATHS, R_COMPAT, R_DESC, R_EMPTY, R_OTHER, R_STORE, R_WORD,
                    TYPE_NB, TYPE_NEW, TYPE_POOL)
from .index import BrandIndex
from .text import (ACCESSORY_RE, COMPAT_RE, NOBRAND_WORDS, blank, clean_raw, extract_bracket,
                   host_brand, looks_like_brand, normalize)

# 各葉節點的基礎信心。實際分數會再依證據強度與扣分調整。
BASE = {
    "1.1": 97, "1.2": 90, "1.3": 98, "1.4": 93, "1.5": 45, "1.6": 85, "1.7": 80,
    "2.1": 88, "2.2": 90, "2.3": 72, "2.4": 58, "2.5": 85, "3.1": 100, "3.2": 100,
}
MATCH_TH = 75          # 低於此分數的候選只列出、不採用（最小顆粒度原則）


def _clamp(x) -> int:
    return max(0, min(100, int(round(x))))


def result(bi, path, conf, explain, ranked=(), bid=None, new_name="", reason="") -> dict:
    typ = PATHS[path][1]
    labels = [bi.label(b) for b, _ in list(ranked)[:3]] + ["", "", ""]
    name = bi.name[bid] if typ == TYPE_POOL else typ
    return {"shp brand name1": labels[0], "shp brand name2": labels[1], "shp brand name3": labels[2],
            "suggest brand name": name,
            "suggest brand id": str(bid) if typ == TYPE_POOL else ("0" if typ == TYPE_NB else ""),
            "新增品牌名稱": new_name if typ == TYPE_NEW else "",
            "No brand原因": reason if typ == TYPE_NB else "",
            CONF_COL: _clamp(conf),
            "判斷路徑": f"{path} {PATHS[path][0]}", "判斷說明": explain,
            "_type": typ, "_bid": bid}


def rank(cands, bi: BrandIndex, cat: str = "", cat_check: bool = False):
    """排序候選：類目相符優先 → 分數 → 證據來源 → 命中長度 → 業績。"""
    trust = {"human": 3, "curated": 2, "natural": 1, "auto": 0}
    def sort_key(kv):
        b, v = kv
        cat_pri = 1 if (cat_check and cat and bi.cat.get(b) == cat) else 0
        return (cat_pri, v[0], trust.get(v[5], 0), v[2], bi.adg_of(b))
    return sorted(cands.items(), key=sort_key, reverse=True)


def resolve_candidates(strong, comp_full, bi: BrandIndex, cat: str = "", cat_check: bool = False):
    """從一組高分候選中選出一個。回傳 (brand_id, 路徑, 說明, 信心調整)。

    順序即為業務規則的優先序：
      ① 人工確認／查證對照 → 先確認是哪個品牌實體
      ② 已確認為同一實體的品牌庫重複 → 取 adg 最高
      ③ 子品牌優先（最具體原則） → 母品牌與子品牌並存時，統一取子品牌
      ④ 名稱互異且同分 → 不武斷選擇，標為低信心
    """
    top_id, top = strong[0]
    if top[5] == "human":
        return top_id, "1.3", "人工確認過的別名", 0
    if top[5] == "curated":
        # curated 先證明品牌實體；若它也把品牌庫中的另一筆完整名稱連到此實體，
        # 再由 adg 選 canonical ID。H&H / Herb & Health 就走這條 1.4。
        same_entity = [(b, v) for b, v in strong if bi.same_entity(b, top_id)]
        others = [(b, v) for b, v in strong
                  if not bi.same_entity(b, top_id) and v[0] >= top[0] - 5]
        if len(same_entity) > 1 and not others:
            pick = max(same_entity, key=lambda kv: bi.adg_of(kv[0]))[0]
            dups = "、".join(dict.fromkeys(bi.name[b] for b, _ in same_entity))
            return pick, "1.4", (f"人工查證為同一品牌實體（{dups}），"
                                  f"依業績取 {bi.name[pick]}"), 0
        return top_id, "1.3", "人工整理並查證的品牌對照", 0

    # ① 品牌庫重複：名稱相同或高度近似的候選視為同一個實體的多筆登錄
    same_entity = [(b, v) for b, v in strong if bi.same_entity(b, top_id)]
    others = [(b, v) for b, v in strong if not bi.same_entity(b, top_id) and v[0] >= top[0] - 5]

    if len(same_entity) > 1 and not others:
        pick = max(same_entity, key=lambda kv: bi.adg_of(kv[0]))[0]
        dups = "、".join(dict.fromkeys(bi.name[b] for b, _ in same_entity))
        return pick, "1.4", f"品牌庫有 {len(same_entity)} 筆近似登錄（{dups}），依業績取 {bi.name[pick]}", 0

    if not others:
        return top_id, "1.3", top[1], 0

    # ② 子品牌優先（最具體原則）：母公司與子品牌同時出現時，統一取最具體的子品牌
    tied = [top_id] + [b for b, v in others if v[0] >= top[0] - 5]
    if cat_check and cat:
        same_cat_tied = [b for b in tied if bi.cat.get(b) == cat]
        if same_cat_tied:
            tied = same_cat_tied

    def match_span(b):
        full, lat, cjk = bi.parts[b]
        for piece in (full, lat, cjk):
            if piece and piece in comp_full:
                return comp_full.find(piece), len(piece)
        return -1, 0

    spans = {b: match_span(b) for b in tied}
    inside = {b: sp for b, sp in spans.items() if sp[0] != -1}
    # 出現位置不是母子品牌關係的證據。只對明確列出的品牌關係採用子品牌。
    known_children = {("kanebo", "kate"), ("asus", "rog")}
    confirmed_children = {child for parent in inside for child in inside if parent != child
                          and (bi.parts[parent][1], bi.parts[child][1]) in known_children}
    if confirmed_children:
        # 在電商命名慣例中，母品牌在前、子品牌在後（例：Kanebo KATE、Apple iPhone、ASUS ROG）
        # 出現位置較後（offset 較大）者為子品牌；若位置相同（巢狀包含），取較長者
        pick = max(confirmed_children, key=lambda b: (inside[b][1], bi.adg_of(b)))
        rest = "、".join(bi.name[b] for b in inside if b != pick)
        return pick, "1.3", f"{rest} 也出現在品牌字串中，依子品牌優先取較具體的 {bi.name[pick]}", -5

    # ③ 名稱互異、分數相當 → 交給人判斷
    if others and others[0][1][0] >= top[0] - 5:
        names = "、".join(bi.name[b] for b in [top_id] + [b for b, _ in others[:2]])
        return top_id, "1.5", f"{names} 同樣符合，名稱不同無法自動決定", 0
    return top_id, "1.3", top[1], 0


def _best(cands, bi, comp, cat, cat_check, scale=1.0, extra_note=""):
    """把一組候選收斂成 (brand_id, 路徑, 信心, 說明, ranked)；沒有夠強的候選回傳 None。"""
    ranked = rank(cands, bi, cat, cat_check)
    strong = [kv for kv in ranked if kv[1][0] >= MATCH_TH]
    if not strong:
        return None, ranked
    bid, path, note, adj = resolve_candidates(strong, comp, bi, cat, cat_check)
    best = dict(strong)[bid]
    # 證據強度：match() 給的分數越高，信心越高
    conf = BASE[path] * scale - max(0, (95 - best[0])) * 0.6 + adj
    if best[6]:
        conf = min(conf, 74)
        note += "；只有 1–3 字元英數縮寫相符，中文名稱未被獨立證實，需人工確認"
    if cat_check and cat and bi.cat.get(bid) and bi.cat[bid] != cat:
        is_partial = best[4] in ("word", "sub", "tok", "fuzzy", "cjkpre")
        conf -= (30 if is_partial else 10)
        note += f"；品牌類目（{bi.cat[bid]}）與商品類目（{cat}）不同"
        if is_partial and conf < MATCH_TH:
            return None, ranked
    if extra_note:
        note += extra_note
    ranked = [(bid, best)] + [kv for kv in ranked if kv[0] != bid]
    return (bid, path, conf, note, ranked), ranked


def tag(raw_brand, title, cat, bi: BrandIndex, cat_check: bool) -> dict:
    rb = clean_raw(raw_brand)
    t = "" if blank(title) else str(title)
    bracket = extract_bracket(t)

    # ══════════════════════════════════════════════ 1. 品牌欄有值
    if rb:
        if normalize(rb) in NOBRAND_WORDS:
            return result(bi, "1.1", BASE["1.1"], "品牌欄明確寫明無品牌", reason=R_WORD)
        if COMPAT_RE.search(rb):
            return result(bi, "1.2", BASE["1.2"], "品牌欄為相容/副廠描述", reason=R_COMPAT)
        if host_brand(rb) and (COMPAT_RE.search(t) or
                              (normalize(rb).startswith("iphone") and ACCESSORY_RE.search(t))):
            return result(bi, "1.2", 65, "主機品牌出現在相容配件，需確認實際製造商品牌", reason=R_COMPAT)

        hit, ranked = _best(bi.match(rb), bi, normalize(rb), cat, cat_check)
        if hit:
            bid, path, conf, note, ranked = hit
            # 品牌欄與【】指向不同品牌 → 降信心
            if bracket and normalize(bracket) != normalize(rb):
                tb = {b for b, v in bi.match(bracket).items() if v[0] >= 95}
                if tb and bid not in tb:
                    conf -= 20
                    note += "；品牌欄與商品名稱【】指向不同品牌"
            return result(bi, path, conf, note, ranked, bid=bid)

        # 沒對到品牌庫 → 是不是品牌？
        ok, why = looks_like_brand(rb)
        if ok:
            weak = ranked and ranked[0][1][0] >= 55
            conf = BASE["1.6"] - (20 if weak else 0)
            note = ("品牌庫查無相符品牌" if not weak else
                    f"品牌庫只有相似但不夠像的候選（{ranked[0][1][1]}），依最小顆粒度不採用")
            return result(bi, "1.6", conf, note, ranked, new_name=rb)
        return result(bi, "1.7", BASE["1.7"], f"品牌欄內容不是品牌：{why}", ranked, reason=R_STORE)

    # ══════════════════════════════════════════════ 2. 品牌欄空白
    # 先看商品名稱開頭【】：若有明確獨立第三方品牌，優先採認（例：【ipega】副廠Switch配件、【犀牛盾】iPhone保護殼）
    if bracket:
        norm_br = normalize(bracket)
        if norm_br in NOBRAND_WORDS or COMPAT_RE.search(bracket):
            return result(bi, "2.1", BASE["2.1"], f"商品名稱【】標示為相容描述或無品牌（{bracket}）", reason=R_COMPAT)

        ok, why = looks_like_brand(bracket)
        # 描述字串不能靠其中一個詞命中品牌；精確的既有品牌仍可採認。
        cands = bi.match(bracket)
        if not ok:
            cands = {b: v for b, v in cands.items() if v[4] in ("surface", "full") and v[0] >= 95}
        hit, ranked = _best(cands, bi, norm_br, cat, cat_check, scale=0.95,
                            extra_note="（依商品名稱【】判斷）")

        # 檢驗【】是否為「相容主機名稱」：例如【Apple】iPhone 15 副廠鋼化膜，【】內填的是主機名但標題宣告為副廠/相容
        is_compat_context = bool(COMPAT_RE.search(t) or ACCESSORY_RE.search(t))
        host_platforms = {"apple", "iphone", "ipad", "macbook", "airpods", "applewatch",
                          "samsung", "galaxy", "switch", "nintendo", "pixel",
                          "playstation", "ps4", "ps5", "xbox", "sony", "dyson",
                          "蘋果", "三星", "任天堂", "小米"}
        is_host_brand = (norm_br in host_platforms) or any(hp in norm_br for hp in ("iphone", "ipad", "switch", "macbook", "galaxy", "airpods"))

        if hit:
            bid, path, conf, note, ranked = hit
            # 若【】是主機平台且標題明確標示副廠/相容，則該主機品牌僅為相容對象，非商品製造商
            if is_compat_context and is_host_brand:
                return result(bi, "2.1", 65, f"商品為配件，【】中之品牌（{bracket}）可能是相容主機，需確認製造商", ranked, reason=R_COMPAT)
            return result(bi, "2.2", conf, note, ranked, bid=bid)

        if ok:
            if is_compat_context and is_host_brand:
                return result(bi, "2.1", 65, f"商品為配件，【】中之描述（{bracket}）可能是相容主機，需確認製造商", ranked, reason=R_COMPAT)
            return result(bi, "2.3", BASE["2.3"], "商品名稱【】看起來是品牌，但品牌庫查無", ranked,
                          new_name=bracket)
        # 【】不是品牌（純規格或促銷描述，如【台灣製造防窺片】） → 往下檢查是否為相容配件或掃描內文

    # 【】無獨立品牌時，若標題屬於相容周邊或配件（如 iPhone 15 鋼化膜） → 歸入 No brand (2.1)
    if COMPAT_RE.search(t) or ACCESSORY_RE.search(t):
        return result(bi, "2.1", BASE["2.1"], "商品為配件/相容品且無獨立品牌，標題中的品牌為其相容對象", reason=R_COMPAT)

    sc = {b: v for b, v in bi.scan_title(t).items() if _distinctive(v[1])}
    if sc:
        hit, ranked = _best(sc, bi, normalize(t), cat, cat_check, scale=1.0,
                            extra_note="（於商品名稱內文比對）")
        if hit:
            bid, _p, conf, note, ranked = hit
            return result(bi, "2.4", BASE["2.4"] - max(0, 95 - dict(ranked)[bid][0]) * 0.3, note,
                          ranked, bid=bid)

    reason = R_DESC if bracket else R_EMPTY
    note = "品牌欄空白，商品名稱【】為描述、內文也找不到品牌" if bracket else "品牌欄空白，商品名稱也找不到品牌"
    return result(bi, "2.5", BASE["2.5"], note, reason=reason)


def _distinctive(desc) -> bool:
    """標題掃到的片段要夠長夠獨特才算證據：英文 ≥6 字元、中文 ≥4 字。"""
    import re

    from .text import CJK_RE
    m = re.search(r"「(.+?)」", desc or "")
    frag = m.group(1) if m else ""
    cjk = len(CJK_RE.findall(frag))
    return cjk >= 4 if cjk else len(frag) >= 6


def apply_rule(res: dict, rule, code: str, bi: BrandIndex):
    """把人工判定蓋到自動結果上。人工判定的信心固定為 100。"""
    from .text import id_str
    typ, bid, name, reason = rule["decision"], rule["brand_id"], rule["brand_name"], rule["nobrand_reason"]
    res.update({"suggest brand name": bi.name.get(int(bid), name) if typ == TYPE_POOL and bid is not None else typ,
                "suggest brand id": id_str(bid) if typ == TYPE_POOL else ("0" if typ == TYPE_NB else ""),
                "新增品牌名稱": name if typ == TYPE_NEW else "",
                "No brand原因": (reason or "") if typ == TYPE_NB else "",
                CONF_COL: 100, "判斷路徑": f"{code} {PATHS[code][0]}",
                "判斷說明": PATHS[code][0], "_type": typ,
                "_bid": bid if typ == TYPE_POOL else (NB_ID if typ == TYPE_NB else None)})

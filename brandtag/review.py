"""審核檔的讀回與產生。

每個 L1 一個常駐檔：review/<site>/<L1>.xlsx
  審核    一列一個品牌字串 —— 你填【人工】欄位的地方
  單品覆蓋 只想改某一個品號時填這裡
  說明    填法與判斷路徑一覽

每次執行會「原地更新」這個檔：先把你填的判斷讀進規則庫，再依最新結果重寫整份檔案。
已審過的列會把你的判斷回填成正規化後的樣子（看得到系統確實記住了），並排到後面；
還沒審的排最前面。
"""
from __future__ import annotations

import re

import numpy as np
import pandas as pd

from .const import (ACCEPT_WORDS, H_NOTE, H_PICK, H_REASON, HUMAN_COLS, NB_ID, RESULT_COLS, REVOKE_WORDS,
                    ST_AUTO, ST_DONE_FIX, ST_DONE_OK, ST_PENDING, ST_SAMPLE, TYPE_NB, TYPE_NEW, TYPE_POOL,
                    CONF_COL, REVIEW_TH)
from .cluster import group_keys
from .store import REVOKED
from .text import NOBRAND_WORDS, blank, clean_raw, normalize

# 欄位順序刻意安排成「由左到右＝決策順序」：
# 先看狀態與品牌欄 → 再看系統建議與三個候選 → 然後就地填【人工】判斷。
# 你要選的候選一定在你要填的格子「左邊」，不需要左右來回捲動。
REVIEW_COLS = ["key", "狀態", "抽查", "歧義", "累積覆蓋", "群組", "品牌欄", "商品數",
               "範例商品名稱", "範例商品網址",
               "suggest brand name", CONF_COL,
               "shp brand name1", "shp brand name2", "shp brand name3"] \
    + HUMAN_COLS \
    + ["判斷路徑", "判斷說明", "商品名稱【】", "主要類目",
       "suggest brand id", "新增品牌名稱", "No brand原因", "上次審核"]
FREEZE_AT = REVIEW_COLS.index(H_NOTE) + 1          # 凍結到【人工】備註，往右捲動時決策資訊不會消失

ITEM_COLS = ["品號", "品牌欄", "商品名稱", "目前結果", "目前brand id"] + HUMAN_COLS + ["上次審核"]


def parse_human(pick, reason, note, row, bi):
    """人工輸入 → 判斷 dict / 錯誤字串 / None（沒填）。

    row 需要有：suggest brand name、suggest brand id、新增品牌名稱、No brand原因、shp brand name1~3。
    """
    if blank(pick) and blank(reason):
        return None
    note = "" if blank(note) else str(note).strip()
    reason = "" if blank(reason) else str(reason).strip()
    v = "" if blank(pick) else str(pick).strip()
    raw_input = v or reason

    if v.lower() in REVOKE_WORDS:                                  # 撤銷先前的判斷
        return {"decision": REVOKED, "brand_id": None, "brand_name": "", "nobrand_reason": "",
                "note": note, "raw_input": raw_input}

    def sugg_type():
        name = "" if blank(row.get("suggest brand name")) else str(row["suggest brand name"])
        return name if name in (TYPE_NB, TYPE_NEW) else TYPE_POOL

    if v.lower() in ACCEPT_WORDS:                                  # v = 同意系統建議
        t = sugg_type()
        if t == TYPE_POOL:
            if blank(row.get("suggest brand id")):
                return "系統沒有建議 brand id，請直接填品牌名或 brand_id"
            return {"decision": TYPE_POOL, "brand_id": int(float(row["suggest brand id"])),
                    "brand_name": str(row["suggest brand name"]), "nobrand_reason": "", "note": note,
                    "raw_input": raw_input}
        if t == TYPE_NEW:
            return {"decision": TYPE_NEW, "brand_id": None, "brand_name": str(row.get("新增品牌名稱") or ""),
                    "nobrand_reason": "", "note": note, "raw_input": raw_input}
        return {"decision": TYPE_NB, "brand_id": NB_ID, "brand_name": TYPE_NB,
                "nobrand_reason": reason or str(row.get("No brand原因") or ""), "note": note,
                "raw_input": raw_input}

    if re.fullmatch(r"[123](\.0)?", v):                            # 1/2/3 = 改選第幾個候選
        cand = row.get(f"shp brand name{int(float(v))}")
        if blank(cand):
            return f"沒有第 {int(float(v))} 個候選"
        v = str(cand)

    if not v or normalize(v) in NOBRAND_WORDS or v == TYPE_NB:
        return {"decision": TYPE_NB, "brand_id": NB_ID, "brand_name": TYPE_NB,
                "nobrand_reason": reason or "人工判定無品牌", "note": note, "raw_input": raw_input}

    if v.startswith((TYPE_NEW, "建議品牌庫新增", "新增", "新品牌")):
        name = re.split(r"[:：]", v, maxsplit=1)[1].strip() if re.search(r"[:：]", v) else ""
        if not name:
            for c in ("新增品牌名稱", "品牌欄", "商品名稱【】"):
                if not blank(row.get(c)):
                    name = clean_raw(row[c])
                    break
        if not name:
            return "請寫成「新增:品牌名」"
        return {"decision": TYPE_NEW, "brand_id": None, "brand_name": name, "nobrand_reason": "",
                "note": note, "raw_input": raw_input}

    bid, err = bi.resolve(v)
    if err:
        return err
    return {"decision": TYPE_POOL, "brand_id": bid, "brand_name": bi.name[bid], "nobrand_reason": "",
            "note": note, "raw_input": raw_input}


def read_existing(path, sheet, dtype=str):
    if not path.exists():
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet, dtype=dtype)
    except (ValueError, KeyError):
        return None


def collect_decisions(path, bi, site) -> tuple[list, list]:
    """從審核檔讀回人工判斷。回傳 (判斷清單, 錯誤訊息清單)。"""
    entries, errors = [], []
    rv = read_existing(path, "審核")
    if rv is not None and H_PICK in rv:
        for _, row in rv.iterrows():
            d = parse_human(row.get(H_PICK), row.get(H_REASON), row.get(H_NOTE), row, bi)
            if d is None:
                continue
            label = row.get("品牌欄") if not blank(row.get("品牌欄")) else row.get("範例商品名稱")
            if isinstance(d, str):
                errors.append(f"審核頁｜{label}：{d}")
                continue
            if blank(row.get("key")):
                errors.append(f"審核頁｜{label}：這一列的 key 不見了（不要刪欄或插入空列）")
                continue
            sysname = "" if blank(row.get("suggest brand name")) else str(row["suggest brand name"])
            systype = sysname if sysname in (TYPE_NB, TYPE_NEW) else TYPE_POOL
            sysid = None if blank(row.get("suggest brand id")) else int(float(row["suggest brand id"]))
            d.update({
                "scope": "brand", "rule_key": str(row["key"]).strip(),
                "raw_brand": "" if blank(row.get("品牌欄")) else str(row["品牌欄"]),
                "sys_decision": systype, "sys_brand_id": sysid,
                "sys_path": str(row.get("判斷路徑") or "").split()[0] if not blank(row.get("判斷路徑")) else "",
                "sys_conf": row.get(CONF_COL),
                "sampled": int(str(row.get("抽查") or "") == "★"),
                "goods": int(float(row["商品數"])) if not blank(row.get("商品數")) else 0,
                "agree": int(d["decision"] == systype and (d["decision"] != TYPE_POOL or d["brand_id"] == sysid)),
            })
            entries.append(d)

    it = read_existing(path, "單品覆蓋")
    if it is not None and H_PICK in it:
        for _, row in it.iterrows():
            if blank(row.get("品號")):
                continue
            d = parse_human(row.get(H_PICK), row.get(H_REASON), row.get(H_NOTE), row, bi)
            if d is None:
                continue
            if isinstance(d, str):
                errors.append(f"單品覆蓋頁｜品號 {row.get('品號')}：{d}")
                continue
            d.update({"scope": "item", "rule_key": str(row["品號"]).strip(),
                      "raw_brand": "" if blank(row.get("品牌欄")) else str(row["品牌欄"]),
                      "sys_decision": None, "sys_brand_id": None, "sys_path": "", "sys_conf": "",
                      "sampled": 0, "goods": 1, "agree": None})
            entries.append(d)
    return entries, errors


def _prefill(rule):
    """已審過的列：把規則庫裡的結論回填成人看得懂的樣子。"""
    if rule is None:
        return "", "", ""
    d = rule["decision"]
    if d == TYPE_POOL:
        pick = f"{rule['brand_name']} [{rule['brand_id']}]"
    elif d == TYPE_NEW:
        pick = f"{TYPE_NEW}:{rule['brand_name']}" if rule["brand_name"] else TYPE_NEW
    else:
        pick = TYPE_NB
    return pick, rule["nobrand_reason"] or "", rule["note"] or ""


def build_review(kinfo, kres, rules, sample_n, rng_seed=42) -> pd.DataFrame:
    rv = kinfo.join(kres).reset_index().rename(columns={
        "raw": "品牌欄", "bracket": "商品名稱【】", "cat": "主要類目", "sku": "商品數", "title": "範例商品名稱",
        "url": "範例商品網址", "index": "key"})
    rv["抽查"] = ""
    decided = rv["key"].map(lambda k: rules.get(k))
    pre = decided.map(_prefill)
    rv[H_PICK] = [p[0] for p in pre]
    rv[H_REASON] = [p[1] for p in pre]
    rv[H_NOTE] = [p[2] for p in pre]
    rv["上次審核"] = decided.map(lambda r: "" if r is None else f"{r['updated_at']}　{r['actor']}")

    # 抽查★：高/中信心、系統自動判的，依商品數加權抽樣（開根號避免大品牌壟斷樣本）
    is_rule = decided.notna()
    conf = pd.to_numeric(rv[CONF_COL], errors="coerce").fillna(0)
    # 只從「影響 ≥3 件商品」的列抽樣。長尾的單件商品列佔了母體大半，
    # 不排除的話樣本會被它們灌滿，量不到真正重要的判斷準不準。
    cand = rv[conf.ge(REVIEW_TH) & ~is_rule & rv["商品數"].astype(float).ge(3)]
    for lo, hi in ((REVIEW_TH, 90), (90, 101)):
        sub = cand[conf.loc[cand.index].between(lo, hi - 1)]
        if len(sub) <= sample_n:
            pick = sub.index
        else:
            w = sub["商品數"].astype(float) ** 0.5
            pick = np.random.default_rng(rng_seed).choice(sub.index, size=sample_n, replace=False,
                                                          p=(w / w.sum()).values)
        rv.loc[pick, "抽查"] = "★"

    # 已審的列再分「同意 / 修正」：拿人工結論跟系統自己原本算出來的（_auto_*，尚未套規則）比
    def same_as_auto(row):
        r = rules.get(row["key"])
        if r is None:
            return False
        if r["decision"] != row["_auto_type"]:
            return False
        return r["decision"] != TYPE_POOL or r["brand_id"] == row["_auto_bid"]

    agreed = rv.apply(same_as_auto, axis=1) if len(rv) else pd.Series(dtype=bool)
    rv["狀態"] = np.select(
        [is_rule & agreed, is_rule, conf.lt(REVIEW_TH), rv["抽查"].eq("★")],
        [ST_DONE_OK, ST_DONE_FIX, ST_PENDING, ST_SAMPLE], default=ST_AUTO)
    # 歧義：系統說高信心，但品牌庫裡還有第 2 個近似品牌 —— 這種其實值得瞄一眼
    rv["歧義"] = np.where(conf.ge(90) & rv["shp brand name2"].astype(str).str.strip().ne(""), "⚠", "")

    # 分群：把長得像的待審品牌字串放在一起，一次決定一群
    pend = rv[rv["狀態"].isin([ST_PENDING, ST_SAMPLE])]
    src = [(k, b if str(b).strip() else t) for k, b, t in
           zip(pend["key"], pend["品牌欄"], pend["商品名稱【】"])]
    gmap = group_keys(src) if len(src) > 1 else {}
    rv["群組"] = rv["key"].map(gmap).fillna("")

    order = {ST_PENDING: 0, ST_SAMPLE: 1, ST_DONE_FIX: 2, ST_DONE_OK: 3, ST_AUTO: 4}
    rv["_o"] = rv["狀態"].map(order)
    # 同群組的排在一起，且用群組內最大商品數當排序值，讓大群組仍然排前面
    rv["_g"] = rv.groupby("群組")["商品數"].transform("max").where(rv["群組"].ne(""), rv["商品數"])
    rv = rv.sort_values(["_o", "_g", "群組", "商品數"], ascending=[True, False, True, False])

    # 累積覆蓋：審到這一列為止，已經涵蓋多少比例的商品。看到 80% 就可以先收工
    tot = max(rv["商品數"].sum(), 1)
    rv["累積覆蓋"] = (rv["商品數"].cumsum() / tot).map("{:.1%}".format)

    for c in REVIEW_COLS:
        if c not in rv:
            rv[c] = ""
    return rv[REVIEW_COLS].copy()


def build_items(full, overrides, col) -> pd.DataFrame:
    """單品覆蓋頁：只列已經被單獨指定過的品號（其他人自己貼上去）。"""
    if not overrides:
        return pd.DataFrame(columns=ITEM_COLS)
    ids = set(overrides)
    sub = full[full[col["id"]].astype(str).isin(ids)]
    rows = []
    for _, r in sub.iterrows():
        o = overrides[str(r[col["id"]])]
        pick, reason, note = _prefill(o)
        rows.append({"品號": r[col["id"]], "品牌欄": r["brand"], "商品名稱": r["title"],
                     "目前結果": r["suggest brand name"], "目前brand id": r["suggest brand id"],
                     H_PICK: pick, H_REASON: reason, H_NOTE: note,
                     "上次審核": f"{o['updated_at']}　{o['actor']}"})
    return pd.DataFrame(rows, columns=ITEM_COLS)

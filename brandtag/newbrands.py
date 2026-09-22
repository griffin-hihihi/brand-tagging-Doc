"""跨 L1 的「待決策品牌總表」。

一個品牌字串可能同時出現在好幾個 L1（例如「3M」在居家、3C、汽車都有）。
一個一個 L1 去審會重複看到同一個品牌；這份總表把它們合併，讓負責品牌庫的人
一次把所有待決策的品牌看完，而且看得到每個決策能解決多少商品。

同時區分兩種完全不同的情況：
  A. 純中文品牌、品牌庫只有英文名 → 需要人工配對（例：理膚寶水 = La Roche Posay）
  B. 品牌庫真的沒有這個品牌      → 需要決定要不要收錄（例：出版社）
兩者的處理方式不一樣，混在一起看會做錯決策。
"""
from __future__ import annotations

import pandas as pd

from .const import CONF_COL, REVIEW_TH, TYPE_NEW
from .text import CJK_RE, split_parts

ACT_PAIR = "A 可能是品牌庫既有品牌的中文名 → 請配對"
ACT_ADD = "B 尚未找到品牌 → 查證是否為品牌及是否已收錄"
ACT_CHECK = "C 有相似候選 → 請確認是否同一品牌"

COLS = ["排名", "品牌字串", "總商品數", "累積覆蓋", "出現在幾個L1", "L1清單", "語言", "建議動作",
        "最相似的品牌庫品牌", "系統判斷", "平均信心", "範例商品名稱"]


def _lang(s: str) -> str:
    _, lat, cjk = split_parts(s)
    if cjk and lat:
        return "中英並列"
    if cjk:
        return "純中文"
    return "純英文"


def build(con, site: str) -> pd.DataFrame:
    """從 item_results 彙總所有待人工判斷的品牌字串。"""
    d = pd.read_sql(
        "SELECT level1, raw_brand, title, 新增品牌名稱, `suggest brand name` AS sugg, "
        "`shp brand name1` AS cand1, `信心指數` AS conf, 判斷路徑 "
        "FROM item_results WHERE site=? AND (CAST(`信心指數` AS REAL) < ? OR `suggest brand name`=?)",
        con, params=(site, REVIEW_TH, TYPE_NEW))
    if d.empty:
        return pd.DataFrame(columns=COLS)

    d["名稱"] = d["新增品牌名稱"].where(d["新增品牌名稱"].astype(str).str.strip().ne(""), d["raw_brand"])
    d["名稱"] = d["名稱"].astype(str).str.strip()
    d = d[d["名稱"].ne("")]

    g = d.groupby("名稱").agg(
        總商品數=("名稱", "size"),
        出現在幾個L1=("level1", "nunique"),
        L1清單=("level1", lambda s: "、".join(sorted(set(s))[:6])),
        最相似的品牌庫品牌=("cand1", lambda s: next((x for x in s if str(x).strip()), "")),
        系統判斷=("sugg", lambda s: s.mode().iloc[0] if len(s.mode()) else ""),
        平均信心=("conf", lambda s: round(pd.to_numeric(s, errors="coerce").mean())),
        範例商品名稱=("title", "first"),
    ).reset_index().rename(columns={"名稱": "品牌字串"})

    g = g.sort_values("總商品數", ascending=False).reset_index(drop=True)
    g["排名"] = g.index + 1
    g["累積覆蓋"] = (g["總商品數"].cumsum() / g["總商品數"].sum()).map("{:.1%}".format)
    g["語言"] = g["品牌字串"].map(_lang)

    has_cand = g["最相似的品牌庫品牌"].astype(str).str.strip().ne("")
    g["建議動作"] = ACT_ADD
    g.loc[has_cand, "建議動作"] = ACT_CHECK
    # 純中文、系統找不到任何候選 → 很可能是國際品牌的中文名，品牌庫只收了英文
    g.loc[g["語言"].eq("純中文") & ~has_cand, "建議動作"] = ACT_PAIR
    return g[COLS]


def summary(g: pd.DataFrame) -> pd.DataFrame:
    if g.empty:
        return pd.DataFrame({"說明": ["沒有待決策品牌"]})
    s = g.groupby("建議動作").agg(品牌數=("品牌字串", "size"), 商品數=("總商品數", "sum")).reset_index()
    s["商品占比"] = (s["商品數"] / g["總商品數"].sum()).map("{:.1%}".format)
    return s.sort_values("商品數", ascending=False)

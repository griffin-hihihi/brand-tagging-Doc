"""給主管看的摘要檔：幾張一眼能看懂的表，不是 54 萬列的原始資料。

Google Sheet 單檔上限 1,000 萬格，全站明細有 927 萬格 —— 塞得進去但會慢到不能用，
而且主管要的是結論。這裡產出的是可以直接上傳、開啟即讀的摘要。
"""
from __future__ import annotations

import pandas as pd

from .const import CONF_COL, REVIEW_TH, TYPE_NB, TYPE_NEW, TYPE_POOL, band


def _type(s: pd.Series) -> pd.Series:
    return s.map(lambda x: x if x in (TYPE_NEW, TYPE_NB) else TYPE_POOL)


def build(con, site: str) -> dict:
    d = pd.read_sql("SELECT level1, `suggest brand name` AS s, `信心指數` AS conf, "
                    "raw_brand, 新增品牌名稱 FROM item_results WHERE site=?", con, params=(site,))
    d["conf"] = pd.to_numeric(d["conf"], errors="coerce").fillna(0)
    tot = max(len(d), 1)
    d["類型"] = _type(d["s"])
    pend = d["conf"].lt(REVIEW_TH)

    總覽 = pd.DataFrame([
        ("商品總數（品號）", f"{tot:,}", ""),
        ("系統自動完成", f"{(~pend).sum():,}", f"{(~pend).mean():.1%}"),
        ("需人工確認", f"{pend.sum():,}", f"{pend.mean():.1%}"),
        ("", "", ""),
        ("對應到品牌庫既有品牌", f"{(d.類型 == TYPE_POOL).sum():,}", f"{(d.類型 == TYPE_POOL).mean():.1%}"),
        ("品牌庫沒有、建議新增", f"{(d.類型 == TYPE_NEW).sum():,}", f"{(d.類型 == TYPE_NEW).mean():.1%}"),
        ("判定為無品牌（白牌）", f"{(d.類型 == TYPE_NB).sum():,}", f"{(d.類型 == TYPE_NB).mean():.1%}"),
        ("", "", ""),
        ("平均信心指數", f"{d['conf'].mean():.1f}", ""),
        ("信心 90 以上（可直接採用）", f"{d['conf'].ge(90).sum():,}", f"{d['conf'].ge(90).mean():.1%}"),
        (f"信心 {REVIEW_TH}–89（大致可信）", f"{d['conf'].between(REVIEW_TH, 89).sum():,}",
         f"{d['conf'].between(REVIEW_TH, 89).mean():.1%}"),
        (f"信心 {REVIEW_TH} 以下（需人工）", f"{d['conf'].lt(REVIEW_TH).sum():,}", f"{d['conf'].lt(REVIEW_TH).mean():.1%}"),
    ], columns=["項目", "數量", "占比"])

    cat = d.groupby("level1").apply(lambda g: pd.Series({
        "商品數": len(g),
        "自動完成率": f"{g['conf'].ge(REVIEW_TH).mean():.0%}",
        "對應到品牌庫": f"{(g.類型 == TYPE_POOL).mean():.0%}",
        "建議新增": f"{(g.類型 == TYPE_NEW).mean():.0%}",
        "無品牌": f"{(g.類型 == TYPE_NB).mean():.0%}",
        "待人工確認": int(g["conf"].lt(REVIEW_TH).sum()),
    }), include_groups=False).reset_index().rename(columns={"level1": "商品類別"})
    cat = cat.sort_values("商品數", ascending=False)

    # 人工確認的投入與回報曲線
    p = d[pend].copy()
    p["名稱"] = p["新增品牌名稱"].where(p["新增品牌名稱"].astype(str).str.strip().ne(""), p["raw_brand"])
    g = p.groupby(p["名稱"].astype(str).str.strip()).size().sort_values(ascending=False)
    g = g[g.index != ""]
    auto = (~pend).sum()
    rows = []
    for n in (50, 100, 200, 500, 1000, 2000):
        if n <= len(g):
            cov = int(g.head(n).sum())
            rows.append({"人工確認品牌數": n, "可解決商品數": f"{cov:,}",
                         "占待確認比例": f"{cov / max(pend.sum(), 1):.0%}",
                         "全站覆蓋率": f"{(auto + cov) / tot:.1%}",
                         "估計投入": {50: "約 1 小時", 100: "約 2 小時", 200: "約半天",
                                    500: "約 1 天", 1000: "約 2 天", 2000: "約 4 天"}[n]})
    投報 = pd.DataFrame(rows)

    return {"總覽": 總覽, "各類別": cat, "人工投入與回報": 投報}


def write(path, tables: dict, top_brands: pd.DataFrame):
    with pd.ExcelWriter(path, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        wb = w.book
        hf = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": True,
                            "valign": "vcenter"})
        big = wb.add_format({"bold": True, "font_size": 13})
        for name, df in tables.items():
            df.to_excel(w, sheet_name=name, index=False, startrow=1)
            ws = w.sheets[name]
            ws.write(0, 0, {"總覽": "momo 品牌對應 — 總覽",
                            "各類別": "各商品類別的對應結果",
                            "人工投入與回報": "人工確認要花多少力氣、能換到多少覆蓋率"}[name], big)
            for j, c in enumerate(df.columns):
                ws.write(1, j, c, hf)
                ws.set_column(j, j, 24 if j == 0 else 15)
            ws.freeze_panes(2, 0)
        top_brands.to_excel(w, sheet_name="待決策品牌Top200", index=False, startrow=1)
        ws = w.sheets["待決策品牌Top200"]
        ws.write(0, 0, "最值得優先處理的 200 個品牌（依影響商品數排序）", big)
        for j, c in enumerate(top_brands.columns):
            ws.write(1, j, c, hf)
            ws.set_column(j, j, 30 if c in ("品牌字串", "建議動作", "最相似的品牌庫品牌") else 14)
        ws.freeze_panes(2, 0)
        ws.autofilter(1, 0, len(top_brands) + 1, len(top_brands.columns) - 1)

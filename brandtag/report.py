"""統計頁：結果分布、判斷路徑佔比、類目拆解、系統累積準確率。"""
from __future__ import annotations

import pandas as pd

from .const import CONF_COL, PATHS, REVIEW_TH, TYPE_NB, TYPE_NEW, TYPE_POOL, band


def type_of(series: pd.Series) -> pd.Series:
    return series.map(lambda x: x if x in (TYPE_NEW, TYPE_NB) else TYPE_POOL)


def build(full: pd.DataFrame, review: pd.DataFrame, con, breakdown, site, level1) -> list:
    tot = max(len(full), 1)
    typ = type_of(full["suggest brand name"])

    conf = pd.to_numeric(full[CONF_COL], errors="coerce").fillna(0)
    s1 = (full.assign(建議類型=typ, 信心分級=conf.map(band))
          .groupby(["建議類型", "信心分級"]).size().rename("商品數").reset_index())
    s1["商品占比"] = (s1["商品數"] / tot).map("{:.1%}".format)

    s2 = review.groupby("判斷路徑").agg(品牌字串數=("key", "size"), 商品數=("商品數", "sum")).reset_index()
    s2["結論類型"] = s2["判斷路徑"].str.split().str[0].map(lambda c: PATHS.get(c, ("", ""))[1] or "")
    s2["商品占比"] = (s2["商品數"] / tot).map("{:.1%}".format)
    s2["條件"] = s2["判斷路徑"].str.split().str[0].map(lambda c: PATHS.get(c, ("", "", ""))[2])
    s2 = s2.sort_values("商品數", ascending=False)

    s3 = []
    for col in [c for c in breakdown if c in full]:
        for val, d in full.assign(_t=typ).groupby(col):
            s3.append({"拆解欄位": col, "值": val, "商品數": len(d),
                       TYPE_POOL: f"{(d._t == TYPE_POOL).mean():.0%}", TYPE_NEW: f"{(d._t == TYPE_NEW).mean():.0%}",
                       TYPE_NB: f"{(d._t == TYPE_NB).mean():.0%}",
                       "平均信心": f"{pd.to_numeric(d[CONF_COL], errors='coerce').mean():.0f}",
                       "待人工": f"{pd.to_numeric(d[CONF_COL], errors='coerce').lt(REVIEW_TH).mean():.0%}"})

    log = pd.read_sql("SELECT * FROM decision_log WHERE scope='brand' AND agree IS NOT NULL", con)
    acc = []
    if len(log):
        log["conf_band"] = pd.to_numeric(log["sys_conf"], errors="coerce").fillna(0).map(band)
        for dim, col in (("判斷路徑", "sys_path"), ("信心分級", "conf_band")):
            for v, d in log[log[col].astype(str) != ""].groupby(col):
                acc.append({"維度": dim, "值": f"{v} {PATHS[v][0]}" if v in PATHS else v, "審核數": len(d),
                            "系統建議正確": int(d.agree.sum()), "準確率": f"{d.agree.mean():.0%}",
                            "商品加權準確率": f"{(d.agree * d.goods).sum() / max(d.goods.sum(), 1):.0%}"})
        s = log[log.sampled == 1]
        if len(s):
            acc.append({"維度": "抽查★", "值": "高/中信心抽查", "審核數": len(s), "系統建議正確": int(s.agree.sum()),
                        "準確率": f"{s.agree.mean():.0%}",
                        "商品加權準確率": f"{(s.agree * s.goods).sum() / max(s.goods.sum(), 1):.0%}"})
        for v, d in log.groupby("level1"):
            acc.append({"維度": "L1", "值": v or "(未標)", "審核數": len(d), "系統建議正確": int(d.agree.sum()),
                        "準確率": f"{d.agree.mean():.0%}",
                        "商品加權準確率": f"{(d.agree * d.goods).sum() / max(d.goods.sum(), 1):.0%}"})
    acc = pd.DataFrame(acc) if acc else pd.DataFrame({"說明": ["還沒有審核記錄，審核後再執行一次就會出現"]})

    return [("① 建議結果 × 信心", s1), ("② 判斷路徑（哪種情況佔多少）", s2),
            ("③ 類目拆解", pd.DataFrame(s3)), ("④ 系統準確率（累積所有 L1 的審核記錄）", acc)]


def status_table(cfg, con, l1_df: pd.DataFrame) -> pd.DataFrame:
    """每個 L1 現在進行到哪。"""
    runs = pd.read_sql("SELECT * FROM run_log WHERE site=? ORDER BY run_id", con, params=(cfg.site,))
    last = runs.drop_duplicates("level1", keep="last").set_index("level1") if len(runs) else None
    rows = []
    for _, r in l1_df.iterrows():
        d = {"L1": r["l1"], "cluster": r["cluster"], "品號": int(r["goods"]), "SKU": int(r["skus"]),
             "已跑過": "", "品牌字串": "", "待審": "", "待審商品": "", "品牌庫命中": "", "建議新增": "",
             "No brand": ""}
        if last is not None and r["l1"] in last.index:
            x = last.loc[r["l1"]]
            g = max(int(x["goods"] or 0), 1)
            d.update({"已跑過": str(x["ran_at"])[:16], "品牌字串": int(x["keys"] or 0),
                      "待審": int(x["pending_keys"] or 0),
                      "待審商品": f"{int(x['pending_goods'] or 0):,} ({(x['pending_goods'] or 0) / g:.0%})",
                      "品牌庫命中": f"{(x['pool_goods'] or 0) / g:.0%}",
                      "建議新增": f"{(x['new_goods'] or 0) / g:.0%}", "No brand": f"{(x['nb_goods'] or 0) / g:.0%}"})
        rows.append(d)
    return pd.DataFrame(rows)

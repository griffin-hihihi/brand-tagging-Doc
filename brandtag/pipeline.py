"""一個 L1 的完整流程。

    ① 讀回 review/<site>/<L1>.xlsx 的人工判斷 → 追加到 decision_log → 重建 brand_rules
    ② 用最新的規則庫 + 品牌庫比對這個 L1 的所有品號
    ③ 原地更新審核檔、產出結果檔、寫進 item_results、記一筆 run_log
"""
from __future__ import annotations

import time
from pathlib import Path

import pandas as pd

from . import excel, report, review as rv, source, store
from .const import (CONF_COL, HUMAN_COLS, RESULT_COLS, REVIEW_TH, TYPE_NB, TYPE_NEW, TYPE_POOL)
from .engine import apply_rule, tag
from .index import BrandIndex, load_pool
from .text import blank, brand_key, clean_raw, extract_bracket, is_bilingual, normalize, split_parts


def locked(path: Path) -> bool:
    """檔案是否正被 Excel 佔用而寫不進去。

    不用 ~$ 暫存檔判斷 —— Excel 當掉後那個檔會殘留，會害之後永遠寫不了。
    直接試著以寫入模式開檔：Windows 上 Excel 會持有獨佔鎖，開不起來才是真的被佔用。
    """
    if not path.exists():
        return False
    try:
        with path.open("r+b"):
            return False
    except (PermissionError, OSError):
        return True


def run_l1(cfg, con, pool, l1: str, log=print, write_excel=True) -> dict:
    t0 = time.time()
    c = cfg.col
    review_path = cfg.site_review_dir / f"{_safe(l1)}.xlsx"
    log(f"\n━━━ {cfg.site} / {l1} ━━━")

    # ── ① 讀回人工判斷 ──────────────────────────────────────────────
    bi = BrandIndex(pool, store.aliases(con), cfg.pool_col_cat, cfg.th, store.entity_links(con))
    n_imp = n_agree = 0
    if review_path.exists():
        entries, errors = rv.collect_decisions(review_path, bi, cfg.site)
        n_imp, n_agree = store.record(con, entries, review_path.name, cfg.site, l1)
        for e in errors[:30]:
            log(f"   ⚠ {e}")
        if len(errors) > 30:
            log(f"   ⚠ 另有 {len(errors) - 30} 筆錯誤未顯示")
        if errors:
            log(f"   ⚠ 以上 {len(errors)} 筆沒寫入，請在更新後的審核檔重新填")
        # 人工確認的品牌 → 存成別名，之後同義寫法也對得到
        pairs = []
        for e in entries:
            if e["scope"] == "brand" and e["decision"] == TYPE_POOL and e.get("raw_brand"):
                full, lat, cjk = split_parts(clean_raw(e["raw_brand"]))
                for a in {full, lat if len(lat) >= 4 else "", cjk if len(cjk) >= 2 else ""} - {""}:
                    pairs.append((a, e["brand_id"]))
        if pairs:
            store.add_aliases(con, pairs, "human")
        log(f"① 讀回人工判斷：新增/異動 {n_imp} 筆"
            + (f"（同意建議 {n_agree}、修正 {n_imp - n_agree}）" if n_imp else ""))
    else:
        log("① 尚未有審核檔（第一次跑這個 L1）")

    # ── ② 比對 ────────────────────────────────────────────────────
    bi = BrandIndex(pool, store.aliases(con), cfg.pool_col_cat, cfg.th,
                    store.entity_links(con))     # 帶上剛學到的別名與同實體關係
    items = source.load_l1(cfg, l1)
    if items.empty:
        log(f"   ⚠ 這個 L1 沒有資料，略過")
        return {}
    items = items.rename(columns={"item_id": c["id"]})
    items["key"] = [brand_key(b, t, i) for b, t, i in zip(items["brand"], items["title"], items[c["id"]])]

    g = items.groupby("key", sort=False)
    kinfo = g.agg(raw=("brand", "first"), title=("title", "first"), url=("url", "first"), sku=(c["id"], "size"))
    kinfo["bracket"] = items.groupby("key", sort=False)["title"].first().map(extract_bracket)
    cc = items.groupby(["key", "l2"]).size().reset_index(name="n").sort_values("n", ascending=False)
    kinfo["cat"] = cc.drop_duplicates("key").set_index("key")["l2"]
    kinfo["l1cat"] = l1

    cat_check = bool(bi.cat) and l1 in set(bi.cat.values())
    rules = store.current(con, "brand")
    log(f"② 品牌庫 {len(pool):,} 品牌｜商品 {len(items):,} 品號｜品牌字串 {len(kinfo):,} 個"
        + ("" if cat_check else "｜（品牌庫沒有這個類目名稱，略過類目檢查）"))

    rows = []
    for k, r in kinfo.iterrows():
        res = tag(r["raw"], r["title"], l1, bi, cat_check)
        res["_auto_type"], res["_auto_bid"] = res["_type"], res["_bid"]
        if k in rules:
            apply_rule(res, rules[k], "3.1", bi)
        rows.append(res)
    kres = pd.DataFrame(rows, index=kinfo.index)
    n_alias = _learn_auto_aliases(kres, kinfo, bi, con)
    hit = len(rules.keys() & set(kinfo.index))
    log(f"③ 判斷完成：套用規則庫 {hit:,} 個字串"
        + (f"｜自動學到 {n_alias} 個新別名" if n_alias else ""))

    # 展開到每個品號 + 單品覆蓋
    full = items.join(kres, on="key")
    ovr = store.current(con, "item", cfg.site)
    if ovr:
        mask = full[c["id"]].astype(str).isin(ovr)
        for i in full.index[mask]:
            res = full.loc[i, RESULT_COLS + ["_type", "_bid"]].to_dict()
            apply_rule(res, ovr[str(full.at[i, c["id"]])], "3.2", bi)
            for col, v in res.items():
                full.at[i, col] = v

    # ── ③ 輸出 ────────────────────────────────────────────────────
    review = rv.build_review(kinfo, kres, rules, cfg.sample_n)
    stats = report.build(full, review, con, cfg.breakdown, cfg.site, l1)
    newb = (review[review["suggest brand name"] == TYPE_NEW]
            [["新增品牌名稱", "品牌欄", "商品數", CONF_COL, "判斷路徑", "shp brand name1", "範例商品名稱"]]
            .rename(columns={"shp brand name1": "最相似的品牌庫品牌"})
            .sort_values("商品數", ascending=False))

    if write_excel:
        if locked(review_path):
            log(f"   ⚠ {review_path.name} 還開著 → 這次不覆寫審核檔（你填的已經存進規則庫了）。"
                f"請關閉 Excel 後再跑一次。")
        else:
            items_sheet = rv.build_items(full, ovr, c)
            excel.write_review(review_path, review, items_sheet)
        out_path = cfg.site_output_dir / f"{_safe(l1)}_result.xlsx"
        if locked(out_path):
            log(f"   ⚠ {out_path.name} 還開著 → 這次不覆寫結果檔。")
        else:
            drop = [x for x in ["_type", "_bid", "_auto_type", "_auto_bid", "key", "sku_codes"] if x in full]
            excel.write_result(out_path, full.drop(columns=drop), newb, stats, "title")

    out_cols = ["item_id", "raw_brand", "title"] + RESULT_COLS
    db_df = full.rename(columns={c["id"]: "item_id", "brand": "raw_brand"})[out_cols]
    store.save_results(con, db_df, cfg.site, l1, RESULT_COLS)

    vc = report.type_of(full["suggest brand name"]).value_counts()
    pend = pd.to_numeric(review[CONF_COL], errors="coerce").fillna(0).lt(REVIEW_TH)
    m = {"site": cfg.site, "level1": l1, "goods": len(full), "keys": len(kinfo), "imported": n_imp,
         "agreed": n_agree, "pending_keys": int(pend.sum()), "pending_goods": int(review.loc[pend, "商品數"].sum()),
         "pool_goods": int(vc.get(TYPE_POOL, 0)), "new_goods": int(vc.get(TYPE_NEW, 0)),
         "nb_goods": int(vc.get(TYPE_NB, 0)), "seconds": round(time.time() - t0, 1)}
    store.log_run(con, **m)

    log(f"④ 完成（{m['seconds']:.0f} 秒）"
        f"｜{TYPE_POOL} {m['pool_goods']:,}（{m['pool_goods'] / len(full):.0%}）"
        f"｜{TYPE_NEW} {m['new_goods']:,}｜{TYPE_NB} {m['nb_goods']:,}")
    log(f"   待你審核：{m['pending_keys']:,} 個品牌字串（涵蓋 {m['pending_goods']:,} 品號）"
        f"＋抽查★ {int(review['抽查'].eq('★').sum())} 個")
    log(f"   審核檔：{review_path.relative_to(cfg.root)}")
    return m


def _safe(s: str) -> str:
    import re
    return re.sub(r'[\\/:*?"<>|]+', "_", s).strip() or "L1"


def _learn_auto_aliases(kres, kinfo, bi, con) -> int:
    """「英文 中文」高信心相符 → 另一半自動存成別名。

    1–3 字元英數縮寫不自動學中文：短縮寫本身無法證明品牌實體，若把中文寫回，
    下一輪會形成「英文命中 + 自己學到的中文命中」的循環證據。
    """
    ok = (pd.to_numeric(kres[CONF_COL], errors="coerce").fillna(0).ge(90) & kres["_type"].eq(TYPE_POOL))
    pairs = []
    for k in kres.index[ok]:
        raw = clean_raw(kinfo.at[k, "raw"])
        if not is_bilingual(raw) or blank(kres.at[k, "_bid"]):
            continue
        bid = int(kres.at[k, "_bid"])
        _, lat, cjk = split_parts(raw)
        if cjk and len(lat) <= 3:
            continue
        for a in (lat if len(lat) >= 4 else "", cjk if len(cjk) >= 2 else ""):
            if a and a not in bi.idx:                          # 品牌庫還沒有這個名稱才學
                pairs.append((a, bid))
    return store.add_aliases(con, pairs, "auto") if pairs else 0

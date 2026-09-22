"""單一大審核表（打包 22 個品類給 Temp 審核，審完直接回灌）。"""
from __future__ import annotations

import sqlite3
from pathlib import Path
import pandas as pd

from .const import (CONF_COL, H_NOTE, H_PICK, H_REASON, HUMAN_COLS, NB_ID,
                    REVIEW_TH, ST_AUTO, ST_DONE_FIX, ST_DONE_OK, ST_PENDING, ST_SAMPLE,
                    TYPE_NB, TYPE_NEW, TYPE_POOL)
from .index import BrandIndex, load_pool
from .review import FREEZE_AT, parse_human
from .text import blank, clean_raw, normalize, split_parts, brand_key
from . import store

BUNDLE_COLS = ["key", "狀態", "抽查", "處理分類", "明確動作", "套用範圍", "跨類目衝突", "各類目建議", "出現在哪些L1", "品牌欄", "商品數",
               "範例商品名稱", "範例商品網址",
               "suggest brand name", CONF_COL,
               "shp brand name1", "shp brand name2", "shp brand name3"] \
    + HUMAN_COLS \
    + ["判斷路徑", "判斷說明", "商品名稱【】", "主要類目",
       "suggest brand id", "新增品牌名稱", "No brand原因", "上次審核", "規則版本"]


def export_bundle(cfg, out_path: Path | None = None, log=print) -> Path:
    """把全站所有 L1 的審核項目彙總去重成單一 Excel，供 Temp 集中審核。"""
    review_dir = cfg.site_review_dir
    files = [f for f in review_dir.glob("*.xlsx")
             if not f.name.startswith("~$") and not f.name.startswith("_")]
    if not files:
        raise FileNotFoundError(f"在 {review_dir} 找不到任何品類審核檔，請先跑 py run.py tag")

    log(f"讀取 {len(files)} 個品類審核檔…")
    merged: dict[str, dict] = {}
    
    for f in files:
        l1_name = f.stem
        try:
            df = pd.read_excel(f, sheet_name="審核", dtype=str)
        except Exception as exc:
            raise ValueError(f"無法讀取審核檔 {f.name}，停止以免產出不完整總表") from exc
        if df.empty or "key" not in df.columns:
            continue
            
        for row in df.to_dict(orient="records"):
            k = str(row.get("key") or "").strip()
            if not k:
                continue
            goods = int(float(row.get("商品數") or 0)) if str(row.get("商品數") or "").strip() else 0
            st = str(row.get("狀態") or "").strip()
            samp = str(row.get("抽查") or "").strip()
            
            if k not in merged:
                item = {col: ("" if pd.isna(row.get(col)) else str(row.get(col))) for col in BUNDLE_COLS if col in row}
                item["key"] = k
                item["_l1s"] = {l1_name}
                item["_goods"] = goods
                item["_st_has_pending"] = st.startswith("①")
                item["_st_has_sample"] = st.startswith("②")
                item["_st_has_done"] = st.startswith("③")
                item["_has_sample_star"] = (samp == "★")
                item["_outcomes"] = set()
                item["_advice"] = []
                merged[k] = item
            else:
                item = merged[k]
                item["_l1s"].add(l1_name)
                item["_goods"] += goods
                if st.startswith("①"):
                    item["_st_has_pending"] = True
                elif st.startswith("②"):
                    item["_st_has_sample"] = True
                elif st.startswith("③"):
                    item["_st_has_done"] = True
                if samp == "★":
                    item["_has_sample_star"] = True
            item = merged[k]
            outcome = tuple("" if blank(row.get(c)) else str(row[c]) for c in
                            ("suggest brand name", "suggest brand id", "新增品牌名稱", "No brand原因"))
            item["_outcomes"].add(outcome)
            item["_advice"].append(f"{l1_name}: {outcome[0]} [{outcome[1]}] {outcome[2]}")
                    
    # 建立 key -> (title, url) 映射庫，確保 100% 每一筆品牌字串都有代表性商品名稱與點擊網址
    key_goods_map: dict[str, tuple[str, str]] = {}
    if hasattr(cfg, "cache_path") and cfg.cache_path and cfg.cache_path.exists():
        log("建立商品名稱與商品網址對照庫…")
        try:
            con_g = sqlite3.connect(cfg.cache_path)
            for b_val, t_val, i_val, u_val in con_g.execute("SELECT brand, title, item_id, url FROM goods"):
                k_val = brand_key(b_val, t_val, i_val)
                if k_val not in key_goods_map:
                    u_final = u_val or f"https://www.momoshop.com.tw/product/{i_val}"
                    key_goods_map[k_val] = (t_val or "", u_final)
            con_g.close()
        except Exception as err:
            log(f"   (快取讀取提示: {err})")

    log(f"彙總完成，全站唯一待審/審核品牌字串共 {len(merged):,} 個，建立排序…")
    
    rows = []
    order_map = {1: 1, 2: 2, 3: 3, 4: 4}
    for item in merged.values():
        item["出現在哪些L1"] = "、".join(sorted(item.pop("_l1s")))
        item["商品數"] = item.pop("_goods")
        conflict = len(item.pop("_outcomes")) > 1
        item["跨類目衝突"] = "是" if conflict else "否"
        item["各類目建議"] = "；".join(item.pop("_advice"))
        
        k = item["key"]
        if k in key_goods_map:
            t_sample, u_sample = key_goods_map[k]
            if not item.get("範例商品名稱"):
                item["範例商品名稱"] = t_sample
            if not item.get("範例商品網址"):
                item["範例商品網址"] = u_sample

        if item.pop("_st_has_pending"):
            item["狀態"] = ST_PENDING
            ord_val = 1
        elif item.pop("_st_has_sample"):
            item["狀態"] = ST_SAMPLE
            ord_val = 2
        elif item.pop("_st_has_done"):
            item["狀態"] = ST_DONE_OK
            ord_val = 3
        else:
            item["狀態"] = ST_AUTO
            ord_val = 4
            
        item["抽查"] = "★" if item.pop("_has_sample_star") else ""
        item["套用範圍"] = "僅此品號 " + k[2:] if k.startswith("I:") else "全站相同 key；不同商品品牌須改用單品覆蓋"
        if conflict:
            item["狀態"], ord_val = ST_PENDING, 1
            item["處理分類"] = "A 跨類目衝突／主管處理"
            item["明確動作"] = "Temp 查各類目商品並記錄證據；人工判斷留白，交主管拆分單品"
            for col in HUMAN_COLS:
                item[col] = ""
        elif item.get("suggest brand name") == TYPE_NEW:
            item["處理分類"] = "B 新品牌查證"
            item["明確動作"] = "確認是品牌而非描述；查品牌庫，已有填 ID，確無則填新增:正式品牌名並附證據"
            if not str(item.get("狀態", "")).startswith("③"):
                item["狀態"], ord_val = ST_PENDING, 1
        elif k.startswith("I:"):
            item["處理分類"] = "C 單品／配件辨識"
            item["明確動作"] = "查此商品的製造品牌；相容主機不是品牌，無法確認留白並記錄待確認"
        elif item.get("suggest brand name") == TYPE_NB:
            item["處理分類"] = "D 無品牌確認"
            item["明確動作"] = "確認頁面無品牌後填 No brand 與原因；有品牌填 ID 或新增:名稱"
        else:
            item["處理分類"] = "E 既有品牌配對"
            item["明確動作"] = "核對品牌實體後填 ID／候選序號；同實體重複 ID 才比較 adg"
        item["_ord"] = ord_val
        rows.append(item)
        
    out_df = pd.DataFrame(rows)
    out_df = out_df.sort_values(["_ord", "商品數"], ascending=[True, False]).reset_index(drop=True)
    out_df = out_df[[c for c in BUNDLE_COLS if c in out_df.columns]]

    target = out_path or (review_dir / f"_{cfg.site}_全站審核總表_Temp用.xlsx")
    target.parent.mkdir(parents=True, exist_ok=True)

    widths = {
        "key": 22, "狀態": 10, "抽查": 6, "出現在哪些L1": 22, "品牌欄": 22, "商品數": 8,
        "範例商品名稱": 42, "範例商品網址": 38,
        "suggest brand name": 22, CONF_COL: 8, "shp brand name1": 22,
        "shp brand name2": 22, "shp brand name3": 22, H_PICK: 16, H_REASON: 20,
        H_NOTE: 18, "判斷路徑": 24, "判斷說明": 34, "商品名稱【】": 18, "主要類目": 14,
        "suggest brand id": 14, "新增品牌名稱": 16, "No brand原因": 16, "上次審核": 18
    }

    log(f"輸出 Excel 檔至 {target.name}…")
    with pd.ExcelWriter(target, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        out_df.to_excel(w, sheet_name="全站審核", index=False)
        ws = w.sheets["全站審核"]
        wb = w.book

        hf = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": True})
        pick_fmt = wb.add_format({"bold": True, "bg_color": "#FFF2CC", "border": 1})  # 人工欄黃色高亮

        for j, col in enumerate(out_df.columns):
            fmt = pick_fmt if col in HUMAN_COLS else hf
            ws.write(0, j, col, fmt)
            ws.set_column(j, j, widths.get(col, 14))

        # 凍結首列與前 6 欄（key, 狀態, 抽查, L1, 品牌欄, 商品數）
        ws.freeze_panes(1, 6)
        ws.autofilter(0, 0, max(len(out_df), 1), len(out_df.columns) - 1)

    log(f"[OK] 成功產出：{target}")
    return target


def import_bundle(cfg, in_path: Path | None = None, log=print) -> tuple[int, int]:
    """將 Temp 審核完成的總表讀回，寫入 brand_rules.db 規則庫。"""
    target = in_path or (cfg.site_review_dir / f"_{cfg.site}_全站審核總表_Temp用.xlsx")
    if not target.exists():
        raise FileNotFoundError(f"找不到檔案：{target}")

    con = store.open_db(cfg.db_path)
    pool = load_pool(cfg)
    bi = BrandIndex(pool, store.aliases(con), cfg.pool_col_cat, cfg.th, store.entity_links(con))

    log(f"讀取審核表 {target.name}…")
    df = pd.read_excel(target, sheet_name="全站審核", dtype=str)
    if H_PICK not in df.columns:
        raise ValueError(f"格式不符：找不到「{H_PICK}」欄位")

    entries, errors = [], []
    for _, row in df.iterrows():
        d = parse_human(row.get(H_PICK), row.get(H_REASON), row.get(H_NOTE), row, bi)
        if d is None:
            continue
        label = row.get("品牌欄") if not blank(row.get("品牌欄")) else row.get("範例商品名稱")
        if isinstance(d, str):
            errors.append(f"{label}：{d}")
            continue
        if blank(row.get("key")):
            errors.append(f"{label}：key 不見了")
            continue

        sysname = "" if blank(row.get("suggest brand name")) else str(row["suggest brand name"])
        systype = sysname if sysname in (TYPE_NB, TYPE_NEW) else TYPE_POOL
        sysid = None if blank(row.get("suggest brand id")) else int(float(row["suggest brand id"]))
        d.update({
            "expected_log_id": int(float(row["規則版本"])) if not blank(row.get("規則版本")) else 0,
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

    conflicts = df.get("跨類目衝突", pd.Series("否", index=df.index)).eq("是")
    if (conflicts & df[H_PICK].fillna("").str.strip().ne("")).any():
        errors.append("跨類目衝突列不可整組匯入，請清空判斷並由主管處理單品覆蓋")
    if errors:
        con.close()
        raise ValueError("匯入驗證未通過，未寫入任何判斷：\n" + "\n".join(errors[:20]))
    n_imp, n_agree = store.record(con, entries, target.name, cfg.site, level1="ALL")

    for e in errors[:20]:
        log(f"   ⚠ {e}")
    if len(errors) > 20:
        log(f"   ⚠ 另有 {len(errors) - 20} 筆錯誤未顯示")

    pairs = []
    for e in entries:
        if e["scope"] == "brand" and e["decision"] == TYPE_POOL and e.get("raw_brand"):
            full, lat, cjk = split_parts(clean_raw(e["raw_brand"]))
            for a in {full, lat if len(lat) >= 4 else "", cjk if len(cjk) >= 2 else ""} - {""}:
                pairs.append((a, e["brand_id"]))
    if pairs:
        store.add_aliases(con, pairs, "human")

    con.close()
    return n_imp, n_agree

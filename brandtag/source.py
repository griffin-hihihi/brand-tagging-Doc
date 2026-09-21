"""從競網 raw SQLite 抽出「品號層」的工作快取。

raw 檔是 SKU 層、每列一包 JSON（787k 筆），每次跑都重新解析太慢。
這裡一次把它攤平成一張 goods 表（545k 筆，品號唯一），之後每個 L1 只要 SELECT 就好。
raw 檔換了（大小或修改時間變動）會自動重建快取。
"""
from __future__ import annotations

import json
import sqlite3
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

GOODS_COLS = ["item_id", "cluster", "l1", "l2", "l3", "brand", "title", "url", "price", "reviews", "sales",
              "momo_l1", "momo_l2", "momo_l3", "sku_count", "sku_codes"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS goods(
  item_id TEXT PRIMARY KEY, cluster TEXT, l1 TEXT, l2 TEXT, l3 TEXT, brand TEXT, title TEXT, url TEXT,
  price TEXT, reviews TEXT, sales TEXT, momo_l1 TEXT, momo_l2 TEXT, momo_l3 TEXT,
  sku_count INTEGER, sku_codes TEXT);
CREATE INDEX IF NOT EXISTS ix_goods_l1 ON goods(l1);
CREATE TABLE IF NOT EXISTS cache_meta(k TEXT PRIMARY KEY, v TEXT);
"""


def _stamp(p: Path) -> str:
    st = p.stat()
    return f"{st.st_size}:{int(st.st_mtime)}"


def ensure_cache(cfg, force=False, log=print) -> Path:
    raw, cache = cfg.source_path, cfg.cache_path
    if not raw.exists():
        sys.exit(f"✖ 找不到 raw 資料 {raw}（請確認 config.toml 的 [source] path）")
    cache.parent.mkdir(parents=True, exist_ok=True)
    want = _stamp(raw)
    if cache.exists() and not force:
        try:
            con = sqlite3.connect(cache)
            got = con.execute("SELECT v FROM cache_meta WHERE k='stamp'").fetchone()
            n = con.execute("SELECT count(*) FROM goods").fetchone()[0]
            con.close()
            if got and got[0] == want and n:
                return cache
        except sqlite3.Error:
            pass
    log(f"   抽取 raw 資料 → 快取（第一次或 raw 有更新，需要一兩分鐘）…")
    _build(cfg, want, log)
    return cache


def _build(cfg, stamp: str, log):
    c = cfg.col
    tmp = cfg.cache_path.with_suffix(".building")
    tmp.unlink(missing_ok=True)
    out = sqlite3.connect(tmp)
    out.executescript(SCHEMA)
    src = sqlite3.connect(f"file:{cfg.source_path}?mode=ro", uri=True)

    counts: Counter = Counter()
    batch, n = [], 0
    ins = ("INSERT OR IGNORE INTO goods(item_id,cluster,l1,l2,l3,brand,title,url,price,reviews,sales,"
           "momo_l1,momo_l2,momo_l3,sku_count,sku_codes) VALUES(" + ",".join("?" * 16) + ")")
    for (payload,) in src.execute(f'SELECT payload FROM "{cfg.source_table}"'):
        d = json.loads(payload)
        gid = str(d.get(c["id"]) or "").strip()
        if not gid:
            continue
        counts[gid] += 1
        if counts[gid] == 1:
            batch.append((gid, d.get(c["cluster"]), d.get(c["l1"]), d.get(c["l2"]), d.get(c["l3"]),
                          d.get(c["brand"]), d.get(c["title"]), d.get(c["url"]),
                          str(d.get(c["price"]) or ""), str(d.get(c["reviews"]) or ""),
                          str(d.get(c["sales"]) or ""), d.get(c["momo_l1"]), d.get(c["momo_l2"]),
                          d.get(c["momo_l3"]), 0, ""))
        n += 1
        if len(batch) >= 20000:
            out.executemany(ins, batch)
            batch.clear()
        if n % 200000 == 0:
            log(f"      …{n:,} SKU")
    if batch:
        out.executemany(ins, batch)
    out.executemany("UPDATE goods SET sku_count=? WHERE item_id=?", ((v, k) for k, v in counts.items()))
    out.execute("INSERT OR REPLACE INTO cache_meta VALUES('stamp',?)", (stamp,))
    out.execute("INSERT OR REPLACE INTO cache_meta VALUES('skus',?)", (str(n),))
    out.commit()
    g = out.execute("SELECT count(*) FROM goods").fetchone()[0]
    out.close()
    src.close()
    cfg.cache_path.unlink(missing_ok=True)
    tmp.rename(cfg.cache_path)
    log(f"   快取完成：{n:,} SKU → {g:,} 品號")


def list_l1(cfg) -> pd.DataFrame:
    con = sqlite3.connect(cfg.cache_path)
    df = pd.read_sql("SELECT l1, cluster, count(*) AS goods, sum(sku_count) AS skus "
                     "FROM goods GROUP BY l1 ORDER BY goods DESC", con)
    con.close()
    return df


def load_l1(cfg, l1: str) -> pd.DataFrame:
    con = sqlite3.connect(cfg.cache_path)
    df = pd.read_sql("SELECT * FROM goods WHERE l1=?", con, params=(l1,))
    con.close()
    return df.fillna("")

"""匯出最新全量與瘦身版 Google Sheet 用 CSV。

包含：
1. output/momo/momo_全站品牌標記結果_GoogleSheet用.csv (完整 14 欄)
2. output/momo/momo_全站品牌標記結果_Part1.csv (前半 27.2 萬筆，完整欄位)
3. output/momo/momo_全站品牌標記結果_Part2.csv (後半 27.2 萬筆，完整欄位)
4. output/momo/momo_全站品牌標記結果_GoogleSheet用_單檔瘦身版.csv (11 欄，嚴格 < 100MB)
"""
from pathlib import Path
import sqlite3
import sys
import time
import pandas as pd

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = Path(__file__).resolve().parents[1]
DB_PATH = ROOT / "brand_rules.db"
CACHE_PATH = ROOT / "cache" / "momo_goods.db"
OUT_DIR = ROOT / "output" / "momo"
OUT_DIR.mkdir(parents=True, exist_ok=True)

t0 = time.time()
con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
con.execute(f"ATTACH DATABASE 'file:{CACHE_PATH}?mode=ro' AS goods_cache")

q = """
SELECT 
    g.item_id AS "品號",
    g.cluster AS "Cluster",
    g.l1 AS "L1",
    g.l2 AS "L2",
    g.l3 AS "L3",
    g.brand AS "品牌",
    g.title AS "商品名稱",
    r.[shp brand name1] AS "候選品牌1",
    r.[shp brand name2] AS "候選品牌2",
    r.[shp brand name3] AS "候選品牌3",
    CASE 
        WHEN r.[suggest brand name] = '建議品牌庫新增品牌' AND r.[新增品牌名稱] != '' 
            THEN r.[新增品牌名稱]
        ELSE r.[suggest brand name]
    END AS "TAGGING品牌",
    r.[信心指數] AS "信心度",
    CASE 
        WHEN CAST(r.[信心指數] AS INTEGER) < 75 THEN '是'
        ELSE '否'
    END AS "是否需要人工",
    r.[判斷路徑] AS "路徑"
FROM item_results r
JOIN goods_cache.goods g ON r.item_id = g.item_id
WHERE r.site='momo'
"""

print(f"正在從資料庫讀取 54.5 萬筆最新標記結果...")
df = pd.read_sql(q, con)
con.close()
print(f"資料讀取完成：共 {len(df):,} 筆（耗時 {time.time()-t0:.1f} 秒）")

# 1. 輸出全量版
p_full = OUT_DIR / "momo_全站品牌標記結果_GoogleSheet用.csv"
df.to_csv(p_full, index=False, encoding="utf-8-sig")
print(f"✔ 全量總表已更新：{p_full.name} ({p_full.stat().st_size / 1024 / 1024:.1f} MB)")

# 2. 輸出 Part1 & Part2
mid = len(df) // 2
p_part1 = OUT_DIR / "momo_全站品牌標記結果_Part1.csv"
p_part2 = OUT_DIR / "momo_全站品牌標記結果_Part2.csv"
df.iloc[:mid].to_csv(p_part1, index=False, encoding="utf-8-sig")
df.iloc[mid:].to_csv(p_part2, index=False, encoding="utf-8-sig")
print(f"✔ Part 1 已更新：{p_part1.name} ({p_part1.stat().st_size / 1024 / 1024:.1f} MB)")
print(f"✔ Part 2 已更新：{p_part2.name} ({p_part2.stat().st_size / 1024 / 1024:.1f} MB)")

# 3. 輸出單檔瘦身版（移除 Cluster、候選2/3，路徑改代碼，嚴格 < 100MB）
p_slim = OUT_DIR / "momo_全站品牌標記結果_GoogleSheet用_單檔瘦身版.csv"
slim = df.drop(columns=["Cluster", "候選品牌2", "候選品牌3"]).copy()
slim["路徑"] = slim["路徑"].astype(str).str.split().str[0]
slim.to_csv(p_slim, index=False, encoding="utf-8-sig")
print(f"✔ 單檔瘦身版已更新：{p_slim.name} ({p_slim.stat().st_size / 1024 / 1024:.1f} MB)")

print(f"\n全部匯出完成！總耗時 {time.time()-t0:.1f} 秒。")

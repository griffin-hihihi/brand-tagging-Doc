#!/usr/bin/env python3
"""競網品牌 Tagging —— 入口。

常用指令（在 brand_tagging 資料夾執行）：

    py run.py status                 看每個 L1 現在進行到哪
    py run.py tag                    跑 config.toml 裡 [run] active 列的 L1
    py run.py tag --l1 Beauty        只跑一個 L1
    py run.py tag --all              22 個 L1 全跑
    py run.py history --key B:dhc    查某個品牌字串的人工判斷歷程
    py run.py export --l1 Beauty     匯出品號層結果 CSV（之後上傳 Google Sheet 用）

流程：跑完 → 打開 review/<站點>/<L1>.xlsx 的「審核」頁 → 填【人工】判斷 → 存檔關閉 → 再跑一次同樣的指令。
你的判斷會進 brand_rules.db 的 decision_log（append-only，永不覆蓋），並且全站共用。
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pandas as pd

from brandtag import bundle, config, newbrands, pipeline, report, source, store, summary
from brandtag.index import load_pool
from brandtag.text import normalize


def _targets(cfg, args, all_l1) -> list:
    if args.all:
        return all_l1
    if args.l1:
        want = [x.strip() for x in ",".join(args.l1).split(",") if x.strip()]
        bad = [x for x in want if x not in all_l1]
        if bad:
            sys.exit(f"✖ 找不到 L1：{', '.join(bad)}\n   可用的有：{', '.join(all_l1)}")
        return want
    if not cfg.active:
        sys.exit("✖ config.toml 的 [run] active 是空的。請填要跑的 L1，或用 --l1 / --all 指定。")
    bad = [x for x in cfg.active if x not in all_l1]
    if bad:
        sys.exit(f"✖ config.toml 的 active 裡有不存在的 L1：{', '.join(bad)}")
    return list(cfg.active)


def cmd_status(cfg, args):
    source.ensure_cache(cfg, log=print)
    con = store.open_db(cfg.db_path)
    l1 = source.list_l1(cfg)
    t = report.status_table(cfg, con, l1)
    with pd.option_context("display.max_rows", None, "display.width", 200,
                           "display.max_colwidth", 30, "display.unicode.east_asian_width", True):
        print(t.to_string(index=False))
    n_rules = con.execute("SELECT count(*) FROM brand_rules").fetchone()[0]
    n_log = con.execute("SELECT count(*) FROM decision_log").fetchone()[0]
    n_ovr = con.execute("SELECT count(*) FROM item_overrides").fetchone()[0]
    n_al = con.execute("SELECT count(*) FROM brand_aliases").fetchone()[0]
    print(f"\n規則庫：品牌規則 {n_rules:,}｜單品覆蓋 {n_ovr:,}｜別名 {n_al:,}｜判斷歷程 {n_log:,} 筆")
    print(f"設定檔 active：{', '.join(cfg.active) or '(空)'}")
    con.close()


def cmd_tag(cfg, args):
    source.ensure_cache(cfg, force=args.rebuild_cache, log=print)
    all_l1 = list(source.list_l1(cfg)["l1"])
    targets = _targets(cfg, args, all_l1)
    con = store.open_db(cfg.db_path)
    pool = load_pool(cfg)
    print(f"要跑的 L1（{len(targets)}）：{', '.join(targets)}")
    done = []
    for l1 in targets:
        m = pipeline.run_l1(cfg, con, pool, l1, write_excel=not args.no_excel)
        if m:
            done.append(m)
    if len(done) > 1:
        print("\n━━━ 總計 ━━━")
        d = pd.DataFrame(done)
        tot = max(int(d["goods"].sum()), 1)
        print(f"{len(done)} 個 L1｜品號 {int(d['goods'].sum()):,}｜品牌字串 {int(d['keys'].sum()):,}")
        print(f"品牌庫命中 {d['pool_goods'].sum() / tot:.0%}｜建議新增 {d['new_goods'].sum() / tot:.0%}"
              f"｜No brand {d['nb_goods'].sum() / tot:.0%}")
        print(f"待審核 {int(d['pending_keys'].sum()):,} 個品牌字串"
              f"（{int(d['pending_goods'].sum()):,} 品號）")
    print(f"\n下一步：打開 {cfg.site_review_dir.relative_to(cfg.root)}\\<L1>.xlsx →「審核」頁 → "
          f"填【人工】判斷 → 存檔關閉 → 再跑一次同樣的指令")
    con.close()


def cmd_history(cfg, args):
    con = store.open_db(cfg.db_path)
    if args.key:
        rows = con.execute("SELECT * FROM decision_log WHERE rule_key=? ORDER BY log_id", (args.key,)).fetchall()
    else:
        rows = con.execute("SELECT * FROM decision_log ORDER BY log_id DESC LIMIT ?", (args.limit,)).fetchall()
        rows = list(reversed(rows))
    if not rows:
        print("（沒有紀錄）")
        return
    for r in rows:
        who = f"{r['actor']}@{r['src_file'] or '-'}"
        what = r["decision"]
        if r["decision"] == store.REVOKED:
            what = "撤銷"
        elif r["brand_name"]:
            what += f" → {r['brand_name']}" + (f" [{r['brand_id']}]" if r["brand_id"] is not None else "")
        sysinfo = f"（系統原建議 {r['sys_decision'] or '-'} / {r['sys_path'] or '-'} / {r['sys_conf'] or '-'}）"
        print(f"#{r['log_id']:<6} {r['logged_at']}  {r['level1'] or '-':<18} {r['rule_key']:<28} "
              f"{r['raw_brand'] or '':<22} {what}")
        print(f"        by {who} {sysinfo}" + (f"  備註：{r['note']}" if r["note"] else ""))
    con.close()


def cmd_export(cfg, args):
    con = store.open_db(cfg.db_path)
    q = "SELECT * FROM item_results WHERE site=?"
    p = [cfg.site]
    if args.l1:
        q += " AND level1 IN (" + ",".join("?" * len(args.l1)) + ")"
        p += list(args.l1)
    df = pd.read_sql(q, con, params=p)
    if df.empty:
        sys.exit("✖ item_results 沒有資料，請先跑 tag")
    out = Path(args.out) if args.out else cfg.site_output_dir / f"{cfg.site}_export.csv"
    out.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"✔ {len(df):,} 筆 → {out}")
    con.close()


def cmd_newbrands(cfg, args):
    """跨 L1 彙總所有待決策的品牌，給品牌庫負責人一次看完。"""
    con = store.open_db(cfg.db_path)
    g = newbrands.build(con, cfg.site)
    if g.empty:
        sys.exit("✖ 沒有待決策品牌，請先跑 tag")
    s = newbrands.summary(g)
    print(s.to_string(index=False))
    print()
    tot = g["總商品數"].sum()
    for n in (50, 100, 200, 500, 1000):
        if n <= len(g):
            cov = g["總商品數"].head(n).sum() / tot
            print(f"  決策前 {n:>5,} 個品牌 → 涵蓋 {cov:6.1%}（{g['總商品數'].head(n).sum():,} 件商品）")
    out = cfg.site_output_dir / f"_{cfg.site}_待決策品牌總表.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        s.to_excel(w, sheet_name="摘要", index=False)
        g.to_excel(w, sheet_name="待決策品牌", index=False)
        for name, df, widths in (("摘要", s, {}), ("待決策品牌", g, {"品牌字串": 28, "L1清單": 30,
                                                                "建議動作": 34, "最相似的品牌庫品牌": 30,
                                                                "範例商品名稱": 44, "系統判斷": 20})):
            ws = w.sheets[name]
            hf = w.book.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": True})
            for j, c in enumerate(df.columns):
                ws.write(0, j, c, hf)
                ws.set_column(j, j, widths.get(c, 14))
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), len(df.columns) - 1)
    print(f"\n✔ {len(g):,} 個待決策品牌 → {out.relative_to(cfg.root)}")
    con.close()


def cmd_curate(cfg, args):
    """把人工整理的品牌對照載入成別名與已查證的同實體 ID 關係。

    curated 低於使用者在審核檔逐筆確認的 human，但高於 natural / auto。
    同實體關係只開放配對內比較 adg，不會擴散到其他相同短縮寫。
    """
    import csv
    path = Path(args.file) if args.file else cfg.root / "curated_aliases.csv"
    if not path.exists():
        sys.exit(f"✖ 找不到 {path}")
    con = store.open_db(cfg.db_path)
    pool = load_pool(cfg)
    valid = set(pool.brand_id)
    pairs, links, bad = [], [], []
    with path.open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            cn, bid = (r.get("中文品牌名") or "").strip(), (r.get("brand_id") or "").strip()
            if not cn or not bid:
                continue
            if not bid.isdigit() or int(bid) not in valid:
                bad.append(f"{cn}：brand_id {bid} 不在品牌庫")
                continue
            pairs.append((normalize(cn), int(bid)))
            other = (r.get("同實體brand_id") or "").strip()
            if other:
                if not other.isdigit() or int(other) not in valid:
                    bad.append(f"{cn}：同實體brand_id {other} 不在品牌庫")
                elif int(other) != int(bid):
                    links.append((int(bid), int(other)))
    link_path = path.with_name("curated_entity_links.csv")
    if link_path.exists():
        with link_path.open(encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                b1 = (r.get("brand_id_a") or "").strip()
                b2 = (r.get("brand_id_b") or "").strip()
                if not b1.isdigit() or int(b1) not in valid or not b2.isdigit() or int(b2) not in valid:
                    bad.append(f"同實體關係 {b1} ↔ {b2}：brand_id 不在品牌庫")
                elif b1 != b2:
                    links.append((int(b1), int(b2)))
    for b in bad:
        print(f"   ⚠ {b}")
    n = store.add_aliases(con, pairs, "curated")
    n_links = store.add_entity_links(con, links, "curated")
    print(f"✔ 讀入 {len(pairs)} 筆對照，新增/升級 {n} 筆別名（來源 curated）")
    print(f"   同實體關係 {len(links)} 筆，新增 {n_links} 筆；關係內依 adg 選 canonical ID")
    print("   下一步：py run.py tag --all  讓新別名生效")
    con.close()


def cmd_summary(cfg, args):
    """產出給主管看的摘要檔，可直接上傳 Google Sheet。"""
    con = store.open_db(cfg.db_path)
    tables = summary.build(con, cfg.site)
    nb = newbrands.build(con, cfg.site)
    top = nb.head(200)[["排名", "品牌字串", "總商品數", "累積覆蓋", "出現在幾個L1",
                        "建議動作", "最相似的品牌庫品牌"]]
    out = cfg.site_output_dir / f"_{cfg.site}_主管摘要.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    summary.write(out, tables, top)
    print(tables["總覽"].to_string(index=False))
    print(f"\n✔ 摘要檔 → {out.relative_to(cfg.root)}")
    print("   4 個分頁：總覽 / 各類別 / 人工投入與回報 / 待決策品牌Top200")
    print("   檔案小，可直接上傳 Google Sheet")
    con.close()


def cmd_bundle_export(cfg, args):
    """將全站 22 個品類的審核表打包去重成單一總表，供 Temp 集中審核。"""
    out = Path(args.out) if args.out else None
    target = bundle.export_bundle(cfg, out_path=out, log=print)
    print(f"\n[OK] 審核總表產出完成：{target.name}")
    print("   可直接上傳至 Google Sheet 供 Temp 團隊填寫「【人工】判斷」欄位。")
    print("   填寫完畢下載 Excel 後，執行：py run.py bundle-import 即可回灌規則庫。")


def cmd_bundle_import(cfg, args):
    """將 Temp 填寫完成的審核總表讀回，寫入 brand_rules.db 規則庫。"""
    in_path = Path(args.file) if args.file else None
    n_imp, n_agree = bundle.import_bundle(cfg, in_path=in_path, log=print)
    print(f"\n[OK] 成功讀入 {n_imp} 筆人工判斷（同意原建議 {n_agree} 筆，修正 {n_imp - n_agree} 筆）！")
    print("   下一步：執行 py run.py tag --all 即可讓全站 22 個品類立即生效。")


def main():
    ap = argparse.ArgumentParser(description="競網品牌 → shp 品牌庫 對應工具 v1.0",
                                 formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("-c", "--config", default="config.toml", help="設定檔（預設 config.toml）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("status", help="看每個 L1 的進度")
    s.set_defaults(fn=cmd_status)

    s = sub.add_parser("tag", help="跑 tagging（讀回人工判斷 → 比對 → 更新審核檔與結果檔）")
    s.add_argument("--l1", nargs="+", help="只跑這些 L1（可多個或用逗號分隔）")
    s.add_argument("--all", action="store_true", help="跑全部 L1")
    s.add_argument("--no-excel", action="store_true", help="不產 Excel，只更新資料庫")
    s.add_argument("--rebuild-cache", action="store_true", help="強制重建 raw 快取")
    s.set_defaults(fn=cmd_tag)

    s = sub.add_parser("history", help="查人工判斷歷程")
    s.add_argument("--key", help="品牌字串 key，例如 B:dhc")
    s.add_argument("--limit", type=int, default=30, help="沒給 key 時顯示最近幾筆")
    s.set_defaults(fn=cmd_history)

    s = sub.add_parser("newbrands", help="跨 L1 彙總待決策品牌，產出總表")
    s.set_defaults(fn=cmd_newbrands)

    s = sub.add_parser("bundle-export", aliases=["export-review"], help="打包全站待審品牌成單一總表（給 Temp/Google Sheet 審核）")
    s.add_argument("--out", help="輸出路徑")
    s.set_defaults(fn=cmd_bundle_export)

    s = sub.add_parser("bundle-import", aliases=["import-review"], help="讀回 Temp 審核完成的總表，寫入 brand_rules.db")
    s.add_argument("--file", help="輸入檔案路徑（預設 review/<site>/_<site>_全站審核總表_Temp用.xlsx）")
    s.set_defaults(fn=cmd_bundle_import)

    s = sub.add_parser("curate", help="載入人工整理的中英文品牌對照表")
    s.add_argument("--file", help="對照表路徑（預設 curated_aliases.csv）")
    s.set_defaults(fn=cmd_curate)

    s = sub.add_parser("summary", help="產出給主管看的摘要檔（可上傳 Google Sheet）")
    s.set_defaults(fn=cmd_summary)

    s = sub.add_parser("export", help="匯出品號層結果 CSV")
    s.add_argument("--l1", nargs="+", help="只匯出這些 L1")
    s.add_argument("--out", help="輸出路徑")
    s.set_defaults(fn=cmd_export)

    args = ap.parse_args()
    cfg = config.load(args.config)
    args.fn(cfg, args)


if __name__ == "__main__":
    main()

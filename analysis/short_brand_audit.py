"""Read-only audit of short Latin brand matches; run with py -B analysis/short_brand_audit.py.

Writes evidence under analysis/short_brand_audit only. Does not import review decisions,
open the store through open_db(), or mutate the production database.
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from brandtag import config
from brandtag.index import BrandIndex, load_pool
from brandtag.text import clean_raw, extract_bracket, normalize, split_parts


def main():
    cfg = config.load(ROOT / 'config.toml')
    out = ROOT / 'analysis' / 'short_brand_audit'
    out.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(cfg.db_path.as_uri() + '?mode=ro', uri=True)
    con.execute('BEGIN')
    sql = '''SELECT site, level1, item_id, raw_brand, title,
        "suggest brand name" AS matched_name, "suggest brand id" AS matched_id,
        "信心指數" AS confidence, 判斷路徑 AS path, tagged_at
        FROM item_results WHERE site = ?'''
    items = pd.read_sql_query(sql, con, params=(cfg.site,)).fillna('')
    aliases = con.execute('SELECT alias,brand_id,source FROM brand_aliases').fetchall()
    runs = pd.read_sql_query('SELECT * FROM run_log WHERE site=? ORDER BY run_id', con, params=(cfg.site,))
    con.close()
    pool = load_pool(cfg)
    bi = BrandIndex(pool, aliases, cfg.pool_col_cat, cfg.th)
    source_map = {(a, b): s for a, b, s in aliases}
    raw_map = {s: clean_raw(s) for s in items.raw_brand.unique()}
    items['source_name'] = items.raw_brand.map(raw_map)
    empty = items.source_name.eq('')
    items['source_field'] = '品牌欄'
    items.loc[empty, 'source_name'] = items.loc[empty, 'title'].map(extract_bracket)
    items.loc[empty, 'source_field'] = '標題【】'
    parts = {s: split_parts(s) for s in items.source_name.unique()}
    items['source_norm'] = items.source_name.map(lambda s: parts[s][0])
    items['latin'] = items.source_name.map(lambda s: parts[s][1])
    items['chinese'] = items.source_name.map(lambda s: parts[s][2])
    # Broad screening, not a declaration that these mappings are wrong.
    sub = items[items.latin.str.len().between(1, 3) & items.chinese.ne('')].copy()
    sub['matched_id_num'] = pd.to_numeric(sub.matched_id, errors='coerce').fillna(-1).astype(int)
    sub['pool_latin'] = sub.matched_id_num.map(lambda b: bi.parts.get(b, ('', '', ''))[1])
    sub['pool_chinese'] = sub.matched_id_num.map(lambda b: bi.parts.get(b, ('', '', ''))[2])
    sub['pool_category'] = sub.matched_id_num.map(bi.cat).fillna('')
    sub['alias_source'] = [source_map.get((n, b), '') or source_map.get((c, b), '')
                           for n, c, b in zip(sub.source_norm, sub.chinese, sub.matched_id_num)]
    sub['latin_length'] = sub.latin.str.len()
    sub['risk'] = (sub.matched_id_num.isin(bi.name) & ~sub.matched_id_num.isin([0, 99999999])
                   & sub.latin.eq(sub.pool_latin) & sub.chinese.ne(sub.pool_chinese)
                   & ~sub.alias_source.isin(['human', 'curated']))
    risk = sub[sub.risk].copy()
    group_cols = ['source_norm', 'matched_id_num']
    grouped = risk.groupby(group_cols).agg(
        source_name=('source_name', 'first'), source_field=('source_field', lambda s: '|'.join(sorted(set(s)))),
        latin=('latin', 'first'), chinese=('chinese', 'first'), latin_length=('latin_length', 'first'),
        matched_name=('matched_name', 'first'), pool_category=('pool_category', 'first'),
        pool_chinese=('pool_chinese', 'first'), alias_source=('alias_source', 'first'),
        goods=('item_id', 'size'), categories=('level1', lambda s: '|'.join(sorted(set(s)))),
        title=('title', 'first'), confidence_min=('confidence', lambda s: pd.to_numeric(s).min()),
        confidence_max=('confidence', lambda s: pd.to_numeric(s).max()),
    ).reset_index().sort_values('goods', ascending=False)
    grouped.to_csv(out / 'risk_groups.csv', index=False, encoding='utf-8-sig')
    risk.to_csv(out / 'risk_items.csv', index=False, encoding='utf-8-sig')
    sub.to_csv(out / 'all_short_bilingual_items.csv', index=False, encoding='utf-8-sig')
    summary = {
        'site': cfg.site, 'total_goods': len(items), 'unique_items': int(items.item_id.nunique()),
        'categories': int(items.level1.nunique()), 'pool_ids': len(pool),
        'result_min': items.tagged_at.min(), 'result_max': items.tagged_at.max(),
        'risk_definition': '有中文及1-3字元英數、現有結果配到英數相同但中文不相同的品牌，且無human/curated別名支持；是待核風險，不是已證實錯配。',
        'bilingual_short_goods': len(sub),
        'risk_goods': len(risk), 'risk_source_names': int(risk.source_norm.nunique()),
        'risk_categories': int(risk.level1.nunique()),
        'by_length': {str(n): {'source_names': int(d.source_norm.nunique()), 'goods': len(d),
                              'auto_alias_goods': int(d.alias_source.eq('auto').sum())}
                      for n, d in risk.groupby('latin_length')},
        'by_category': risk.groupby('level1').size().to_dict(),
        'latest_runs': runs.drop_duplicates('level1', keep='last')[['level1', 'ran_at']].to_dict('records'),
    }
    (out / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')
    (out / 'query.sql').write_text(sql, encoding='utf-8')
    print(json.dumps({k:v for k,v in summary.items() if k not in ['latest_runs','by_category']}, ensure_ascii=False, indent=2))
    print(grouped.to_string(index=False, columns=['source_name','latin_length','matched_name','pool_category','alias_source','goods']))
    print('AUTO ALIASES FOR RISK GROUPS')
    for r in grouped.itertuples():
        if r.alias_source == 'auto':
            print(r.source_name, r.matched_name, [(a,b,s) for a,b,s in aliases if b == r.matched_id_num])


if __name__ == '__main__':
    main()

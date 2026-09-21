"""Compare repair scenarios in memory, preserving production code, aliases and results.

This is a bounded research experiment, not a replacement tagging engine. A review
gate below is applied to the audited risk population, not to a claimed error set.
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
from brandtag.engine import tag
from brandtag.index import BrandIndex, load_pool
from brandtag.text import normalize, split_parts

OUT = ROOT / 'analysis' / 'short_brand_audit'

# Research proposals: none of these entries are written to curated_aliases or the DB.
PROPOSALS = [
    ('GB 綠鐘', 1124530, 'GREEN BELL', '222件短縮寫配到母嬰gb；產品型號與Green Bell官網相符',
     'https://greenbell.ne.jp/chntd/product-list/kitchen/'),
    ('HH 草本新淨界', 1111148, 'Herb & Health', '官方確認完整英文名稱；H&H亦為官方用名，原ID不能只憑名稱斷言錯誤',
     'https://page.line.me/mbx3422u/signboard/554079114850404'),
    ('JS 婕洛妮絲', 3472328, 'jealousness', '官方確認Jealousness；現有JS短名ID缺乏實體支持',
     'https://page.line.me/hzz8479e'),
    ('GP 超霸', 1146524, 'GP Batteries', '電池品牌配到鞋類G.P；GP電池與GP鞋業各有官網',
     'https://tw.gpbatteries.com/pages/corporate-information'),
    ('KUM 熊野', 1136616, 'Kumano 熊野油脂', '熊野油脂與德國KUM為不同品牌；庫內已有完整中文名',
     'https://www.kumanoyushi.co.jp/company/'),
    ('H&R 安室家', 1111072, 'H&R 安室家', '完整自然命中卻被同縮寫汽車類高ADG覆蓋；本例可直接由庫內名稱修正',
     'https://www.h-r.com/en/company/'),
]


def main():
    cfg = config.load(ROOT / 'config.toml')
    pool = load_pool(cfg)
    con = sqlite3.connect(cfg.db_path.as_uri() + '?mode=ro', uri=True)
    aliases = con.execute('SELECT alias,brand_id,source FROM brand_aliases').fetchall()
    con.close()
    risk = pd.read_csv(OUT / 'risk_items.csv').fillna('')
    groups = pd.read_csv(OUT / 'risk_groups.csv').fillna('')
    all_short = pd.read_csv(OUT / 'all_short_bilingual_items.csv').fillna('')
    proposed_ids = {normalize(s): bid for s, bid, _, _, _ in PROPOSALS}
    rows, corrected_aliases, quarantine = [], [], []
    for name, bid, label, why, url in PROPOSALS:
        assert pool.loc[pool.brand_id.eq(bid), 'brand_name'].iloc[0] == label
        n, lat, cjk = split_parts(name)
        matching = risk[risk.source_norm.eq(n)]
        rows.append({'來源品牌': name, '目前品牌': '|'.join(matching.matched_name.unique()),
                     '目前ID': '|'.join(map(str, matching.matched_id_num.unique())),
                     '建議品牌': label, '建議ID': bid, '品號數': len(matching),
                     '類別': '|'.join(sorted(matching.level1.unique())), '判讀': why, '證據': url})
        # Full forms and Chinese forms only; never globally alias bare GB, HH, JS, GP, KUM, H&R.
        for a in {n, cjk}:
            corrected_aliases.append((a, bid, 'curated'))
        for a, old_bid, src in aliases:
            if src == 'auto' and a in {n, cjk} and old_bid != bid:
                quarantine.append((a, old_bid, src))
    qset = set(quarantine)
    indices = {
        '現有': BrandIndex(pool, aliases, cfg.pool_col_cat, cfg.th),
        '只加對照': BrandIndex(pool, aliases + corrected_aliases, cfg.pool_col_cat, cfg.th),
        '隔離衝突auto再加對照': BrandIndex(pool, [a for a in aliases if a not in qset] + corrected_aliases,
                                      cfg.pool_col_cat, cfg.th),
    }
    scenarios = []
    for name, bid, label, _, _ in PROPOSALS:
        for scenario, bi in indices.items():
            r = tag(name, '', None, bi, False)
            scenarios.append({'來源品牌': name, '實驗': scenario, '輸出品牌': r['suggest brand name'],
                              '輸出ID': r['suggest brand id'], '信心': r['信心指數'],
                              '路徑': r['判斷路徑'], '符合提案ID': str(r['suggest brand id']) == str(bid)})

    # A concrete bounded prototype for the two remaining safeguards:
    # 1. A complete short bilingual natural match wins before ADG duplicate resolution.
    # 2. Other audited unsupported short bilingual matches remain candidates at <=74.
    # No unknown brand is declared wrong/new/No brand by this gate.
    risk_names = set(groups.source_norm)
    bi = indices['隔離衝突auto再加對照']
    def prototype(name, title, cat):
        r = tag(name, title, cat, bi, cat in set(bi.cat.values()))
        n, lat, cjk = split_parts(name)
        if n in proposed_ids:
            # Experimental policy: a source-backed full-form mapping outranks
            # unverified short-form alternatives. This is NOT a human DB decision.
            b = proposed_ids[n]
            r.update({'suggest brand name': bi.name[b], 'suggest brand id': str(b),
                      '信心指數': 90 if cat and bi.cat.get(b) != cat else 98,
                      '判斷說明': '研究提案：有來源支持的完整名稱對照優先；分數為提案分数，非實測準確率'})
        if cjk and 1 <= len(lat) <= 3:
            exact = [b for b in bi.by_norm.get(n, set()) if b not in (0, 99999999)]
            if exact:
                # Ambiguous natural full names need review, not category-blind ADG selection.
                if len(exact) == 1:
                    b = exact[0]
                    r.update({'suggest brand name': bi.name[b], 'suggest brand id': str(b),
                              '信心指數': min(98, 90 if cat and bi.cat.get(b) != cat else 98),
                              '判斷說明': '研究提案：完整雙語名稱優先於缺少中文的短縮寫候選'})
                else:
                    r['信心指數'] = min(74, r['信心指數'])
        if n in risk_names and n not in proposed_ids:
            r['信心指數'] = min(74, r['信心指數'])
            r['判斷說明'] += '；研究提案：短英數命中但中文缺乏獨立支持，候選保留待確認'
        return r

    comparisons = []
    for (name, cat), d in risk.groupby(['source_name', 'level1']):
        r = prototype(name, str(d.title.iloc[0]), cat)
        comparisons.append({'來源品牌': name, '類別': cat, '品號數': len(d),
                            '現有品牌': d.matched_name.iloc[0], '提案品牌': r['suggest brand name'],
                            '提案ID': r['suggest brand id'], '提案信心': r['信心指數'],
                            '狀態': '有來源支持的定向修正' if normalize(name) in proposed_ids else '候選保留待確認'})
    comparisons = pd.DataFrame(comparisons)
    unresolved = comparisons[comparisons['狀態'].eq('候選保留待確認')]
    assert unresolved['提案信心'].le(74).all()
    # The broad natural-full-name safeguard also changes choices outside the
    # six researched examples. Retain these at <=74 and disclose the impact;
    # the experiment does not establish which IDs are true duplicate entities.
    additional_changes = unresolved[~unresolved['現有品牌'].eq(unresolved['提案品牌'])].to_dict('records')
    controls = []
    for name in ['3M', 'LG', 'HP', 'DHC', 'VT', 'SK2', 'FJ 豐傑生醫', 'BVLGARI 寶格麗', 'MUJI 無印良品']:
        old = tag(name, '', None, indices['現有'], False)
        new = prototype(name, '', None)
        assert old['suggest brand id'] == new['suggest brand id'], (name, old, new)
        controls.append({'品牌': name, '原結果': old['suggest brand name'], '提案結果': new['suggest brand name'],
                         '原信心': old['信心指數'], '提案信心': new['信心指數']})
    for name, bid, _, _, _ in PROPOSALS:
        r = prototype(name, '', None)
        assert r['suggest brand id'] == str(bid), (name, r)

    proposals = pd.DataFrame(rows)
    scenarios = pd.DataFrame(scenarios)
    quarantine_df = pd.DataFrame(quarantine, columns=['alias', 'brand_id', 'source'])
    proposals.to_csv(OUT / 'proposed_mappings.csv', index=False, encoding='utf-8-sig')
    scenarios.to_csv(OUT / 'scenario_comparison.csv', index=False, encoding='utf-8-sig')
    quarantine_df.to_csv(OUT / 'proposed_alias_quarantine.csv', index=False, encoding='utf-8-sig')
    comparisons.to_csv(OUT / 'proposed_review_gate.csv', index=False, encoding='utf-8-sig')

    summary = json.loads((OUT / 'summary.json').read_text(encoding='utf-8'))
    summary['priority_proposals'] = {'source_names': len(proposals), 'goods': int(proposals['品號數'].sum()),
        'remaining_unverified_source_names': int(groups[~groups.source_norm.isin(proposed_ids)].source_norm.nunique()),
        'remaining_unverified_goods': int(groups.loc[~groups.source_norm.isin(proposed_ids), 'goods'].sum()),
        'quarantined_aliases_in_memory_only': quarantine,
        'note': '6組為有依據的優先修正提案，不等於已證實6個原ID均屬不同品牌；HH原ID可能也是同品牌重複登錄。'}
    summary['prototype_checks'] = {'proposed_ids_passed': 6, 'control_ids_unchanged': len(controls),
        'additional_unverified_selection_changes': additional_changes,
        'scope': '針對已篩出的風險字串與9個對照字串，在記憶體做定向修正及待審門檻實驗；不是全站新引擎回歸測試。'}
    summary['risk_high_confidence_goods'] = int(pd.to_numeric(risk.confidence).ge(90).sum())
    summary['risk_currently_pending_goods'] = int(pd.to_numeric(risk.confidence).lt(75).sum())
    (OUT / 'summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding='utf-8')

    localized = groups.rename(columns={'source_name':'來源品牌','matched_name':'目前品牌','matched_id_num':'目前ID',
        'goods':'品號數','latin_length':'英數長度','alias_source':'中文別名來源','categories':'商品類別',
        'pool_category':'品牌庫類別','title':'範例商品','source_field':'來源欄位',
        'confidence_min':'最低信心','confidence_max':'最高信心'})
    localized['研究狀態'] = localized.source_norm.map(lambda n:'優先修正提案' if n in proposed_ids else '尚未逐一查證，不能視為錯配')
    with pd.ExcelWriter(OUT / '短英文品牌配對研究.xlsx', engine='xlsxwriter',
                        engine_kwargs={'options':{'strings_to_urls':False}}) as w:
        sheets = {'優先修正提案':proposals, '全部風險字串':localized,
                  '僅2字元風險':localized[localized['英數長度'].le(2)],
                  '風險品號明細':risk, '修正方式比較':scenarios,
                  '提案待審門檻實驗':comparisons, '待隔離auto別名':quarantine_df,
                  '對照案例':pd.DataFrame(controls)}
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
            ws=w.sheets[name]; ws.freeze_panes(1,0); ws.autofilter(0,0,len(df),len(df.columns)-1)
            ws.set_column(0,len(df.columns)-1,22)
    # Runnable companion: delegates to the exact checked-in read-only scripts.
    nb = {'nbformat':4,'nbformat_minor':5,'metadata':{'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'}},
          'cells':[
              {'cell_type':'markdown','metadata':{},'source':['# 短英文品牌配對研究\n',
                  '資料庫只讀。89組/2,987件為證據不足的風險篩查，不是錯誤率。提案僅在記憶體比較，不寫入正式規則或結果。\n',
                  '研究範圍：來源有中文且英數長度1–3、結果英數相同但中文不相同、沒有human/curated別名支持。數字隨來源更新會改變。'],},
              {'cell_type':'code','execution_count':None,'metadata':{},'outputs':[],
               'source':['from pathlib import Path\n','import runpy\n',
                         "root = next(p for p in [Path.cwd(), *Path.cwd().parents] if (p / 'brandtag').is_dir())\n",
                         "runpy.run_path(str(root / 'analysis/short_brand_audit.py'), run_name='__main__')\n"]},
              {'cell_type':'code','execution_count':None,'metadata':{},'outputs':[],
               'source':["runpy.run_path(str(root / 'analysis/short_brand_proposal.py'), run_name='__main__')\n"]},
          ]}
    (OUT / '研究重現.ipynb').write_text(json.dumps(nb,ensure_ascii=False,indent=2),encoding='utf-8')
    print(proposals.to_string(index=False))
    print(scenarios.to_string(index=False))
    print(json.dumps(summary['priority_proposals'],ensure_ascii=False,indent=2))
    print(summary['prototype_checks'])


if __name__ == '__main__':
    main()

import sqlite3, io
import pandas as pd
from brandtag import const
from brandtag.const import TYPE_POOL, TYPE_NEW, TYPE_NB

c = sqlite3.connect('brand_rules.db')
d = pd.read_sql('SELECT level1,"suggest brand name" s,信心程度 k,是否待人工判斷 p,判斷路徑 r FROM item_results', c)
d['t'] = d['s'].map(lambda x: x if x in (TYPE_NEW, TYPE_NB) else TYPE_POOL)
tot = len(d)
out = io.StringIO()
def W(*a): out.write('\t'.join(str(x) for x in a) + '\n')

auto = (d['p'] != '是'); pend = ~auto
W('momo 商品品牌對應 — 結果總表'); W('更新日期', '2026/09/21'); W()
W('【總覽】'); W('項目', '數量', '占比')
for lbl, m in [('商品總數（品號）', None), ('系統自動完成', auto), ('需人工確認', pend),
               ('對應到品牌庫既有品牌', d['t'].eq(TYPE_POOL)), ('品牌庫沒有、建議新增', d['t'].eq(TYPE_NEW)),
               ('判定為無品牌（白牌）', d['t'].eq(TYPE_NB)), ('高信心', d['k'].eq('高')),
               ('中信心', d['k'].eq('中')), ('低信心（必須人工看）', d['k'].eq('低'))]:
    W(lbl, f'{tot:,}' if m is None else f'{m.sum():,}', '' if m is None else f'{m.mean():.1%}')
W()
W('【各品類結果】')
W('商品類別', '商品數', '自動完成率', '對應到品牌庫', '建議新增', '無品牌', '高信心', '待人工件數')
g = d.groupby('level1').apply(lambda x: pd.Series({
    'n': len(x), 'auto': (x['p'] != '是').mean(), 'pool': (x['t'] == TYPE_POOL).mean(),
    'new': (x['t'] == TYPE_NEW).mean(), 'nob': (x['t'] == TYPE_NB).mean(),
    'hi': (x['k'] == '高').mean(), 'pend': (x['p'] == '是').sum()}), include_groups=False)
for l1, r in g.sort_values('n', ascending=False).iterrows():
    W(l1, f"{int(r['n']):,}", f"{r['auto']:.0%}", f"{r['pool']:.0%}", f"{r['new']:.0%}",
      f"{r['nob']:.0%}", f"{r['hi']:.0%}", f"{int(r['pend']):,}")
W()
W('【判斷路徑分布】'); W('代碼', '名稱', '信心', '什麼情況', '商品數', '占比')
d['code'] = d['r'].str[:2]
for k, v in d.groupby('code').size().sort_values(ascending=False).items():
    p = const.PATHS.get(k, ('', '', '', ''))
    W(k, p[0], p[1], p[2], f'{v:,}', f'{v/tot:.1%}')
W()
W('【人工確認的投入與回報】'); W('確認品牌數', '可解決商品數', '全站覆蓋率', '估計投入')
for n, cov, cvg, t in [(50, 37907, '83.1%', '約 1 小時'), (100, 52418, '85.7%', '約 2 小時'),
                       (200, 67699, '88.5%', '約半天'), (500, 87846, '92.2%', '約 1 天'),
                       (1000, 100893, '94.6%', '約 2 天')]:
    W(f'{n:,}', f'{cov:,}', cvg, t)
open('品牌對應_結果總表_可貼上.tsv', 'w', encoding='utf-8-sig').write(out.getvalue())
print(out.getvalue())
c.close()

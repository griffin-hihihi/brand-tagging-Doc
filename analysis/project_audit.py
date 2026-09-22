"""Read-only audit receipt: py analysis/project_audit.py [before|after]."""
from pathlib import Path
import json
import sqlite3
import sys

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'output' / 'audit_20260922'
OUT.mkdir(parents=True, exist_ok=True)
con = sqlite3.connect(f'file:{ROOT / "brand_rules.db"}?mode=ro', uri=True)
con.execute('ATTACH DATABASE ? AS src', (f'file:{ROOT / "cache/momo_goods.db"}?mode=ro',))
queries = {
    'integrity': 'PRAGMA quick_check',
    'coverage': 'SELECT count(*),count(distinct item_id),count(distinct level1),min(tagged_at),max(tagged_at) FROM item_results',
    'source_rows': 'SELECT count(*) FROM src.goods',
    'missing_results': 'SELECT count(*) FROM src.goods WHERE item_id NOT IN (SELECT item_id FROM item_results WHERE site="momo")',
    'orphan_results': 'SELECT count(*) FROM item_results r LEFT JOIN src.goods g ON r.item_id=g.item_id WHERE g.item_id IS NULL',
    'human_decisions': 'SELECT count(*) FROM decision_log',
    'paths': 'SELECT [判斷路徑],count(*) FROM item_results GROUP BY 1',
    'iphone_brands': 'SELECT [suggest brand name],count(*) FROM item_results WHERE lower([suggest brand name]) LIKE "%iphone%" GROUP BY 1',
    'known_description_new': '''SELECT [新增品牌名稱],count(*) FROM item_results
        WHERE [suggest brand name]='建議品牌庫新增品牌' AND [新增品牌名稱] IN
        ('買一送一','超抗刮','洗衣機過濾網','運動相機通用','製造','防摔專家') GROUP BY 1''',
    'new_bracket_examples': '''SELECT [新增品牌名稱],count(*),min(title) FROM item_results
        WHERE [判斷路徑] LIKE '2.3%' GROUP BY 1 ORDER BY 2 DESC LIMIT 100''',
}
data = {name: {'query': sql, 'rows': con.execute(sql).fetchall()} for name, sql in queries.items()}
con.close()
stage = sys.argv[1] if len(sys.argv) > 1 else 'after'
target = OUT / f'{stage}.json'
target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
print(target)
for name in ('integrity', 'coverage', 'missing_results', 'orphan_results', 'iphone_brands', 'known_description_new'):
    print(name, data[name]['rows'])

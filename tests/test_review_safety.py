import sqlite3
import uuid
import unittest
from pathlib import Path
from types import SimpleNamespace

import pandas as pd

from brandtag import bundle, store
from brandtag.const import H_PICK, TYPE_NB


class ReviewSafetyTests(unittest.TestCase):
    def test_bundle_exposes_cross_category_conflict_and_newbrand_action(self):
        root = Path(__file__).resolve().parents[1] / 'output' / ('test_review_' + uuid.uuid4().hex)
        root.mkdir(parents=True)
        try:
            for l1, bid in [('Beauty', '10'), ('Health', '20')]:
                pd.DataFrame([{'key': 'B:test', '狀態': '④ 自動', '商品數': 3,
                               'suggest brand name': 'Test', 'suggest brand id': bid}]).to_excel(
                                   root / f'{l1}.xlsx', sheet_name='審核', index=False)
            target = bundle.export_bundle(SimpleNamespace(site_review_dir=root, site='momo'), log=lambda *a: None)
            row = pd.read_excel(target).iloc[0]
            self.assertEqual(row['跨類目衝突'], '是')
            self.assertEqual(row['商品數'], 6)
            self.assertTrue(row['狀態'].startswith('①'))
            self.assertIn('Beauty', row['各類目建議'])
            self.assertIn('Health', row['各類目建議'])
        finally:
            for name in ('Beauty.xlsx', 'Health.xlsx', '_momo_全站審核總表_Temp用.xlsx'):
                (root / name).unlink(missing_ok=True)
            root.rmdir()

    def test_stale_review_cannot_overwrite_newer_decision(self):
        con = sqlite3.connect(':memory:')
        con.row_factory = sqlite3.Row
        con.executescript(store.SCHEMA)
        entry = {'scope': 'brand', 'rule_key': 'B:test', 'decision': TYPE_NB,
                 'brand_id': 0, 'brand_name': TYPE_NB, 'note': '', 'nobrand_reason': '人工確認',
                 'expected_log_id': 0}
        self.assertEqual(store.record(con, [entry, entry], 'first.xlsx', 'momo', 'ALL')[0], 1)
        self.assertEqual(store.record(con, [entry], 'same.xlsx', 'momo', 'ALL')[0], 0)
        with self.assertRaisesRegex(ValueError, '更新的人工判斷'):
            store.record(con, [{**entry, 'note': 'outdated edit'}], 'stale.xlsx', 'momo', 'ALL')
        self.assertEqual(con.execute('SELECT count(*) FROM decision_log').fetchone()[0], 1)
        con.close()

    def test_conflicting_duplicate_keys_are_atomic(self):
        con = sqlite3.connect(':memory:')
        con.row_factory = sqlite3.Row
        con.executescript(store.SCHEMA)
        entry = {'scope': 'brand', 'rule_key': 'B:test', 'decision': TYPE_NB, 'note': ''}
        with self.assertRaisesRegex(ValueError, '互相衝突'):
            store.record(con, [entry, {**entry, 'note': 'different'}], 'bad.xlsx', 'momo', 'ALL')
        self.assertEqual(con.execute('SELECT count(*) FROM decision_log').fetchone()[0], 0)
        con.close()

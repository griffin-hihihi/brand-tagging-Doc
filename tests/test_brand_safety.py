from __future__ import annotations

import unittest

from brandtag.engine import tag
from brandtag.index import BrandIndex
from brandtag.text import brand_key, extract_bracket
from brandtag.const import TYPE_NEW, TYPE_NB, CONF_COL
from test_short_brand_conflicts import pool, TH


class BrandSafetyTests(unittest.TestCase):
    def setUp(self):
        self.bi = BrandIndex(pool(
            (10, 'Apple', 'Mobile & Gadgets', 1000),
            (11, 'iPhone Apple', 'Mobile & Gadgets', 2000),
            (12, 'iphone', 'Mobile & Gadgets', 9999),
            (13, 'APPLE HOUSE 蘋果屋', 'Books', 999999),
            (14, 'RHINOSHIELD 犀牛盾', 'Mobile & Gadgets', 100),
            (15, 'Samsung', 'Mobile & Gadgets', 900)), [], 'level1_category', TH)

    def test_apple_canonical_adg_excludes_product_and_unrelated_apple(self):
        for value in ('Apple', 'iphone', 'Apple iPhone 15', 'iPhone Apple'):
            with self.subTest(value=value):
                self.assertEqual(tag(value, '', 'Mobile & Gadgets', self.bi, True)['suggest brand id'], '11')

    def test_descriptions_are_not_new_brands(self):
        for value in ('買一送一', '超抗刮', '洗衣機過濾網', '台灣製造', '運動相機通用'):
            with self.subTest(value=value):
                r = tag('', f'【{value}】商品', 'Mobile & Gadgets', self.bi, True)
                self.assertNotEqual(r['suggest brand name'], TYPE_NEW)
                self.assertTrue(brand_key('', f'【{value}】商品', '123').startswith('I:'))

    def test_empty_and_square_brackets(self):
        self.assertEqual(extract_bracket('【】[Samsung] 手機'), 'Samsung')
        self.assertEqual(tag('', '【】Samsung 手機', 'Mobile & Gadgets', self.bi, True)['suggest brand id'], '15')

    def test_title_after_80_characters(self):
        r = tag('', '一般商品描述' * 20 + ' Samsung 手機', 'Mobile & Gadgets', self.bi, True)
        self.assertEqual(r['suggest brand id'], '15')

    def test_known_third_party_accessory(self):
        r = tag('', '【RHINOSHIELD 犀牛盾】iPhone 15 保護殼', 'Mobile & Gadgets', self.bi, True)
        self.assertEqual(r['suggest brand id'], '14')

    def test_host_accessory_is_pending_and_item_scoped(self):
        r = tag('', '【Apple】iPhone 15 保護殼', 'Mobile & Gadgets', self.bi, True)
        self.assertEqual(r['suggest brand name'], TYPE_NB)
        self.assertLess(r[CONF_COL], 75)
        self.assertEqual(brand_key('', '【Apple】iPhone 15 保護殼', '123'), 'I:123')
        self.assertEqual(brand_key('', '【Apple】iPhone 15 手機', '456'), 'T:apple')
        r = tag('iphone', 'iPhone 15 保護殼', 'Mobile & Gadgets', self.bi, True)
        self.assertEqual(r['suggest brand name'], TYPE_NB)
        self.assertLess(r[CONF_COL], 75)

    def test_unrelated_brands_do_not_use_last_word_as_subbrand(self):
        r = tag('Apple Samsung', '', 'Mobile & Gadgets', self.bi, True)
        self.assertLess(r[CONF_COL], 75)


if __name__ == '__main__':
    unittest.main()

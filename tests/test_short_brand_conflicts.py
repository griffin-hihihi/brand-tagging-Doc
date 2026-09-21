from __future__ import annotations

import sqlite3
import unittest
from unittest.mock import patch

import pandas as pd

from brandtag.const import CONF_COL, TYPE_POOL, TYPE_NEW
from brandtag.engine import tag
from brandtag.index import BrandIndex
from brandtag.pipeline import _learn_auto_aliases
from brandtag import store


TH = {"auto": 95, "suggest": 75, "tier_gap": 5, "near_ratio": 85,
      "title_factor": 0.95, "fuzzy_cutoff": 88}


def pool(*rows):
    return pd.DataFrame(rows, columns=["brand_id", "brand_name", "level1_category", "adg"])


class ShortBrandConflictTests(unittest.TestCase):
    def test_curated_full_mapping_beats_old_auto_alias(self):
        p = pool((1, "gb", "Mother & Baby", 4209), (2, "GREEN BELL", "Beauty", 20497))
        aliases = [("綠鐘", 1, "auto"), ("gb綠鐘", 2, "curated")]
        bi = BrandIndex(p, aliases, "level1_category", TH)

        r = tag("GB 綠鐘", "", "Beauty", bi, True)

        self.assertEqual(r["suggest brand id"], "2")
        self.assertEqual(r[CONF_COL], 98)

    def test_unverified_short_latin_match_stays_pending(self):
        p = pool((1, "MSI", "Hardware & 3C", 1000))
        aliases = [("微星", 1, "auto")]
        bi = BrandIndex(p, aliases, "level1_category", TH)

        r = tag("MSI 微星", "", "Hardware & 3C", bi, True)

        self.assertEqual(r["suggest brand id"], "1")
        self.assertLess(r[CONF_COL], 75)
        self.assertIn("中文名稱未被獨立證實", r["判斷說明"])

    def test_curated_same_entity_uses_adg_for_hh_duplicate_ids(self):
        p = pool((1, "H&H", "Home & Living", 2392),
                 (2, "Herb & Health", "Beauty", 181621))
        aliases = [("hh草本新淨界", 2, "curated")]
        bi = BrandIndex(p, aliases, "level1_category", TH, [(1, 2, "curated")])

        r = tag("HH 草本新淨界", "", "Beauty", bi, True)

        self.assertTrue(bi.same_entity(1, 2))
        self.assertEqual(r["suggest brand id"], "2")
        self.assertTrue(r["判斷路徑"].startswith("1.4"))
        self.assertIn("依業績取 Herb & Health", r["判斷說明"])

        exact = tag("H&H", "", "Beauty", bi, True)
        self.assertEqual(exact["suggest brand id"], "2")
        self.assertTrue(exact["判斷路徑"].startswith("1.4"))

    def test_surface_exact_hh_does_not_inherit_h_and_h_entity_link(self):
        p = pool((1, "H&H", "Home & Living", 2392),
                 (2, "Herb & Health", "Beauty", 181621),
                 (3, "HH", "Mobile & Gadgets", 1815))
        bi = BrandIndex(p, [], "level1_category", TH, [(1, 2, "curated")])

        r = tag("HH", "", "Mobile & Gadgets", bi, True)

        self.assertEqual(r["suggest brand id"], "3")

    def test_exact_bilingual_brand_is_not_replaced_by_high_adg_short_name(self):
        p = pool((1, "H&R", "Motors", 100000), (2, "H&R 安室家", "Home & Living", 300))
        bi = BrandIndex(p, [], "level1_category", TH)

        r = tag("H&R 安室家", "", "Home & Living", bi, True)

        self.assertEqual(r["suggest brand id"], "2")
        self.assertEqual(r[CONF_COL], 98)
        self.assertFalse(bi.same_entity(1, 2))

    def test_long_shared_english_name_can_still_be_same_entity(self):
        p = pool((1, "Liese", "Beauty", 100), (2, "Liese 莉婕", "Beauty", 200))
        bi = BrandIndex(p, [], "level1_category", TH)
        self.assertTrue(bi.same_entity(1, 2))

    def test_short_bilingual_result_does_not_learn_auto_chinese_alias(self):
        p = pool((1, "MSI", "Hardware & 3C", 1000))
        bi = BrandIndex(p, [], "level1_category", TH)
        kres = pd.DataFrame({CONF_COL: [93], "_type": [TYPE_POOL], "_bid": [1]}, index=["k"])
        kinfo = pd.DataFrame({"raw": ["MSI 微星"]}, index=["k"])
        with patch("brandtag.pipeline.store.add_aliases") as add:
            self.assertEqual(_learn_auto_aliases(kres, kinfo, bi, object()), 0)
            add.assert_not_called()

    def test_long_bilingual_result_can_learn_chinese_alias(self):
        p = pool((1, "Liese", "Beauty", 1000))
        bi = BrandIndex(p, [], "level1_category", TH)
        kres = pd.DataFrame({CONF_COL: [98], "_type": [TYPE_POOL], "_bid": [1]}, index=["k"])
        kinfo = pd.DataFrame({"raw": ["Liese 莉婕"]}, index=["k"])
        with patch("brandtag.pipeline.store.add_aliases", return_value=1) as add:
            self.assertEqual(_learn_auto_aliases(kres, kinfo, bi, object()), 1)
            self.assertIn(("莉婕", 1), add.call_args.args[1])

    def test_alias_source_is_promoted_without_duplicate_row(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(store.SCHEMA)
        store.add_aliases(con, [("綠鐘", 2)], "auto")
        changed = store.add_aliases(con, [("綠鐘", 2)], "curated")
        row = con.execute("SELECT source FROM brand_aliases WHERE alias='綠鐘' AND brand_id=2").fetchone()
        self.assertEqual(changed, 1)
        self.assertEqual(row["source"], "curated")
        self.assertEqual(con.execute("SELECT count(*) FROM brand_aliases").fetchone()[0], 1)

    def test_entity_link_is_direction_independent_and_idempotent(self):
        con = sqlite3.connect(":memory:")
        con.row_factory = sqlite3.Row
        con.executescript(store.SCHEMA)
        self.assertEqual(store.add_entity_links(con, [(2, 1)], "curated"), 1)
        self.assertEqual(store.add_entity_links(con, [(1, 2)], "curated"), 0)
        row = store.entity_links(con)[0]
        self.assertEqual((row["brand_id_a"], row["brand_id_b"]), (1, 2))

    def test_subbrand_priority_kanebo_kate(self):
        p = pool((1, "Kanebo", "Beauty", 50000), (2, "KATE 凱婷", "Beauty", 30000))
        bi = BrandIndex(p, [], "level1_category", TH)

        r = tag("Kanebo KATE 凱婷 零瑕肌密微霧粉底液", "", "Beauty", bi, True)

        self.assertEqual(r["suggest brand id"], "2")
        self.assertEqual(r["suggest brand name"], "KATE 凱婷")

    def test_subbrand_priority_apple_iphone(self):
        p = pool((1, "Apple", "Mobile & Gadgets", 1000000), (2, "iphone", "Mobile & Gadgets", 500000))
        bi = BrandIndex(p, [], "level1_category", TH)

        r = tag("Apple iPhone 15 Pro Max 256G", "", "Mobile & Gadgets", bi, True)

        self.assertEqual(r["suggest brand id"], "2")
        self.assertEqual(r["suggest brand name"], "iphone")

    def test_subbrand_fallback_to_parent_when_subbrand_not_in_pool(self):
        p = pool((1, "Apple", "Mobile & Gadgets", 1000000))
        bi = BrandIndex(p, [], "level1_category", TH)

        r = tag("Apple iPad Air 5 64G", "", "Mobile & Gadgets", bi, True)

        self.assertEqual(r["suggest brand id"], "1")
        self.assertEqual(r["suggest brand name"], "Apple")

    def test_cross_category_partial_word_rejected(self):
        p = pool((1, "GOLF", "Motors", 50000))
        bi = BrandIndex(p, [], "level1_category", TH)

        r = tag("KING GOLF 漫畫 第1集", "", "Books", bi, True)

        # Cross-category GOLF should be rejected, falling back to 1.6
        self.assertEqual(r["suggest brand id"], "")
        self.assertEqual(r["suggest brand name"], TYPE_NEW)
        self.assertTrue(r["判斷路徑"].startswith("1.6"))

    def test_short_cjk_demoted_for_bilingual_pool_brand(self):
        # WAHEI FREIZ 和平 in Home & Living
        p = pool((1, "WAHEI FREIZ 和平", "Home & Living", 50000))
        bi = BrandIndex(p, [], "level1_category", TH)

        # Input only has 2 CJK characters "和平", no English, in different category Books
        r = tag("和平的日常生活", "", "Books", bi, True)

        # Should be demoted so that it doesn't auto-match WAHEI FREIZ with high confidence
        self.assertEqual(r["suggest brand id"], "")
        self.assertEqual(r["suggest brand name"], TYPE_NEW)
        self.assertTrue(r["判斷路徑"].startswith("1.6"))


if __name__ == "__main__":
    unittest.main()


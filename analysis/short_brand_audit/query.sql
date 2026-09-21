SELECT site, level1, item_id, raw_brand, title,
        "suggest brand name" AS matched_name, "suggest brand id" AS matched_id,
        "信心指數" AS confidence, 判斷路徑 AS path, tagged_at
        FROM item_results WHERE site = ?
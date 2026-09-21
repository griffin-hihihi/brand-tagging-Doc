"""讀 config.toml。所有可調參數集中在這裡，程式其他地方不寫死路徑。"""
from __future__ import annotations

import sys
import tomllib
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Config:
    root: Path
    site: str
    source_path: Path
    source_table: str
    cache_path: Path
    pool_path: Path
    pool_table: str | None
    pool_col_cat: str
    db_path: Path
    review_dir: Path
    output_dir: Path
    active: list[str]
    col: dict[str, str]
    th: dict[str, float]
    breakdown: list[str] = field(default_factory=list)
    write_sku_sheet: bool = False
    sample_n: int = 30

    @property
    def site_review_dir(self) -> Path:
        return self.review_dir / self.site

    @property
    def site_output_dir(self) -> Path:
        return self.output_dir / self.site


DEFAULT_COLS = {"id": "品號", "brand": "品牌", "title": "商品名稱", "cluster": "cluster",
                "l1": "level1_category", "l2": "level2_category", "l3": "level3_category",
                "sku": "SKU CODE", "price": "售價", "reviews": "評價數", "sales": "總銷量",
                "url": "商品網址", "momo_l1": "momo 原始 L1", "momo_l2": "momo 原始 L2",
                "momo_l3": "momo 原始 L3"}
DEFAULT_TH = {"auto": 95, "suggest": 75, "tier_gap": 5, "near_ratio": 85, "title_factor": 0.95, "fuzzy_cutoff": 88}


def load(path="config.toml") -> Config:
    p = Path(path).resolve()
    if not p.exists():
        sys.exit(f"✖ 找不到設定檔 {p}")
    with p.open("rb") as f:
        raw = tomllib.load(f)
    root = p.parent

    def rel(v) -> Path:
        v = Path(v)
        return v if v.is_absolute() else root / v

    src, pool, store = raw.get("source", {}), raw.get("pool", {}), raw.get("store", {})
    paths, run = raw.get("paths", {}), raw.get("run", {})
    cfg = Config(
        root=root,
        site=raw.get("site", {}).get("name", "site"),
        source_path=rel(src.get("path", "")),
        source_table=src.get("table", "output"),
        cache_path=rel(src.get("cache", "cache/goods.db")),
        pool_path=rel(pool.get("path", "brand_pool.csv")),
        pool_table=pool.get("table") or None,
        pool_col_cat=pool.get("col_cat", "level1_category"),
        db_path=rel(store.get("db", "brand_rules.db")),
        review_dir=rel(paths.get("review", "review")),
        output_dir=rel(paths.get("output", "output")),
        active=list(run.get("active", [])),
        col={**DEFAULT_COLS, **raw.get("columns", {})},
        th={**DEFAULT_TH, **raw.get("thresholds", {})},
        breakdown=list(raw.get("report", {}).get("breakdown", ["level2_category", "level3_category"])),
        write_sku_sheet=bool(raw.get("report", {}).get("write_sku_sheet", False)),
        sample_n=int(raw.get("review", {}).get("sample_n", 30)),
    )
    return cfg

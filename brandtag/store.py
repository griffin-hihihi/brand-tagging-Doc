"""規則庫：人工判斷的儲存規則。

設計原則 —— 只有一個真相來源：decision_log，append-only，永不覆蓋、永不刪除。

    decision_log   每一次人工判斷都是一列。誰、何時、從哪個檔案、原本系統建議什麼、人改成什麼。
                   改錯了也救得回來（撤銷只是再追加一列，舊的那列還在）。
    brand_rules    由 decision_log 推導出的「現況」：每個品牌字串取最新一列。程式查這張表。
    item_overrides 同上，但針對單一品號。
    brand_aliases  人工確認 / 自動學習的品牌別名，讓「莉婕」對得到「Liese 莉婕」。
    brand_entity_links  人工查證為同一品牌實體的 brand_id 配對；配對內再依 adg 選 canonical ID。

brand_rules 與 item_overrides 每次開啟資料庫時由 decision_log 重建，所以它們永遠一致；
要回溯任何一筆判斷的歷程，查 decision_log 即可。
"""
from __future__ import annotations

import datetime as dt
import getpass
import hashlib
import sqlite3
from pathlib import Path

from .const import NB_ID, OLD_TYPES, TYPE_NB, TYPE_NEW, TYPE_POOL

REVOKED = "__revoked__"

SCHEMA = """
-- ── 真相來源：append-only，永不 UPDATE / DELETE ─────────────────────────
CREATE TABLE IF NOT EXISTS decision_log(
  log_id      INTEGER PRIMARY KEY AUTOINCREMENT,
  logged_at   TEXT NOT NULL,
  actor       TEXT,              -- 誰做的判斷（預設電腦登入帳號）
  src_file    TEXT,              -- 來自哪個審核檔
  scope       TEXT NOT NULL,     -- 'brand'（品牌字串，全站共用）| 'item'（單一品號）
  rule_key    TEXT NOT NULL,     -- brand: B:xxx / T:xxx；item: 品號
  site        TEXT, level1 TEXT,
  decision    TEXT NOT NULL,     -- 品牌庫品牌 / 建議品牌庫新增品牌 / No brand / __revoked__
  brand_id    INTEGER, brand_name TEXT, nobrand_reason TEXT, note TEXT,
  raw_brand   TEXT,              -- 當時的品牌欄原文（人看的）
  raw_input   TEXT,              -- 人實際在 Excel 填的字
  fp          TEXT,              -- 這個判斷的指紋，用來擋重複匯入
  sys_decision TEXT, sys_brand_id INTEGER, sys_path TEXT, sys_conf TEXT,
  sampled     INTEGER, goods INTEGER, agree INTEGER);
CREATE INDEX IF NOT EXISTS ix_dl_key ON decision_log(scope, rule_key, log_id);

-- ── 現況（由 decision_log 重建，不要手動改）──────────────────────────
CREATE TABLE IF NOT EXISTS brand_rules(
  rule_key TEXT PRIMARY KEY, decision TEXT, brand_id INTEGER, brand_name TEXT, nobrand_reason TEXT,
  note TEXT, raw_brand TEXT, site TEXT, level1 TEXT, log_id INTEGER, updated_at TEXT, actor TEXT);
CREATE TABLE IF NOT EXISTS item_overrides(
  site TEXT, item_id TEXT, decision TEXT, brand_id INTEGER, brand_name TEXT, nobrand_reason TEXT,
  note TEXT, log_id INTEGER, updated_at TEXT, actor TEXT, PRIMARY KEY(site, item_id));

CREATE TABLE IF NOT EXISTS brand_aliases(
  alias TEXT, brand_id INTEGER, source TEXT, created_at TEXT, PRIMARY KEY(alias, brand_id));
CREATE TABLE IF NOT EXISTS brand_entity_links(
  brand_id_a INTEGER, brand_id_b INTEGER, source TEXT, created_at TEXT,
  PRIMARY KEY(brand_id_a, brand_id_b));

CREATE TABLE IF NOT EXISTS run_log(
  run_id INTEGER PRIMARY KEY AUTOINCREMENT, ran_at TEXT, site TEXT, level1 TEXT, goods INTEGER, keys INTEGER,
  imported INTEGER, agreed INTEGER, pending_keys INTEGER, pending_goods INTEGER,
  pool_goods INTEGER, new_goods INTEGER, nb_goods INTEGER, seconds REAL);
"""


def now() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def actor() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return "unknown"


def fingerprint(scope, key, decision, brand_id, brand_name, reason, note) -> str:
    raw = "|".join(str(x or "") for x in (scope, key, decision, brand_id, brand_name, reason, note))
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:16]


# ---------------------------------------------------------------- 開啟 / 升級
def open_db(path: Path) -> sqlite3.Connection:
    con = sqlite3.connect(path)
    con.row_factory = sqlite3.Row
    _migrate_v04(con)
    con.executescript(SCHEMA)
    con.commit()
    rebuild_state(con)
    return con


def _migrate_v04(con):
    """舊版 brand_tagger.py (v0.2–v0.4) 的資料 → decision_log，一次性。"""
    cols = [r[1] for r in con.execute("PRAGMA table_info(brand_rules)")]
    if not cols or "rule_key" in cols:
        return                                                   # 沒有舊表，或已經是新格式
    con.executescript(SCHEMA.replace("brand_rules(", "brand_rules_new(")
                      .replace("item_overrides(", "item_overrides_new("))
    key_col = "key" if "key" in cols else "rule_key"
    moved = 0
    if "decision_type" in cols:
        for r in con.execute(f'SELECT * FROM brand_rules'):
            d = OLD_TYPES.get(r["decision_type"], r["decision_type"])
            con.execute(
                "INSERT INTO decision_log(logged_at,actor,src_file,scope,rule_key,site,level1,decision,"
                "brand_id,brand_name,nobrand_reason,note,raw_brand,raw_input,fp) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (r["updated_at"] or now(), "migrated-v04", r["source"] if "source" in r.keys() else "",
                 "brand", r[key_col], r["site"], r["category"] if "category" in r.keys() else "", d,
                 r["brand_id"], r["brand_name"], r["nobrand_reason"], r["note"],
                 r["raw_example"] if "raw_example" in r.keys() else "", "",
                 fingerprint("brand", r[key_col], d, r["brand_id"], r["brand_name"], r["nobrand_reason"], r["note"])))
            moved += 1
    con.execute("ALTER TABLE brand_rules RENAME TO brand_rules_v04")
    if con.execute("SELECT name FROM sqlite_master WHERE name='item_overrides'").fetchone():
        for r in con.execute("SELECT * FROM item_overrides"):
            d = OLD_TYPES.get(r["decision_type"], r["decision_type"])
            con.execute(
                "INSERT INTO decision_log(logged_at,actor,src_file,scope,rule_key,site,level1,decision,"
                "brand_id,brand_name,nobrand_reason,note,raw_brand,raw_input,fp) "
                "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (r["updated_at"] or now(), "migrated-v04", "", "item", r["item_id"], r["site"], "", d,
                 r["brand_id"], r["brand_name"], r["nobrand_reason"], r["note"], "", "",
                 fingerprint("item", r["item_id"], d, r["brand_id"], r["brand_name"], r["nobrand_reason"], r["note"])))
            moved += 1
        con.execute("ALTER TABLE item_overrides RENAME TO item_overrides_v04")
    con.execute("DROP TABLE IF EXISTS brand_rules_new")
    con.execute("DROP TABLE IF EXISTS item_overrides_new")
    con.commit()
    if moved:
        print(f"   ↻ 已把舊版 {moved} 筆人工判斷轉入 decision_log（舊表保留為 *_v04）")


# ---------------------------------------------------------------- 現況重建
def rebuild_state(con):
    """brand_rules / item_overrides 一律由 decision_log 重算，確保兩者永遠一致。"""
    con.execute("DELETE FROM brand_rules")
    con.execute("DELETE FROM item_overrides")
    latest = """
        SELECT d.* FROM decision_log d
        JOIN (SELECT scope, rule_key, MAX(log_id) AS m FROM decision_log GROUP BY scope, rule_key) t
          ON d.log_id = t.m
        WHERE d.scope = ? AND d.decision != ?"""
    for r in con.execute(latest, ("brand", REVOKED)).fetchall():
        con.execute("INSERT INTO brand_rules VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
                    (r["rule_key"], r["decision"], r["brand_id"], r["brand_name"], r["nobrand_reason"],
                     r["note"], r["raw_brand"], r["site"], r["level1"], r["log_id"], r["logged_at"], r["actor"]))
    for r in con.execute(latest, ("item", REVOKED)).fetchall():
        con.execute("INSERT OR REPLACE INTO item_overrides VALUES(?,?,?,?,?,?,?,?,?,?)",
                    (r["site"], r["rule_key"], r["decision"], r["brand_id"], r["brand_name"],
                     r["nobrand_reason"], r["note"], r["log_id"], r["logged_at"], r["actor"]))
    con.commit()


def current(con, scope="brand", site=None) -> dict:
    if scope == "brand":
        return {r["rule_key"]: dict(r) for r in con.execute("SELECT * FROM brand_rules")}
    return {r["item_id"]: dict(r) for r in
            con.execute("SELECT * FROM item_overrides WHERE site=?", (site,))}


# ---------------------------------------------------------------- 寫入
def record(con, entries, src_file: str, site: str, level1: str) -> tuple[int, int]:
    """追加人工判斷。與現況相同的（例如重跑同一個沒改過的審核檔）自動略過。

    回傳 (實際寫入筆數, 其中同意系統建議的筆數)。
    """
    cur_b = current(con, "brand")
    cur_i = current(con, "item", site)
    who, ts = actor(), now()
    n = agreed = 0
    for e in entries:
        scope, key = e["scope"], e["rule_key"]
        fp = fingerprint(scope, key, e["decision"], e.get("brand_id"), e.get("brand_name"),
                         e.get("nobrand_reason"), e.get("note"))
        prev = (cur_b if scope == "brand" else cur_i).get(key)
        if prev:
            prev_fp = fingerprint(scope, key, prev["decision"], prev["brand_id"], prev["brand_name"],
                                  prev["nobrand_reason"], prev["note"])
            if prev_fp == fp:
                continue                                          # 沒變 → 不重複記錄
        elif e["decision"] == REVOKED:
            continue                                              # 本來就沒有規則，不用撤銷
        con.execute(
            "INSERT INTO decision_log(logged_at,actor,src_file,scope,rule_key,site,level1,decision,brand_id,"
            "brand_name,nobrand_reason,note,raw_brand,raw_input,fp,sys_decision,sys_brand_id,sys_path,sys_conf,"
            "sampled,goods,agree) VALUES(" + ",".join("?" * 22) + ")",
            (ts, who, src_file, scope, key, site, level1, e["decision"], e.get("brand_id"), e.get("brand_name"),
             e.get("nobrand_reason"), e.get("note"), e.get("raw_brand"), e.get("raw_input"), fp,
             e.get("sys_decision"), e.get("sys_brand_id"), e.get("sys_path"), e.get("sys_conf"),
             e.get("sampled", 0), e.get("goods", 0), e.get("agree")))
        n += 1
        agreed += int(bool(e.get("agree")))
    if n:
        con.commit()
        rebuild_state(con)
    return n, agreed


def add_aliases(con, pairs, source="human") -> int:
    ts, n = now(), 0
    priority = {"auto": 1, "curated": 2, "human": 3}
    for alias, bid in pairs:
        if alias and bid is not None:
            bid = int(bid)
            old = con.execute("SELECT source FROM brand_aliases WHERE alias=? AND brand_id=?",
                              (alias, bid)).fetchone()
            if old is None:
                con.execute("INSERT INTO brand_aliases VALUES(?,?,?,?)", (alias, bid, source, ts))
                n += 1
            elif priority.get(source, 0) > priority.get(old["source"], 0):
                con.execute("UPDATE brand_aliases SET source=?, created_at=? WHERE alias=? AND brand_id=?",
                            (source, ts, alias, bid))
                n += 1
    con.commit()
    return n


def aliases(con) -> list:
    return con.execute("SELECT alias, brand_id, source FROM brand_aliases").fetchall()


def add_entity_links(con, pairs, source="curated") -> int:
    """加入人工查證的同實體 ID 配對。ID 固定小者在前，避免反向重複。"""
    ts, n = now(), 0
    for b1, b2 in pairs:
        a, b = sorted((int(b1), int(b2)))
        if a == b:
            continue
        old = con.execute("SELECT 1 FROM brand_entity_links WHERE brand_id_a=? AND brand_id_b=?",
                          (a, b)).fetchone()
        if old is None:
            con.execute("INSERT INTO brand_entity_links VALUES(?,?,?,?)", (a, b, source, ts))
            n += 1
    con.commit()
    return n


def entity_links(con) -> list:
    return con.execute("SELECT brand_id_a, brand_id_b, source FROM brand_entity_links").fetchall()


def log_run(con, **kw):
    cols = ["site", "level1", "goods", "keys", "imported", "agreed", "pending_keys", "pending_goods",
            "pool_goods", "new_goods", "nb_goods", "seconds"]
    con.execute("INSERT INTO run_log(ran_at," + ",".join(cols) + ") VALUES(" + ",".join("?" * (len(cols) + 1)) + ")",
                [now()] + [kw.get(c) for c in cols])
    con.commit()


def save_results(con, df, site, level1, cols):
    """最新一次的品號層結果，供之後上傳 Google Sheet 用。"""
    schema = ["site", "level1", "item_id", "raw_brand", "title"] + cols + ["tagged_at"]
    have = [r[1] for r in con.execute("PRAGMA table_info(item_results)")]
    if have and have != schema:
        con.execute("DROP TABLE item_results")
    con.execute("CREATE TABLE IF NOT EXISTS item_results(" + ",".join(f'"{c}" TEXT' for c in schema) + ")")
    con.execute("CREATE INDEX IF NOT EXISTS ix_ir ON item_results(site, level1)")
    con.execute("DELETE FROM item_results WHERE site=? AND level1=?", (site, level1))
    df = df.copy()
    df["site"], df["level1"], df["tagged_at"] = site, level1, now()
    df[schema].astype(str).to_sql("item_results", con, if_exists="append", index=False, chunksize=20000)
    con.commit()

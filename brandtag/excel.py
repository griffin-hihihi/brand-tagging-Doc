"""Excel 輸出：審核檔（你要編輯的）與結果檔（給下游看的）分開。"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from .const import (CONF_COL, H_PICK, H_REASON, HUMAN_COLS, PATHS, PICK_OPTIONS, REASONS, RESULT_COLS,
                    REVIEW_TH, TYPE_NB, TYPE_NEW)
from .review import FREEZE_AT, ITEM_COLS, REVIEW_COLS

# 商品網址欄如果讓 xlsxwriter 自動轉成超連結，會撞到 Excel 每張工作表 65,530 個連結的上限：
# 超過的部分被默默丟掉，還會噴幾十萬行警告拖慢速度。網址當純文字存就好。
_OPTS = {"options": {"strings_to_urls": False}}


def _writer(path):
    return pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs=_OPTS)

GUIDE = [
    "【這個檔案怎麼用】",
    "  這是「審核檔」，一個商品類別一個，檔名固定、不會越積越多。你直接在這裡填，存檔關閉，再跑一次同樣的指令即可。",
    "  跑完程式會原地更新這個檔：你的判斷會寫進規則庫，已審的列回填成正式結論並排到後面，還沒審的排最前面。",
    "",
    "【每一輪怎麼做】",
    "1. 到「審核」頁。已經照「狀態」排好：① 待審 → ② 抽查 → ③ 已審 → ④ 自動（同狀態內商品數多的在前）",
    "   你只需要看 ① 和 ②。③④ 是給你回頭查的。",
    "2. 在黃色的【人工】判斷 填（有下拉選單，也可以直接打字）：",
    "      v　　　　　　　　　 同意系統建議",
    "      1 / 2 / 3　　　　　 改選 shp brand name1 / 2 / 3（候選就在填答格左邊）",
    "      brand_id 或品牌名　 改成品牌庫的其他品牌（可直接貼「品牌名 [12345]」這種格式）",
    f"      {TYPE_NB}　　　　　 白牌、沒有品牌（再選【人工】No brand原因）",
    f"      {TYPE_NEW}　 有品牌但品牌庫沒有（名稱預設用品牌欄，要改名寫「新增:品牌名」）",
    "      -　　　　　　　　　 撤銷先前的判斷，讓這個品牌字串回到系統自動判斷",
    "   不確定的留空就好。",
    "3. 存檔 → 關閉 Excel → 再執行一次同樣的指令",
    "",
    "【重要：留空 ≠ 刪除】",
    "  已審過的列會自動回填你上次的判斷。把它清成空白不會刪掉規則（避免誤刪）。",
    "  真的要撤銷，請填「-」。撤銷也只是追加一筆紀錄，原本的判斷在 decision_log 裡永遠查得到。",
    "",
    "",
    "═══════════════════════════════════════════════════════════════════",
    "【完整判斷樹】每一筆商品必定落在、而且只落在一個節點",
    "═══════════════════════════════════════════════════════════════════",
    "",
    "  問題一：品牌欄有沒有填東西？",
    "",
    "  ┌─ 有填 ───────────────────────────────────────────────────────",
    "  │",
    "  │   寫的是「無品牌/其他/副廠/OEM」嗎？",
    "  │     └─ 是 ────────────────────────────────► 1.1  No brand",
    "  │",
    "  │   寫的是「適用/相容/副廠」這種描述嗎？",
    "  │     └─ 是 ────────────────────────────────► 1.2  No brand",
    "  │",
    "  │   拿去比對品牌庫，對得到嗎？",
    "  │     ├─ 對到，而且品牌庫只有這一筆 ──────────► 1.3  品牌庫品牌",
    "  │     ├─ 對到，而且有證據確認品牌庫好幾筆是同一個品牌",
    "  │     │    （名稱相同、只差公司後綴、長英文名一致，或有人工查證對照）",
    "  │     │    → 確認同實體後，取 adg 最高的一筆 ───► 1.4  品牌庫品牌",
    "  │     ├─ 對到好幾個「名稱不同」的品牌，無法自動決定",
    "  │     │    → 取最高分但信心很低，交給你判斷 ──► 1.5  品牌庫品牌",
    "  │     └─ 對不到",
    "  │          這串字看起來像品牌名嗎？",
    "  │            ├─ 像 ──────────────────────────► 1.6  建議新增品牌",
    "  │            └─ 不像（是品類詞/規格/通用詞）──► 1.7  No brand",
    "  │",
    "  └─ 沒填（空白）───────────────────────────────────────────────",
    "      ",
    "      商品是配件/相容品嗎？（保護貼、手機殼、支架…）",
    "        └─ 是。標題裡的品牌是「它相容的品牌」不是它自己的",
    "             ─────────────────────────────────► 2.1  No brand",
    "      ",
    "      商品名稱開頭有【】嗎？",
    "        ├─ 有 →【】拿去比對品牌庫",
    "        │       ├─ 對得到 ───────────────────► 2.2  品牌庫品牌",
    "        │       └─ 對不到",
    "        │            【】看起來像品牌嗎？",
    "        │              ├─ 像 ────────────────► 2.3  建議新增品牌",
    "        │              └─ 不像 → 往下走",
    "        │",
    "        └─ 沒有【】或【】不是品牌",
    "             改從商品名稱內文找（要夠長夠獨特：英文≥6字元、中文≥4字）",
    "               ├─ 找到 ──────────────────────► 2.4  品牌庫品牌",
    "               └─ 找不到 ────────────────────► 2.5  No brand",
    "",
    "  另外：這個品牌字串你之前審過 ─────────────────► 3.1  直接套用你的判斷",
    "        這個品號你單獨指定過 ───────────────────► 3.2  直接套用你的判斷",
    "",
    "",
    "【多個候選時怎麼選】順序固定，由上往下：",
    "   ① 先確認證據與品牌實體：人工確認／查證對照優先；1–3 字元短縮寫不能單獨證明同實體",
    "   ② 已確認同實體才比 adg：取業績最高的 ID（H&H / Herb & Health → Herb & Health）",
    "   ③ 最小顆粒度 → 名稱完整出現在品牌字串中、而且最具體的優先（Apple iPhone → iPhone）",
    "   ④ 名稱互異又同分 → 不武斷選擇，標為低信心交給你",
    "   auto 別名不能和同一筆短英文互相佐證；中文未獨立證實時，信心最高 74",
    "",
    "",
    "【信心指數 0–100】一個數字，代表系統對這個結論有多確定。",
    "  與結論類型無關 —— 判定為 No brand 也可以是 97 分。",
    "",
    "     90 以上　可直接採用（綠色）",
    "     75–89 　 大致可信，靠抽查★驗證（黃色）",
    "     75 以下　需要你看（紅色）—— 審核頁狀態欄會標「① 待審」",
    "",
    "  沒有「是否待人工判斷」這一欄了，用信心指數篩選就好。",
    "",
    "",
    "【你的判斷會被記到哪裡】",
    "  brand_rules.db 的 decision_log —— append-only，誰、何時、從哪個檔、系統原本建議什麼、你改成什麼，全都留著。",
    "  規則以「品牌字串」為單位，全站共用：在美妝審過的品牌，保健、居家遇到同樣字串直接套用，不用再審一次。",
    "  反過來說，判錯一個大品牌也會擴散到所有類別，所以商品數上千的那幾個要特別謹慎。",
    "",
    "【想針對單一商品修改】到「單品覆蓋」頁填品號，只會影響那一個品號，不影響同品牌的其他商品。",
    "",
    "【判斷路徑一覽】",
]


def _formats(wb):
    spec = {
        "h_blue": {"bold": True, "bg_color": "#DDEBF7", "border": 1, "text_wrap": True, "valign": "vcenter"},
        "h_yel": {"bold": True, "bg_color": "#FFD966", "border": 1, "text_wrap": True, "valign": "vcenter"},
        "h": {"bold": True, "bg_color": "#F2F2F2", "border": 1, "text_wrap": True, "valign": "vcenter"},
        "yel": {"bg_color": "#FFF2CC"}, "hi": {"bg_color": "#C6EFCE"}, "mid": {"bg_color": "#FFEB9C"},
        "lo": {"bg_color": "#FFC7CE"}, "title": {"bold": True, "font_size": 12},
        "pend": {"bg_color": "#FCE4D6", "bold": True}, "done": {"font_color": "#808080"},
    }
    return {k: wb.add_format(v) for k, v in spec.items()}


def _guide_sheet(w, F):
    wb = w.book
    ws = wb.add_worksheet("說明")
    w.sheets["說明"] = ws
    for i, line in enumerate(GUIDE):
        ws.write(i, 0, line, F["title"] if line.startswith("【") else None)
    r0 = len(GUIDE)
    for j, h in enumerate(["代碼", "節點名稱", "結論", "判斷條件", "例子"]):
        ws.write(r0, j, h, F["h"])
    for i, (k, v) in enumerate(PATHS.items(), r0 + 1):
        ws.write_row(i, 0, [k, v[0], v[1], v[2], v[3]])
    ws.set_column(0, 0, 16)
    ws.set_column(1, 1, 28)
    ws.set_column(2, 2, 6)
    ws.set_column(3, 4, 60)


def _human_sheet(w, F, df, name, widths, freeze_col, conf_col=True):
    df.to_excel(w, sheet_name=name, index=False)
    ws = w.sheets[name]
    cols, n = list(df.columns), len(df)
    for j, c in enumerate(cols):
        ws.write(0, j, c, F["h_yel"] if c in HUMAN_COLS else F["h_blue"] if c in RESULT_COLS else F["h"])
        ws.set_column(j, j, widths.get(c, 22 if c in RESULT_COLS + HUMAN_COLS else 14))
    ws.set_row(0, 32)
    ws.freeze_panes(1, freeze_col)
    ws.autofilter(0, 0, max(n, 1), len(cols) - 1)
    if n:
        if conf_col and CONF_COL in cols:
            j = cols.index(CONF_COL)
            ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": ">=", "value": 90, "format": F["hi"]})
            ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "between", "minimum": REVIEW_TH,
                                               "maximum": 89, "format": F["mid"]})
            ws.conditional_format(1, j, n, j, {"type": "cell", "criteria": "<", "value": REVIEW_TH,
                                               "format": F["lo"]})
        if "狀態" in cols:
            j = cols.index("狀態")
            ws.conditional_format(1, j, n, j,
                                  {"type": "text", "criteria": "begins with", "value": "①", "format": F["pend"]})
            ws.conditional_format(1, j, n, j,
                                  {"type": "text", "criteria": "begins with", "value": "④", "format": F["done"]})
        for c in HUMAN_COLS:
            j = cols.index(c)
            ws.conditional_format(1, j, n, j, {"type": "no_errors", "format": F["yel"]})
        j = cols.index(H_PICK)
        ws.data_validation(1, j, n, j, {"validate": "list", "source": PICK_OPTIONS, "show_error": False,
                                        "input_title": "填寫方式",
                                        "input_message": "v=同意｜1/2/3=選候選｜brand_id 或品牌名｜- =撤銷"})
        j = cols.index(H_REASON)
        ws.data_validation(1, j, n, j, {"validate": "list", "source": REASONS, "show_error": False})
    ws.write_comment(0, cols.index(H_PICK),
                     f"v = 同意建議\n1/2/3 = 改選第幾個候選\nbrand_id 或品牌名 = 改成該品牌\n"
                     f"{TYPE_NB} = 白牌\n{TYPE_NEW} = 品牌庫沒有\n- = 撤銷先前判斷",
                     {"x_scale": 1.8, "y_scale": 1.8})
    if "key" in cols:
        ws.set_column(cols.index("key"), cols.index("key"), None, None, {"hidden": True})


def write_review(path: Path, review: pd.DataFrame, items: pd.DataFrame):
    """審核檔：你編輯的那一份。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    with _writer(path) as w:
        F = _formats(w.book)
        _guide_sheet(w, F)
        _human_sheet(w, F, review, "審核",
                     {"狀態": 11, "抽查": 5, "歧義": 5, "累積覆蓋": 9, "群組": 7, "品牌欄": 22,
                      "商品數": 7, "suggest brand name": 24, CONF_COL: 9,
                      "shp brand name1": 26, "shp brand name2": 26, "shp brand name3": 26,
                      "範例商品名稱": 40, "範例商品網址": 36, "判斷說明": 45, "判斷路徑": 22, "上次審核": 18,
                      "商品名稱【】": 18, "主要類目": 16}, FREEZE_AT)
        if items.empty:
            items = pd.DataFrame([{c: "" for c in ITEM_COLS}])
        _human_sheet(w, F, items, "單品覆蓋",
                     {"品號": 14, "商品名稱": 46, "品牌欄": 20, "目前結果": 24, "上次審核": 20}, 1, conf_col=False)


def write_result(path: Path, full: pd.DataFrame, newb: pd.DataFrame, stats: list, title_col: str):
    """結果檔：下游要用的那一份，每次重新產生，不要編輯它。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    if len(full) > 1_000_000:
        raise SystemExit("✖ 這個 L1 超過 100 萬筆，Excel 放不下，請改用 --no-excel 走資料庫輸出")
    with _writer(path) as w:
        wb = w.book
        F = _formats(wb)
        full.to_excel(w, sheet_name="總結果", index=False)
        ws = w.sheets["總結果"]
        cols = list(full.columns)
        for j, c in enumerate(cols):
            ws.write(0, j, c, F["h_blue"] if c in RESULT_COLS else F["h"])
            ws.set_column(j, j, 42 if c == title_col else 40 if c == "判斷說明" else
                          22 if c in RESULT_COLS else 14)
        ws.set_row(0, 30)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, max(len(full), 1), len(cols) - 1)
        if len(full) and CONF_COL in cols:
            j = cols.index(CONF_COL)
            ws.conditional_format(1, j, len(full), j,
                                  {"type": "cell", "criteria": ">=", "value": 90, "format": F["hi"]})
            ws.conditional_format(1, j, len(full), j,
                                  {"type": "cell", "criteria": "between", "minimum": REVIEW_TH,
                                   "maximum": 89, "format": F["mid"]})
            ws.conditional_format(1, j, len(full), j,
                                  {"type": "cell", "criteria": "<", "value": REVIEW_TH, "format": F["lo"]})

        newb.to_excel(w, sheet_name="新增品牌建議", index=False)
        ws = w.sheets["新增品牌建議"]
        for j, c in enumerate(newb.columns):
            ws.write(0, j, c, F["h"])
            ws.set_column(j, j, 40 if c == "範例商品名稱" else 24)

        ws = wb.add_worksheet("統計")
        w.sheets["統計"] = ws
        r = 0
        for t, df in stats:
            ws.write(r, 0, t, F["title"])
            if len(df):
                df.to_excel(w, sheet_name="統計", startrow=r + 1, index=False)
                for j, c in enumerate(df.columns):
                    ws.write(r + 1, j, c, F["h"])
            r += len(df) + 4
        ws.set_column(0, 0, 32)
        ws.set_column(1, 9, 16)

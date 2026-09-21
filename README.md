# 競網品牌 Tagging

把競網（momo）商品的品牌欄，對應到 shp 品牌庫的 `brand_id`。
對不到的分成兩種：**建議品牌庫新增品牌**（有品牌但品牌庫沒有）與 **No brand**（本來就是白牌）。

---

## 為什麼是一個 L1 一個 L1 跑

545,572 個品號，收斂後只有 **11,407 個不重複的品牌字串**。
各 L1 分開算會是 17,682 個 —— 代表約 35% 的品牌字串跨 L1 重複。

規則庫是**以品牌字串為單位、全站共用**的，所以在 Beauty 審過的「DHC」，
Health、Mother & Baby 遇到同一個字串會直接套用，不用再審一次。
越後面跑的 L1，需要人工審的比例越低。這是整個流程會收斂的原因。

---

## 三個指令

```bat
py run.py status                 :: 每個 L1 現在進行到哪、還剩多少要審
py run.py tag                    :: 跑 config.toml 裡 [run] active 列的 L1
py run.py tag --l1 Beauty        :: 臨時只跑一個
py run.py tag --all              :: 22 個全跑
```

## 一輪的流程

```
py run.py tag --l1 Beauty
        ↓
開 review/momo/Beauty.xlsx →「審核」頁
        ↓
「狀態」欄已排好：① 待審 → ② 抽查 → ③ 已審 → ④ 自動
只要看 ① 和 ②，在黃色的【人工】判斷 填
        ↓
存檔、關閉 Excel
        ↓
py run.py tag --l1 Beauty        ← 同樣的指令，再跑一次
```

第二次跑的時候，程式會先把你填的判斷讀進規則庫，再用新規則重跑，
然後**原地更新同一個檔案**：已審的列回填成正式結論並排到後面，還沒審的排最前面。
檔名固定，不會越積越多。

### 【人工】判斷 可以填什麼

| 填 | 意思 |
|---|---|
| `v` | 同意系統建議 |
| `1` / `2` / `3` | 改選 shp brand name1 / 2 / 3 |
| `brand_id` 或品牌名 | 改成品牌庫的其他品牌（可直接貼 `品牌名 [12345]`） |
| `No brand` | 白牌，再選【人工】No brand原因 |
| `建議品牌庫新增品牌` | 有品牌但品牌庫沒有；要改名寫 `新增:品牌名` |
| `-` | **撤銷**先前的判斷，讓它回到系統自動判斷 |
| 留空 | 不確定，跳過 |

> **留空 ≠ 刪除。** 已審過的列會自動回填你上次的判斷，把它清成空白不會刪規則（避免誤刪）。
> 真的要撤銷請填 `-`。

只想改**某一個品號**、不想影響同品牌的其他商品 → 到「單品覆蓋」頁填品號。

---

## 人工判斷存在哪裡

`brand_rules.db`，核心是 **`decision_log`：append-only，永不覆蓋、永不刪除**。

| 表 | 內容 |
|---|---|
| `decision_log` | **唯一的真相來源**。每一次人工判斷一列：誰、何時、從哪個檔、系統原本建議什麼、你改成什麼、同不同意 |
| `brand_rules` | 由 `decision_log` 推導出的「現況」（每個品牌字串取最新一列）。程式查這張表 |
| `item_overrides` | 同上，但針對單一品號 |
| `brand_aliases` | 人工確認 / 自動學習的別名，讓「莉婕」對得到「Liese 莉婕」 |
| `item_results` | 最新一次的品號層結果，之後上傳 Google Sheet 用 |
| `run_log` | 每次執行的統計，`status` 指令讀這張 |

`brand_rules` / `item_overrides` 每次開啟資料庫時由 `decision_log` 重建，所以兩者永遠一致。
改錯了也救得回來 —— 撤銷只是再追加一列，原本那列還在。

```bat
py run.py history --key B:dhc    :: 查某個品牌字串的完整判斷歷程
py run.py history --limit 50     :: 看最近 50 筆異動
```

### 冪等

同一個審核檔重複跑不會重複記錄 —— 只有**跟現況不一樣**的判斷才會追加一列。
所以放心重跑。

---

## 檔案長怎樣

```
brand_tagging/
  config.toml                 ← 控制台：跑哪些 L1、路徑、閾值
  run.py                      ← 入口
  brandtag/                   ← 程式
    const.py    共用常數（三種結論、No brand 原因、判斷路徑一覽）
    text.py     文字正規化
    source.py   raw sqlite → 品號層快取
    index.py    品牌庫索引（7.6 萬品牌的多重索引）
    engine.py   判斷引擎與決策樹
    store.py    規則庫：decision_log 與現況重建
    review.py   審核檔的讀回與產生
    excel.py    Excel 輸出
    report.py   統計
  review/momo/<L1>.xlsx       ← ★ 你編輯的：審核 / 單品覆蓋 / 說明
  output/momo/<L1>_result.xlsx ← 產出：總結果 / 新增品牌建議 / 統計（每次重產，別編輯）
  cache/momo_goods.db         ← raw 抽取快取（raw 換檔會自動重建）
  brand_rules.db              ← 規則庫
```

**審核檔與結果檔是分開的。** 審核檔只有品牌字串層（幾千列，幾百 KB，開起來很快），
結果檔才是品號層（Books 有 16.9 萬列、19 MB）。你每天開的是前者。

---

## 判斷邏輯

三種結論：`品牌庫品牌` / `建議品牌庫新增品牌` / `No brand`，
各自帶判斷路徑代碼（A1–A7、B1–B2、C1、N1–N3、Z1–Z4、R1–R2）與高/中/低信心。
完整對照表在每個審核檔的「說明」頁。

多個候選時走決策樹：

1. **證據與實體** —— 先判斷候選是不是同一品牌；人工確認／查證對照優先。1–3 字元的短縮寫不能單獨證明同實體
2. **同實體才比業績** —— 品牌庫重複 ID 已確認為同一實體後，選 `adg` 最高者（H&H / Herb & Health → Herb & Health）
3. **最小顆粒度** —— 名稱完整出現在競品品牌中、而且最具體的優先（Apple iPhone → iPhone）
4. **歧義保留** —— 名稱不同卻同分時降為低信心，交給人判斷

系統自己學到的 `auto` 別名不能和同一筆短英文互相佐證。短縮寫的中文名稱沒有獨立證據時，信心最高 74，會進待審清單。

**信心與抽查**：高 = 可直接用；中 = 大致可信；低 = 一定要看。
系統每輪會從高/中信心裡各抽 30 個標★，你順手審一下，
結果檔「統計」頁的④就會累積出各路徑、各信心的實際準確率 —— 這是你判斷「中信心能不能放心用」的依據。

---

## 之後：上傳 Google Sheet

`item_results` 表已經是品號層的完整結果，隨時可以匯出：

```bat
py run.py export --l1 Beauty
py run.py export                 :: 全部 L1
```

要接 Google Sheet 時，從這張表讀就好，不用重跑 tagging。

---

## 環境

```bat
py -m pip install pandas openpyxl xlsxwriter rapidfuzz opencc-python-reimplemented
```

Python 3.11 以上（用到 stdlib 的 `tomllib` 讀 config.toml）。

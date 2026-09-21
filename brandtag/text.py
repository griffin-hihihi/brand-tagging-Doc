"""文字正規化與品牌字串的基礎判斷。

這一層只做「字串怎麼看」的事，不碰品牌庫、不碰資料庫，方便單獨測試。
"""
from __future__ import annotations

import re
import unicodedata

import pandas as pd

try:
    from opencc import OpenCC

    to_trad = OpenCC("s2t").convert
except Exception:                                        # pragma: no cover - 沒裝 opencc 時退化
    def to_trad(s):
        return s

CJK_RE = re.compile(r"[㐀-鿿豈-﫿]")

# 品牌欄常見的雜訊詞，比對前先拿掉
NOISE_WORDS = ["官方直營", "官方旗艦店", "旗艦店", "官方授權", "官方", "台灣總代理", "總代理", "即期品",
               "台灣公司貨", "公司貨", "原廠", "正品", "現貨", "免運"]
# 【】內常見的產地前綴，去掉才對得到品牌
ORIGIN_PREFIX = ["紐西蘭", "日本", "韓國", "美國", "法國", "德國", "英國", "澳洲", "義大利", "台灣", "泰國"]
# 品牌欄寫這些字 = 明確宣告沒有品牌
NOBRAND_WORDS = {"無品牌", "其他", "其它", "副廠", "oem", "nobrand", "none", "無", "通用", "自有品牌", "mit",
                 "白牌", "0"}
# 【】裡出現這些字，多半是商品描述不是品牌
GENERIC_WORDS = ["收納", "居家", "生活", "買一送一", "限時", "特價", "熱銷", "新品", "任選", "組合", "入組",
                 "超值", "現貨", "免運", "團購", "禮盒", "必備", "推薦", "精選", "款"]
COMPAT_RE = re.compile(r"(副廠|通用款|(適用|相容|兼容|for)\s*(iphone|ipad|apple|samsung|switch|airpods|galaxy|"
                       r"macbook|pixel|蘋果|三星|小米))", re.I)
# 配件類商品：標題裡出現的品牌通常是「它相容的品牌」，不是「它自己的品牌」。
# 例：「iPhone SE 非滿版鋼化膜」的品牌不是 Apple，是那個做保護貼的白牌廠商。
ACCESSORY_RE = re.compile(r"(保護貼|保護殼|保護套|手機殼|鋼化膜|玻璃貼|滿版|皮套|背蓋|空壓殼|"
                          r"支架|杯架|車架|掛架|收納架|置物架|轉接|傳輸線|充電線|充電座|替換|"
                          r"專用|配件|副廠|相容|適用|通用)", re.I)
# 這些英文單字太泛用，單獨命中不足以當品牌證據
STOP_TOKENS = {"the", "body", "shop", "house", "home", "life", "care", "baby", "beauty", "paris", "pro",
               "professional", "plus", "health", "healthcare", "lab", "labs", "official", "store", "company",
               "international", "taiwan", "japan", "korea", "group", "entertainment", "choice", "goals",
               "signature", "coffee", "tech", "village", "face", "north", "balance", "good", "smile", "mobile",
               "national", "geographic", "digital", "western", "space", "power", "cook", "dream", "trend",
               "future", "biotech", "educational", "nutrition", "optimum", "nature", "natural", "organic"}


def blank(v) -> bool:
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() in ("", "nan", "None", "<NA>")


def clean_raw(s) -> str:
    """去雜訊詞、收斂空白，但保留大小寫與中英文原貌（顯示用）。"""
    s = "" if blank(s) else unicodedata.normalize("NFKC", str(s)).strip()
    for w in NOISE_WORDS:
        s = s.replace(w, " ")
    return re.sub(r"\s+", " ", s).strip()


def normalize(s) -> str:
    """比對用的正規化：簡轉繁、轉小寫、去標點與空白，只留英數與中日韓字。"""
    s = to_trad(unicodedata.normalize("NFKC", "" if blank(s) else str(s))).lower()
    s = s.replace("’", "").replace("'", "").replace("`", "")
    return "".join(ch for ch in s if ch.isalnum() or CJK_RE.match(ch))


def surface_key(s) -> str:
    """保留標點的精確比對鍵；用來區分 H&H 與 HH 這類短名稱。"""
    return re.sub(r"\s+", "", to_trad(clean_raw(s)).casefold())


def split_parts(s):
    """回傳 (完整正規化字串, 只留英數的部分, 只留中文的部分)。"""
    n = normalize(s)
    return n, "".join(re.findall(r"[a-z0-9]+", n)), "".join(CJK_RE.findall(n))


def latin_tokens(s) -> set:
    s = to_trad(unicodedata.normalize("NFKC", str(s or ""))).lower().replace("’", "").replace("'", "")
    return {t for t in re.findall(r"[a-z0-9]+", s) if len(t) >= 4 and t not in STOP_TOKENS}


def is_bilingual(s: str) -> bool:
    """「英文 中文」或「中文 英文」中間有空格 → 正常雙語品牌名（MUK 潮嘜）；LG生活健康 則不是。"""
    return bool(re.fullmatch(r"[^㐀-鿿]*[A-Za-z0-9][^㐀-鿿]*\s+[㐀-鿿].*", s)
                or re.fullmatch(r"[㐀-鿿][^A-Za-z]*\s+[A-Za-z].*", s))


def extract_bracket(title) -> str:
    """取商品名稱開頭的【】內容，並剝掉產地前綴。"""
    m = re.search(r"【(.*?)】", "" if blank(title) else str(title))
    if not m:
        return ""
    b = clean_raw(m.group(1))
    for p in ORIGIN_PREFIX:
        if b.startswith(p) and len(b) > len(p):
            b = b[len(p):]
    return b.strip()


def is_generic(b: str) -> bool:
    """【】內容看起來是商品描述而非品牌。"""
    return (not b) or any(w in b for w in GENERIC_WORDS) or len(normalize(b)) > 14


def long_enough(a: str) -> bool:
    """字串長到足以當「內含品牌名」的證據（避免 2 個字母就命中）。"""
    cjk = len(CJK_RE.findall(a))
    return (cjk == 0 and len(a) >= 5) or cjk >= 3


def id_str(x) -> str:
    return "" if blank(x) else str(int(float(x)))


def brand_key(raw_brand, title, item_id) -> str:
    """品牌字串的穩定 key：優先用品牌欄，其次【】，都沒有才退化成單品。

    這個 key 是全站、跨 L1 共用的 —— 在 Beauty 審過的品牌字串，Health 遇到同樣的字串直接套用。
    """
    b = normalize(clean_raw(raw_brand))
    if b:
        return "B:" + b
    br = extract_bracket(title)
    if br and not is_generic(br):
        return "T:" + normalize(br)
    return "I:" + str(item_id)


# ================================================================ 這串字是不是品牌？
# 競網的品牌欄常被填入品類詞、規格或促銷語。把它們判成「建議新增品牌」會污染品牌庫的
# 待決策清單，所以在判定為新品牌之前，先過這一關。

# 單獨出現時幾乎不可能是品牌的通用詞
NOT_BRAND_WORDS = {
    "國家", "幸福", "生活", "精品", "時尚", "經典", "自然", "健康", "美麗", "快樂", "溫馨",
    "農會", "漁會", "合作社", "生產者", "小農", "產地", "批發", "零售", "量販", "賣場",
    "百貨", "超市", "商行", "商店", "專櫃", "專賣店", "直營", "代理", "進口", "外銷",
    "通用", "萬用", "多功能", "高品質", "高質感", "超值", "優惠", "特價", "促銷", "限量",
    "熱銷", "暢銷", "推薦", "嚴選", "精選", "優選", "良品", "好物", "選物", "雜貨",
    "general", "universal", "standard", "classic", "premium", "quality", "value",
    "original", "genuine", "official", "nil", "none", "null", "other", "misc", "etc",
}
# 品類詞：出現在字串裡就代表這是商品描述不是品牌
CATEGORY_WORDS = [
    "收納", "置物", "整理", "清潔", "洗劑", "保養", "護理", "保健", "營養", "食品",
    "用品", "配件", "工具", "器材", "設備", "零件", "耗材", "文具", "玩具", "服飾",
    "鞋款", "包款", "飾品", "家具", "寢具", "廚具", "餐具", "杯架", "掛勾", "衣架",
    "保護貼", "手機殼", "充電", "傳輸", "轉接", "電池", "燈具", "照明",
]
# 規格樣式：數量、容量、尺寸
SPEC_RE = re.compile(r"(\d+\s*(入|組|件|包|盒|片|支|條|個|雙|ml|ML|cc|CC|g|G|kg|KG|cm|CM|mm|吋|寸|L))"
                     r"|([0-9]{3,})")


def looks_like_brand(s: str) -> tuple[bool, str]:
    """判斷這串字是否像品牌名。回傳 (是否像品牌, 不像的理由)。

    刻意保守：只在有明確證據時才判定「不是品牌」，避免把冷門品牌誤殺成 No brand。
    """
    raw = clean_raw(s)
    if not raw:
        return False, "空白"
    n = normalize(raw)
    if not n:
        return False, "沒有可辨識的文字"
    if n in NOT_BRAND_WORDS or raw.lower() in NOT_BRAND_WORDS:
        return False, "通用詞，非特定品牌"
    if n.isdigit():
        return False, "純數字"
    for w in CATEGORY_WORDS:
        if w in raw:
            return False, f"含品類詞「{w}」，屬商品描述"
    if SPEC_RE.search(raw):
        return False, "含規格或數量，屬商品描述"
    if len(CJK_RE.findall(n)) == 0 and len(n) <= 1:
        return False, "過短"
    if len(n) > 25:
        return False, "過長，不像品牌名"
    return True, ""

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MoneyDJ Word Formatter v1.2 — 全自動版
======================================
雙擊 run.command 即可執行。
自動抓取 MoneyDJ 最新國際股市新聞 → 萃取內容 → 產出 Word 檔案。
"""

import os
import re
import sys
import time
import io
from datetime import datetime, timedelta

import streamlit as st
import requests
import urllib3
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

# MoneyDJ 的 SSL 憑證缺少 Subject Key Identifier，停用嚴格驗證
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================================
# 全域設定
# ============================================================================

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = SCRIPT_DIR  # 產出於腳本同目錄
FONT_LATIN = "Times New Roman"
FONT_CJK = "DFKai-SB"  # Windows／Office 使用的標楷體正式名稱

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/127.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

# 9 種新聞分類定義
CATEGORY_PATTERNS = [
    "《美股》", "《美債》", "《歐股》",
    "《陸股》", "《港股》", "《陸港股》",
    "《日股》", "《韓股》", "《日韓股》",
]

# 地區分組規則（合併為 4 大區塊）
REGION_MAPPING = {
    "《美股》": "美國",
    "《美債》": "美國",
    "《歐股》": "歐洲",
    "《陸股》": "陸港股",
    "《港股》": "陸港股",
    "《陸港股》": "陸港股",
    "《日股》": "亞股",
    "《韓股》": "亞股",
    "《日韓股》": "亞股",
}

REGION_ORDER = ["美國", "歐洲", "陸港股", "亞股"]

# MoneyDJ 主題代碼（只用於擴大候選範圍，不作為最終分類依據）
TOPIC_CODES = [
    "X0100009",  # 港股
    "X0100012",  # 日股
    "X0100013",  # 韓股
    "X0100014",  # 美股
    "X0100015",  # 既有候選來源
    "X0100016",  # 既有候選來源
    "X0100017",  # 既有候選來源
    "X0100018",  # 既有候選來源
    "X0200008",  # 歐股候選來源
]

# MoneyDJ 新聞列表頁（保留舊入口，並加入目前站內分類入口）
NEWS_LIST_URLS = [
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=CB010000",
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=CB020000",
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb03",      # 國際股市
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb10",      # 深滬港股
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb030200",  # 亞股
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb030100",  # 美股
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb030300",  # 歐股
    "https://www.moneydj.com/kmdj/news/newsreallist.aspx?a=mb040200",  # 債券市場
]


# ============================================================================
# 模組一：資料索敵與驗證 (Data Scraping & Validation)
# ============================================================================

def log(msg):
    """帶時間戳的日誌輸出。"""
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}")
    if 'log_container' in st.session_state:
        st.session_state.log_container.text(f"[{ts}] {msg}")


def fetch_page(url, timeout=15):
    """安全地抓取網頁，回傳 BeautifulSoup 物件或 None。"""
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
        resp.encoding = "utf-8"
        if resp.status_code == 200:
            return BeautifulSoup(resp.text, "html.parser")
    except Exception as e:
        log(f"  抓取失敗 {url}: {e}")
    return None


def matches_target_title(text):
    """列表標題包含九種目標分類之一時才視為候選文章。"""
    normalized = re.sub(r"\s+", "", text or "")
    return any(pattern in normalized for pattern in CATEGORY_PATTERNS)


def normalize_article_url(href):
    """將 MoneyDJ 文章連結轉成完整 URL。"""
    if not href or "newsviewer.aspx?a=" not in href.lower():
        return None
    if href.startswith("/"):
        return "https://www.moneydj.com" + href
    if not href.startswith("http"):
        return "https://www.moneydj.com/kmdj/news/" + href
    return href


def collect_target_links(soup):
    """從列表頁擷取標題符合目標分類的文章連結。"""
    target_urls = set()
    links = soup.find_all("a", href=re.compile(r"newsviewer\.aspx\?a=", re.IGNORECASE))
    for link in links:
        visible_title = " ".join(filter(None, [
            link.get_text(" ", strip=True),
            link.get("title", ""),
            link.get("aria-label", ""),
        ]))
        if not matches_target_title(visible_title):
            continue
        full_url = normalize_article_url(link.get("href", ""))
        if full_url:
            target_urls.add(full_url)
    return target_urls


def discover_article_urls():
    """從 MoneyDJ 各頁面蒐集標題符合九種分類的文章 URL。"""
    found_urls = set()

    # 策略 1：從主題列表頁抓取
    for code in TOPIC_CODES:
        url = f"https://www.moneydj.com/kmdj/common/listnewarticles.aspx?svc=NW&a={code}"
        log(f"掃描主題頁 {code}...")
        soup = fetch_page(url)
        if soup:
            found_urls.update(collect_target_links(soup))
        time.sleep(0.3)

    # 策略 2：從新聞列表頁抓取
    for list_url in NEWS_LIST_URLS:
        log(f"掃描新聞列表頁 {list_url.rsplit('=', 1)[-1]}...")
        soup = fetch_page(list_url)
        if soup:
            found_urls.update(collect_target_links(soup))
        time.sleep(0.3)

    # 策略 3：從 RSS Feed 抓取
    rss_urls = [
        "https://www.moneydj.com/kmdj/RssCenter.aspx?svc=NW&fno=1&arg=X0000000",
        "https://www.moneydj.com/kmdj/RssCenter.aspx?svc=NW&fno=1&arg=X0100000",
    ]
    for rss_url in rss_urls:
        log("掃描 RSS Feed...")
        try:
            resp = requests.get(rss_url, headers=HEADERS, timeout=10, verify=False)
            resp.encoding = "utf-8"
            if resp.status_code == 200:
                rss_soup = BeautifulSoup(resp.content, "xml")
                for item in rss_soup.find_all("item"):
                    title_tag = item.find("title")
                    link_tag = item.find("link")
                    title = title_tag.get_text(" ", strip=True) if title_tag else ""
                    href = link_tag.get_text(strip=True) if link_tag else ""
                    if matches_target_title(title):
                        full_url = normalize_article_url(href)
                        if full_url:
                            found_urls.add(full_url)
        except Exception as e:
            log(f"  RSS 解析失敗：{e}")
        time.sleep(0.3)

    # 策略 4：從新聞首頁抓取
    log("掃描新聞首頁...")
    soup = fetch_page("https://www.moneydj.com/kmdj/news/newshome.aspx")
    if soup:
        found_urls.update(collect_target_links(soup))

    log(f"共發現 {len(found_urls)} 篇候選文章 URL")
    return list(found_urls)


def parse_article(url):
    """進入文章內頁，解析標題、發布日期、內文段落。
    回傳 dict 或 None。
    """
    soup = fetch_page(url)
    if not soup:
        return None

    # 取得頁面全文字
    page_text = soup.get_text(separator="\n")
    lines = [ln.strip() for ln in page_text.split("\n") if ln.strip()]

    # 解析標題（從 <title> 或 og:title）
    title = ""
    title_tag = soup.find("title")
    if title_tag:
        title = title_tag.get_text().strip()
        # 移除 " - MoneyDJ理財網" 後綴
        title = re.sub(r"\s*-\s*MoneyDJ.*$", "", title).strip()

    # 檢查是否符合目標分類
    category = None
    for pat in CATEGORY_PATTERNS:
        if pat in title:
            category = pat
            break

    if not category:
        return None

    # 解析發布日期（格式：MoneyDJ新聞 2026-08-14 06:14:14 XXX 發佈）
    pub_date = None
    date_pattern = re.compile(r"MoneyDJ新聞\s+(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})")
    for line in lines:
        m = date_pattern.search(line)
        if m:
            try:
                pub_date = datetime.strptime(m.group(1), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                pass
            break

    if not pub_date:
        return None

    # 解析內文（在日期行之後、相關新聞/標籤之前的文字）
    body_lines = []
    in_body = False
    for line in lines:
        if date_pattern.search(line):
            in_body = True
            continue
        if in_body:
            # 停止條件：遇到相關新聞、編者按、廣告、標籤等
            if any(stop in line for stop in [
                "＊編者按", "編者按", "圖片來源", "（圖片來源",
                "MoneyDJ理財網", "財經知識庫", "基金頻道",
                "新聞首頁", "最新頭條", "推薦新聞",
                "首頁", "財經百科", "Back To Top",
                "econsult@", "加入會員", "手機版", "iQuote",
            ]):
                break
            # 跳過過短的行或導航文字
            if len(line) > 15:
                body_lines.append(line)

    if not body_lines:
        return None

    return {
        "url": url,
        "title": title,
        "category": category,
        "region": REGION_MAPPING.get(category, "其他"),
        "pub_date": pub_date,
        "body_lines": body_lines,
    }


def select_latest_articles(all_articles, max_age_days=5):
    """針對每個分類，選取最新的一篇文章。
    若超過 max_age_days 天仍無結果，標示為放棄。
    """
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    cutoff = today - timedelta(days=max_age_days)

    # 按分類分組
    by_category = {}
    for art in all_articles:
        cat = art["category"]
        if cat not in by_category:
            by_category[cat] = []
        by_category[cat].append(art)

    selected = {}
    for cat in CATEGORY_PATTERNS:
        candidates = by_category.get(cat, [])
        # 篩選在時間範圍內的
        valid = [a for a in candidates if a["pub_date"] >= cutoff]
        if valid:
            # 取最新的一篇
            latest = max(valid, key=lambda x: x["pub_date"])
            selected[cat] = latest
            log(f"  {cat} → {latest['title']} ({latest['pub_date'].strftime('%Y-%m-%d %H:%M')})")
        else:
            log(f"  {cat} → 過去 {max_age_days} 日內未找到")

    return selected


def scrape_all_news():
    """模組一主入口：搜尋、抓取、驗證所有新聞。"""
    log("=" * 60)
    log("模組一：資料索敵與驗證")
    log("=" * 60)

    # 1. 發現文章 URL
    urls = discover_article_urls()

    if not urls:
        log("無法從 MoneyDJ 發現任何文章 URL，請確認網路連線。")
        return {}

    # 2. 逐一解析文章
    log(f"\n開始解析 {len(urls)} 篇候選文章...")
    all_articles = []
    for i, url in enumerate(urls):
        art = parse_article(url)
        if art:
            all_articles.append(art)
        # 限制請求速率
        if (i + 1) % 10 == 0:
            log(f"  已解析 {i + 1}/{len(urls)} 篇...")
        time.sleep(0.2)

    log(f"共解析到 {len(all_articles)} 篇符合分類的文章")

    # 3. 選取每個分類的最新文章
    log("\n選取各分類最新文章：")
    selected = select_latest_articles(all_articles)

    return selected


# ============================================================================
# 模組二：內容組織（保留原文完整內容）
# ============================================================================

MARKET_BENCHMARKS = {
    "《美股》": [
        ("dow", ("道瓊工業平均指數", "道瓊工業指數", "道瓊指數", "道瓊")),
        ("nasdaq", ("那斯達克綜合指數", "那斯達克指數", "NASDAQ", "那指")),
        ("sp500", ("標準普爾500指數", "標普500指數", "S&P 500", "標普500")),
        ("sox", ("費城半導體指數", "費半指數", "費半")),
    ],
    "《歐股》": [
        ("stoxx", ("STOXX 600指數", "STOXX 600", "泛歐STOXX")),
        ("dax", ("德國DAX指數", "DAX指數", "DAX")),
        ("ftse", ("英國FTSE 100指數", "FTSE 100指數", "FTSE 100")),
        ("cac", ("法國CAC 40指數", "CAC 40指數", "CAC 40")),
    ],
    "《陸股》": [
        ("shanghai", ("上證綜合指數", "上證指數", "滬指")),
        ("shenzhen", ("深證成份指數", "深證成指", "深成指")),
        ("chinext", ("創業板指數", "創業板指")),
        ("star50", ("科創50指數", "科創50")),
    ],
    "《港股》": [
        ("hsi", ("香港恆生指數", "恆生指數", "恆指")),
        ("hstech", ("恆生科技指數", "恆科指數", "恆科指")),
        ("hscei", ("恆生中國企業指數", "國企指數", "國指")),
    ],
    "《日股》": [
        ("nikkei", ("日經225指數", "日經225", "日經指數", "日經")),
        ("topix", ("東證股價指數", "TOPIX指數", "TOPIX")),
    ],
    "《韓股》": [
        ("kospi", ("韓國KOSPI指數", "KOSPI指數", "KOSPI")),
        ("kosdaq", ("韓國KOSDAQ指數", "KOSDAQ指數", "KOSDAQ")),
    ],
}

# 合併型文章沿用兩個市場的指數字典。
MARKET_BENCHMARKS["《陸港股》"] = (
    MARKET_BENCHMARKS["《陸股》"] + MARKET_BENCHMARKS["《港股》"]
)
MARKET_BENCHMARKS["《日韓股》"] = (
    MARKET_BENCHMARKS["《日股》"] + MARKET_BENCHMARKS["《韓股》"]
)

SUMMARY_PREFIXES = {
    "《美股》": ("美國四大指數", "美股四大指數", "美國主要指數"),
    "《歐股》": ("歐洲主要指數", "歐股主要指數"),
    "《陸股》": ("中國主要指數", "陸股主要指數"),
    "《港股》": ("香港主要指數", "港股主要指數"),
    "《陸港股》": ("陸港股主要指數",),
    "《日股》": ("日本主要指數", "日股主要指數"),
    "《韓股》": ("韓國主要指數", "韓股主要指數"),
    "《日韓股》": ("日韓主要指數",),
}

MOVE_PATTERN = re.compile(r"上漲|下跌|收漲|收跌|走高|走低|勁揚|挫跌|漲|跌|揚|挫")
CLOSE_PATTERN = re.compile(r"收(?:於|在|盤)?|終場|尾盤")
PERCENT_PATTERN = re.compile(r"\d[\d,]*(?:\.\d+)?\s*%")
POINT_PATTERN = re.compile(r"\d[\d,]*(?:\.\d+)?\s*點")
BOND_TERM_PATTERN = re.compile(r"(?<!\d)(2|5|10|20|30)\s*年期")


def _split_sentences(text):
    """依中文句末符號切句並保留標點。"""
    return [
        sentence.strip()
        for sentence in re.split(r"(?<=[。！？!?])", text or "")
        if sentence.strip()
    ]


def _matched_benchmarks(sentence, category):
    """回傳句子中出現的主要指數代號集合。"""
    matched = set()
    for key, aliases in MARKET_BENCHMARKS.get(category, []):
        if any(alias.lower() in sentence.lower() for alias in aliases):
            matched.add(key)
    return matched


def _trim_to_market_data(sentence, category):
    """移除行情資料之前的敘述性引言，保留有意義的市場摘要前綴。"""
    starts = []
    lowered = sentence.lower()
    for _, aliases in MARKET_BENCHMARKS.get(category, []):
        for alias in aliases:
            index = lowered.find(alias.lower())
            if index >= 0:
                starts.append(index)

    for prefix in SUMMARY_PREFIXES.get(category, ()):
        index = sentence.find(prefix)
        if index >= 0:
            starts.append(index)

    return sentence[min(starts):].strip() if starts else sentence.strip()


def _is_equity_market_sentence(sentence, category):
    """判斷句子是否包含完整的主要股價指數收盤資料。"""
    benchmarks = _matched_benchmarks(sentence, category)
    if not benchmarks:
        return False
    has_value = bool(PERCENT_PATTERN.search(sentence) or POINT_PATTERN.search(sentence))
    return (
        has_value
        and bool(MOVE_PATTERN.search(sentence))
        and bool(CLOSE_PATTERN.search(sentence))
    )


def _extract_bond_summary(paragraphs):
    """擷取首次出現的各年期美債殖利率行情句。"""
    summaries = []
    covered_terms = set()
    for paragraph in paragraphs:
        for sentence in _split_sentences(paragraph):
            terms = set(BOND_TERM_PATTERN.findall(sentence))
            if (
                terms
                and "殖利率" in sentence
                and PERCENT_PATTERN.search(sentence)
                and terms - covered_terms
            ):
                cleaned = re.sub(r"（註[:：].*?[。）]$", "", sentence).strip()
                summaries.append(cleaned)
                covered_terms.update(terms)
    return summaries


def extract_market_summaries(category, paragraphs):
    """從全文擷取主要指數／美債殖利率行情；原正文不受影響。"""
    if category == "《美債》":
        return _extract_bond_summary(paragraphs)

    summaries = []
    covered_benchmarks = set()
    for paragraph in paragraphs:
        current_block = []
        for sentence in _split_sentences(paragraph):
            if not _is_equity_market_sentence(sentence, category):
                if current_block:
                    summaries.append("".join(current_block))
                    current_block = []
                continue
            matched = _matched_benchmarks(sentence, category)
            # 每個主要指數只採首次出現的正式行情，避免後文歷史比較重複。
            if not matched - covered_benchmarks:
                if current_block:
                    summaries.append("".join(current_block))
                    current_block = []
                continue
            current_block.append(_trim_to_market_data(sentence, category))
            covered_benchmarks.update(matched)
        if current_block:
            summaries.append("".join(current_block))
    return summaries


def organize_for_word(selected_articles):
    """將選取的文章組織為 Word 渲染所需的資料結構。"""
    regions = {}

    for cat, article in selected_articles.items():
        region = article["region"]
        if region not in regions:
            regions[region] = {
                "region_label": region,
                "articles": [],
            }

        paragraphs = article["body_lines"] or ["（內文待更新）"]
        regions[region]["articles"].append({
            "title": article["title"],
            "date": article["pub_date"].strftime("%Y-%m-%d %H:%M:%S"),
            "market_summaries": extract_market_summaries(cat, paragraphs),
            "paragraphs": paragraphs,
        })

    return regions


# ============================================================================
# 模組三：Word 底層樣式建設 (Style Initialization)
# ============================================================================

def _set_font_xml(rpr_element, font_en=FONT_LATIN, font_ea=FONT_CJK):
    """鎖定中英雙軌字體與語言，避免 Windows 套用主題字型。"""
    rfonts = rpr_element.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = parse_xml(
            f'<w:rFonts {nsdecls("w")} '
            f'w:ascii="{font_en}" w:hAnsi="{font_en}" '
            f'w:cs="{font_en}" w:eastAsia="{font_ea}"/>'
        )
        rpr_element.insert(0, rfonts)
    else:
        rfonts.set(qn("w:ascii"), font_en)
        rfonts.set(qn("w:hAnsi"), font_en)
        rfonts.set(qn("w:cs"), font_en)
        rfonts.set(qn("w:eastAsia"), font_ea)

    # Word 的 Theme 字型優先權可能蓋過上述實體字型，必須明確移除。
    for attr_name in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        attr = qn(f"w:{attr_name}")
        if attr in rfonts.attrib:
            del rfonts.attrib[attr]

    lang = rpr_element.find(qn("w:lang"))
    if lang is None:
        lang = parse_xml(
            f'<w:lang {nsdecls("w")} w:val="en-US" w:eastAsia="zh-TW"/>'
        )
        rpr_element.append(lang)
    else:
        lang.set(qn("w:val"), "en-US")
        lang.set(qn("w:eastAsia"), "zh-TW")


def _set_document_default_fonts(doc):
    """設定文件層級預設值，封住未指定文字回退到 Verdana 的缺口。"""
    styles_element = doc.styles.element
    doc_defaults = styles_element.find(qn("w:docDefaults"))
    if doc_defaults is None:
        doc_defaults = parse_xml(f'<w:docDefaults {nsdecls("w")}/>')
        styles_element.insert(0, doc_defaults)

    rpr_default = doc_defaults.find(qn("w:rPrDefault"))
    if rpr_default is None:
        rpr_default = parse_xml(f'<w:rPrDefault {nsdecls("w")}/>')
        doc_defaults.insert(0, rpr_default)

    rpr = rpr_default.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        rpr_default.append(rpr)
    _set_font_xml(rpr)


def _ensure_char_style(
    doc,
    name,
    color,
    bold,
    size_pt,
    font_en=FONT_LATIN,
    font_ea=FONT_CJK,
):
    """建立或取得 Character Style。"""
    try:
        style = doc.styles[name]
    except KeyError:
        style = doc.styles.add_style(name, 2)  # CHARACTER
        style.base_style = doc.styles["Default Paragraph Font"]

    style.font.size = Pt(size_pt)
    style.font.bold = bold
    style.font.color.rgb = color
    style.font.name = font_en

    rpr = style.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        style.element.append(rpr)
    _set_font_xml(rpr, font_en=font_en, font_ea=font_ea)
    return style


def init_styles(doc):
    """初始化所有樣式。"""
    _set_document_default_fonts(doc)
    _ensure_char_style(doc, "NewsTitle", RGBColor(0, 0, 0xFF), True, 14)
    _ensure_char_style(doc, "NewsBody", RGBColor(0, 0, 0), False, 12)
    _ensure_char_style(
        doc,
        "MarketSummary",
        RGBColor(0xFF, 0, 0),
        True,
        12,
        font_en=FONT_LATIN,
        font_ea=FONT_CJK,
    )

    normal = doc.styles["Normal"]
    normal.font.size = Pt(12)
    normal.font.name = FONT_LATIN
    rpr = normal.element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        normal.element.append(rpr)
    _set_font_xml(rpr)


# ============================================================================
# 模組四：表格佈局與內容渲染 (Table Layout & Rendering)
# ============================================================================

def _zero_spacing(p):
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.0


def _set_cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="0" w:type="dxa"/>')
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(cm * 567)))
    tcW.set(qn("w:type"), "dxa")


def _set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    ))


def _add_styled_run(
    p,
    text,
    style_name,
    doc,
    font_en=FONT_LATIN,
    font_ea=FONT_CJK,
):
    run = p.add_run(text)
    try:
        run.style = doc.styles[style_name]
    except KeyError:
        pass
    run.font.name = font_en
    rpr = run._element.find(qn("w:rPr"))
    if rpr is None:
        rpr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
        run._element.insert(0, rpr)
    _set_font_xml(rpr, font_en=font_en, font_ea=font_ea)
    return run


def _next_numbering_id(numbering, element_name, attribute_name):
    """取得 numbering.xml 中下一個可用的數字 ID。"""
    ids = []
    for element in numbering.findall(qn(f"w:{element_name}")):
        value = element.get(qn(f"w:{attribute_name}"))
        if value is not None:
            try:
                ids.append(int(value))
            except ValueError:
                pass
    return max(ids, default=-1) + 1


def _add_list_definition(doc, list_type):
    """建立單層項目符號或十進位編號格式，回傳 abstractNumId。"""
    numbering = doc.part.numbering_part.element
    abstract_id = _next_numbering_id(
        numbering, "abstractNum", "abstractNumId"
    )

    if list_type == "bullet":
        number_format = "bullet"
        level_text = "•"
        marker_color = "FF0000"
        marker_bold = '<w:b w:val="true"/>'
        left_indent = 420
        hanging_indent = 300
    else:
        number_format = "decimal"
        level_text = "%1."
        marker_color = "000000"
        marker_bold = ""
        left_indent = 540
        hanging_indent = 360

    abstract_num = parse_xml(
        f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="singleLevel"/>'
        f'<w:lvl w:ilvl="0">'
        f'<w:start w:val="1"/>'
        f'<w:numFmt w:val="{number_format}"/>'
        f'<w:lvlText w:val="{level_text}"/>'
        f'<w:lvlJc w:val="left"/>'
        f'<w:pPr>'
        f'<w:tabs><w:tab w:val="num" w:pos="{left_indent}"/></w:tabs>'
        f'<w:ind w:left="{left_indent}" w:hanging="{hanging_indent}"/>'
        f'</w:pPr>'
        f'<w:rPr>'
        f'<w:rFonts w:ascii="{FONT_LATIN}" w:hAnsi="{FONT_LATIN}" '
        f'w:cs="{FONT_LATIN}" w:eastAsia="{FONT_CJK}"/>'
        f'{marker_bold}'
        f'<w:color w:val="{marker_color}"/>'
        f'<w:sz w:val="24"/><w:szCs w:val="24"/>'
        f'<w:lang w:val="en-US" w:eastAsia="zh-TW"/>'
        f'</w:rPr>'
        f'</w:lvl>'
        f'</w:abstractNum>'
    )

    # abstractNum 必須排在所有 num 元素之前。
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(abstract_num)
    else:
        first_num.addprevious(abstract_num)
    return abstract_id


def _new_list_instance(doc, abstract_id, restart_at_one=False):
    """建立獨立的 Word 清單實例；正文可藉此確實從 1 重新開始。"""
    numbering = doc.part.numbering_part.element
    num_id = _next_numbering_id(numbering, "num", "numId")
    start_override = (
        '<w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/>'
        '</w:lvlOverride>'
        if restart_at_one
        else ""
    )
    numbering.append(parse_xml(
        f'<w:num {nsdecls("w")} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/>'
        f'{start_override}'
        f'</w:num>'
    ))
    return num_id


def _apply_list_numbering(paragraph, num_id):
    """把段落連結到指定的 Word 清單實例。"""
    ppr = paragraph._p.get_or_add_pPr()
    old_num_pr = ppr.find(qn("w:numPr"))
    if old_num_pr is not None:
        ppr.remove(old_num_pr)
    num_pr = parse_xml(
        f'<w:numPr {nsdecls("w")}>'
        f'<w:ilvl w:val="0"/>'
        f'<w:numId w:val="{num_id}"/>'
        f'</w:numPr>'
    )
    ppr.append(num_pr)


def _split_summary_list_items(summary_text):
    """紅色摘要遇到中文分號或句號時，切成獨立的清單段落。"""
    return [
        item.strip()
        for item in re.split(r"(?<=[；。])", summary_text or "")
        if item.strip()
    ]


def build_tables(doc, regions_data):
    """以單一表格渲染四大地區及其新聞。"""
    bullet_abstract_id = _add_list_definition(doc, "bullet")
    number_abstract_id = _add_list_definition(doc, "number")
    bullet_num_id = _new_list_instance(doc, bullet_abstract_id)

    table = doc.add_table(rows=1, cols=2)
    table.style = doc.styles["Table Grid"]
    table.autofit = False

    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        table._tbl.insert(0, tblPr)
    tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
    tw = tblPr.find(qn("w:tblW"))
    if tw is None:
        tw = parse_xml(f'<w:tblW {nsdecls("w")} w:w="0" w:type="dxa"/>')
        tblPr.append(tw)
    tw.set(qn("w:w"), str(int(18.6 * 567)))
    tw.set(qn("w:type"), "dxa")

    old_grid = table._tbl.find(qn("w:tblGrid"))
    if old_grid is not None:
        table._tbl.remove(old_grid)
    grid = parse_xml(
        f'<w:tblGrid {nsdecls("w")}>'
        f'<w:gridCol w:w="{int(1.6 * 567)}"/>'
        f'<w:gridCol w:w="{int(17 * 567)}"/>'
        f'</w:tblGrid>'
    )
    tblPr.addnext(grid)

    header = table.rows[0]
    trPr = header._tr.get_or_add_trPr()
    trPr.append(parse_xml(
        f'<w:tblHeader {nsdecls("w")} w:val="true"/>'
    ))
    for index, (label, width) in enumerate((("分類", 1.6), ("盤後新聞", 17))):
        cell = header.cells[index]
        _set_cell_bg(cell, "92D050")
        _set_cell_width(cell, width)
        p = cell.paragraphs[0]
        p.style = doc.styles["Normal"]
        _zero_spacing(p)
        run = _add_styled_run(p, label, "NewsBody", doc)
        run.bold = True

    for region_key in REGION_ORDER:
        rd = regions_data.get(region_key)
        if not rd:
            continue

        row = table.add_row()
        left, right = row.cells
        _set_cell_width(left, 1.6)
        _set_cell_width(right, 17)

        left_p = left.paragraphs[0]
        left_p.style = doc.styles["Normal"]
        _zero_spacing(left_p)
        _add_styled_run(left_p, region_key, "NewsBody", doc)

        first_p = right.paragraphs[0]
        is_first = True
        for art in rd["articles"]:
            # 每篇新聞使用不同 numId，避免 Word 自動接續上一篇的編號。
            body_num_id = _new_list_instance(
                doc, number_abstract_id, restart_at_one=True
            )
            if not is_first:
                separator = right.add_paragraph()
                separator.style = doc.styles["Normal"]
                _zero_spacing(separator)

            title_p = first_p if is_first else right.add_paragraph()
            is_first = False
            title_p.style = doc.styles["Normal"]
            _zero_spacing(title_p)
            _add_styled_run(title_p, art["title"], "NewsTitle", doc)

            date_p = right.add_paragraph()
            date_p.style = doc.styles["Normal"]
            _zero_spacing(date_p)
            _add_styled_run(
                date_p,
                f'MoneyDJ新聞 {art["date"]} 發佈',
                "NewsBody",
                doc,
            )

            # 將主要指數／美債殖利率行情複製到時間後方並醒目標示。
            for summary_text in art["market_summaries"]:
                for item_text in _split_summary_list_items(summary_text):
                    summary_p = right.add_paragraph()
                    summary_p.style = doc.styles["Normal"]
                    _zero_spacing(summary_p)
                    _apply_list_numbering(summary_p, bullet_num_id)
                    _add_styled_run(
                        summary_p,
                        item_text,
                        "MarketSummary",
                        doc,
                        font_en=FONT_LATIN,
                        font_ea=FONT_CJK,
                    )

            for para_text in art["paragraphs"]:
                body_p = right.add_paragraph()
                body_p.style = doc.styles["Normal"]
                _zero_spacing(body_p)
                _apply_list_numbering(body_p, body_num_id)
                _add_styled_run(body_p, para_text, "NewsBody", doc)


# ============================================================================
# 主程式
# ============================================================================

def generate_report():
    log("=" * 60)
    log("MoneyDJ Word Formatter (Web App)")
    log("=" * 60)

    selected = scrape_all_news()
    if not selected:
        log("\n未抓取到任何符合條件的文章。可能原因：網路連線問題或 MoneyDJ 頁面結構變更。")
        return None, None

    log("\n" + "=" * 60)
    log("模組二：內容組織")
    log("=" * 60)
    regions_data = organize_for_word(selected)
    log(f"已組織 {len(regions_data)} 個區塊")

    log("\n" + "=" * 60)
    log("模組三/四：Word 樣式建設 + 表格渲染")
    log("=" * 60)

    doc = Document()
    for sec in doc.sections:
        sec.left_margin = Cm(1.5)
        sec.right_margin = Cm(1.5)
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)

    init_styles(doc)
    build_tables(doc, regions_data)

    today_str = datetime.now().strftime("%Y%m%d")
    filename = f"MoneyDJ_盤後新聞_{today_str}.docx"
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    log(f"\nWord 檔案已產出至記憶體：{filename}")
    log("=" * 60)
    return doc_io, filename


def main():
    st.set_page_config(page_title="MoneyDJ 報表產生器", page_icon="📈")
    st.title("MoneyDJ 盤後新聞報表產生器")
    st.write("點擊下方按鈕，系統將自動從 MoneyDJ 網站抓取最新國際股市新聞並整理成 Word 檔案。")

    if st.button("開始產出報表", type="primary"):
        st.session_state.log_container = st.empty()
        
        with st.spinner('報表產出中，請稍候...'):
            doc_io, filename = generate_report()
            
            if doc_io:
                st.success("產出成功！請點擊下方按鈕下載檔案。")
                st.download_button(
                    label="📥 下載 Word 報表",
                    data=doc_io,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            else:
                st.error("產出失敗，請查看上方日誌了解詳細原因。")

if __name__ == "__main__":
    main()

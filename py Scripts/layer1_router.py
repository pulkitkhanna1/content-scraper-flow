#!/usr/bin/env python3
"""
layer1_router.py — Content Scraper: Layer 1

Given any input URL from supported novel/manhwa platforms, this script:
  1. Detects the platform and its language (ZH / KR / EN / etc.)
  2. Extracts metadata: title, author, chapter count, description, cover
  3. If source is Chinese/Korean → searches for existing English translation
  4. Returns a structured JSON payload for N8N (or CLI use) to drive Layer 2

Usage (CLI):
    python layer1_router.py --url "https://www.webnovel.com/book/1234/"
    python layer1_router.py --url "https://www.69shuba.com/book/12345/"
    python layer1_router.py --url "https://novelbin.com/b/my-book/"

Output (JSON to stdout):
    {
        "status": "ok",
        "platform": "webnovel",
        "source_language": "en",
        "needs_translation": false,
        "book_name": "...",
        "author": "...",
        "chapter_count": 500,
        "description": "...",
        "cover_url": "...",
        "canonical_url": "...",
        "scraper_script": "webnovel_content_uc.py",
        "scraper_args": { ... },
        "translation_candidates": []
    }
"""

import argparse
import json
import re
import sys
import time
from typing import Optional
from urllib.parse import urlparse, urljoin, quote_plus

import requests
from bs4 import BeautifulSoup

# ---------------------------------------------------------------------------
# Platform Registry
# ---------------------------------------------------------------------------
# Each entry: domain fragment → platform config
# source_language: "en" | "zh" | "ko" | "ja"
# needs_translation: True when source is not English
# scraper_script: relative path to the scraper from py Scripts/
# needs_browser: True → UC Chrome required; False → requests-only

PLATFORMS = {
    "webnovel.com": {
        "platform": "webnovel",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "webnovel_content_uc.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "69shuba.com": {
        "platform": "69shuba",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "69shuba_new.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "69shu.com": {
        "platform": "69shuba",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "69shuba_new.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "novelbin.com": {
        "platform": "novelbin",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "novelbin.py",
        "needs_browser": False,
        "needs_login": False,
    },
    "freewebnovel.com": {
        "platform": "freewebnovel",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "freewebnovel/script_new.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "babelnovel.com": {
        "platform": "babelnovel",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "babelnovel_content.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "royalroad.com": {
        "platform": "royalroad",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "royalroad_content.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "tapas.io": {
        "platform": "tapas",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "tapas_content_fixed.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "wuxiaworld.com": {
        "platform": "wuxiaworld",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "wuxiaworld_next.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "wattpad.com": {
        "platform": "wattpad",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": "wattpad_content.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "kakao.com": {
        "platform": "kakao",
        "source_language": "ko",
        "needs_translation": True,
        "scraper_script": "kakao_content.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "kakaopage.com": {
        "platform": "kakao",
        "source_language": "ko",
        "needs_translation": True,
        "scraper_script": "kakao_content.py",
        "needs_browser": True,
        "needs_login": True,
    },
    "qidian.com": {
        "platform": "qidian",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "qdmm_content_new.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "qdmm.com": {
        "platform": "qdmm",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "qdmm_content_new.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "hengyan.com": {
        "platform": "hengyan",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "hengyan/s.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "ihengyan.com": {
        "platform": "hengyan",
        "source_language": "zh",
        "needs_translation": True,
        "scraper_script": "hengyan/s.py",
        "needs_browser": True,
        "needs_login": False,
    },
    "novelupdates.com": {
        "platform": "novelupdates",
        "source_language": "en",
        "needs_translation": False,
        "scraper_script": None,  # index site only
        "needs_browser": False,
        "notes": "Use as translation-search source only, not for scraping.",
    },
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

# ---------------------------------------------------------------------------
# Platform Detection
# ---------------------------------------------------------------------------

def detect_platform(url: str) -> Optional[dict]:
    """Return platform config dict for the given URL, or None if unknown."""
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    # Strip exactly the "www." prefix (not character-by-character)
    if host.startswith("www."):
        host = host[4:]
    for domain, config in PLATFORMS.items():
        if host == domain or host.endswith("." + domain):
            result = dict(config)
            result["canonical_url"] = url
            return result
    return None


# ---------------------------------------------------------------------------
# Metadata Extractors (per platform, requests + BS4 where possible)
# ---------------------------------------------------------------------------

def _fetch_html(url: str, timeout: int = 15) -> Optional[BeautifulSoup]:
    try:
        resp = requests.get(url, headers=HEADERS, timeout=timeout)
        resp.raise_for_status()
        return BeautifulSoup(resp.content, "html.parser")
    except Exception as e:
        return None


def _text(soup, *selectors) -> str:
    """Try multiple CSS selectors, return first non-empty text found."""
    for sel in selectors:
        el = soup.select_one(sel)
        if el:
            t = el.get_text(strip=True)
            if t:
                return t
    return ""


def extract_metadata_webnovel(url: str) -> dict:
    """webnovel.com — uses their JSON API (no browser needed for metadata)."""
    parsed = urlparse(url)
    parts = parsed.path.strip("/").split("/")
    # Path formats: /book/{bookId} or /book/{bookId}/{chapterId}
    book_id = None
    for i, part in enumerate(parts):
        if part == "book" and i + 1 < len(parts):
            bid = parts[i + 1]
            # May be numeric or "slug_12345" style
            m = re.search(r"(\d{10,})", bid)
            if m:
                book_id = m.group(1)
            elif bid.isdigit():
                book_id = bid
            break

    if not book_id:
        return {}

    api_url = f"https://www.webnovel.com/go/pcm/bookInfo/getBookWeb?bookId={book_id}"
    try:
        resp = requests.get(api_url, headers=HEADERS, timeout=15)
        data = resp.json()
        book = data.get("data", {}).get("bookInfo", {})
        return {
            "book_name": book.get("bookName", ""),
            "author": book.get("authorName", ""),
            "chapter_count": book.get("chapterCount", 0),
            "description": book.get("description", ""),
            "cover_url": book.get("coverUrl", ""),
            "book_id": book_id,
            "scraper_args": {
                "book_url": f"https://www.webnovel.com/book/{book_id}",
                "start_chapter": 1,
                "end_chapter": book.get("chapterCount", 100),
                "out_dir": book.get("bookName", f"webnovel_{book_id}"),
            },
        }
    except Exception:
        pass

    # Fallback: HTML scrape
    soup = _fetch_html(f"https://www.webnovel.com/book/{book_id}")
    if not soup:
        return {}
    return {
        "book_name": _text(soup, "h1.pt4", "h1", "title"),
        "author": _text(soup, ".ell.c_s", "[class*='author']"),
        "chapter_count": 0,
        "description": _text(soup, ".g_txt_2.br", "[class*='desc']"),
        "cover_url": "",
        "book_id": book_id,
        "scraper_args": {
            "book_url": f"https://www.webnovel.com/book/{book_id}",
            "start_chapter": 1,
            "end_chapter": 100,
            "out_dir": "",
        },
    }


def extract_metadata_69shuba(url: str) -> dict:
    """69shuba.com — HTML scrape of book info page."""
    # URL formats: /book/{id}/ or /txt/{id}/{chapterId}
    parsed = urlparse(url)
    path = parsed.path.strip("/")
    m = re.search(r"(?:book|txt)/(\d+)", path)
    book_id = m.group(1) if m else None

    book_url = f"https://www.69shuba.com/book/{book_id}/" if book_id else url
    soup = _fetch_html(book_url)
    if not soup:
        return {}

    title = _text(soup, "h1.booktitle", ".bookname h1", "h1")
    author = _text(soup, ".booknav2 p a", ".author a", "[class*='author'] a")
    desc = _text(soup, ".intro", "[class*='intro']", "#intro")

    # Chapter count from chapter list length or text mention
    chapter_count = 0
    count_el = soup.select_one(".chapter-count, [class*='chapter'] span")
    if count_el:
        m2 = re.search(r"(\d+)", count_el.get_text())
        if m2:
            chapter_count = int(m2.group(1))

    # First chapter URL for the scraper
    first_chapter_url = ""
    first_link = soup.select_one(".chapter-list a, .catalog a, ul.chapter-list li:first-child a")
    if first_link:
        href = first_link.get("href", "")
        first_chapter_url = href if href.startswith("http") else urljoin(book_url, href)

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": "",
        "book_id": book_id,
        "scraper_args": {
            "start_url": first_chapter_url or url,
            "start_chapter": 1,
            "end_chapter": chapter_count or 100,
            "book_name": title,
        },
    }


def extract_metadata_novelbin(url: str) -> dict:
    """novelbin.com — HTML scrape."""
    soup = _fetch_html(url)
    if not soup:
        return {}

    title = _text(soup, "h3.title", "h1.title", ".book-info h3")
    author = _text(soup, ".author a", "[class*='author'] a", ".info-item a")
    desc = _text(soup, "#tab-description p", ".desc-text p", ".summary p")
    cover = ""
    img = soup.select_one(".book-img img, .cover img")
    if img:
        cover = img.get("src", "")

    # Chapter count
    chapter_count = 0
    ch_count_el = soup.select_one(".chapter-count, .l-chapter a")
    if ch_count_el:
        m = re.search(r"(\d+)", ch_count_el.get_text())
        if m:
            chapter_count = int(m.group(1))

    # First chapter URL
    first_link = soup.select_one(".l-chapter a, .chapter-list a")
    first_url = ""
    if first_link:
        href = first_link.get("href", "")
        first_url = href if href.startswith("http") else urljoin("https://novelbin.com", href)

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": cover,
        "scraper_args": {
            "start_url": first_url or url,
            "start_chapter": 1,
            "end_chapter": chapter_count or 100,
            "book_name": title,
        },
    }


def extract_metadata_royalroad(url: str) -> dict:
    """royalroad.com — HTML scrape."""
    # Normalize to fiction page (not chapter page)
    parsed = urlparse(url)
    path = parsed.path
    m = re.search(r"/fiction/(\d+)", path)
    fiction_id = m.group(1) if m else None

    base_url = url
    if fiction_id:
        slug_match = re.search(r"/fiction/\d+/([^/]+)", path)
        slug = slug_match.group(1) if slug_match else "book"
        base_url = f"https://www.royalroad.com/fiction/{fiction_id}/{slug}"

    soup = _fetch_html(base_url)
    if not soup:
        return {}

    title = _text(soup, ".fic-header h1", "h1.font-white", "h1")
    author = _text(soup, ".fic-header a.a-white", ".author-name a", "[property='author']")
    desc = _text(soup, ".description .hidden-content", ".description p")
    cover = ""
    img = soup.select_one(".thumbnail img, .cover img, .fic-image img")
    if img:
        cover = img.get("src", "")

    # Chapter count — count rows in the chapter table (most reliable)
    chapter_rows = soup.select("table#chapters tbody tr")
    chapter_count = len(chapter_rows)

    # Fallback: stats widget text e.g. "351 Chapters"
    if chapter_count == 0:
        for li in soup.select(".stats-content li, .fiction-stats li"):
            m2 = re.search(r"(\d[\d,]*)\s+[Cc]hapter", li.get_text())
            if m2:
                chapter_count = int(m2.group(1).replace(",", ""))
                break

    # Fallback: fa-list icon sibling span
    if chapter_count == 0:
        ch_el = soup.select_one(".fa-list + span, .fa-list ~ span")
        if ch_el:
            m2 = re.search(r"(\d+)", ch_el.get_text())
            if m2:
                chapter_count = int(m2.group(1))

    # First chapter URL — prefer "Start Reading" button, then first table row
    first_url = ""
    start_btn = soup.select_one("a.btn-read-now")
    if start_btn:
        href = start_btn.get("href", "")
        first_url = href if href.startswith("http") else f"https://www.royalroad.com{href}"

    # Find chapter 1 link from the table — RoyalRoad lists oldest-first by default
    if not first_url and chapter_rows:
        link = chapter_rows[0].select_one("td a[href*='/chapter/']")
        if link:
            href = link.get("href", "")
            first_url = href if href.startswith("http") else f"https://www.royalroad.com{href}"

    # Fallback: any chapter link on the page
    if not first_url:
        link = soup.select_one("a[href*='/chapter/']")
        if link:
            href = link.get("href", "")
            first_url = href if href.startswith("http") else f"https://www.royalroad.com{href}"

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": cover,
        "fiction_id": fiction_id,
        "scraper_args": {
            "url": first_url or url,
            "title": title,
        },
    }


def extract_metadata_freewebnovel(url: str) -> dict:
    """freewebnovel.com — HTML scrape."""
    # Normalize to book main page
    parsed = urlparse(url)
    path = parsed.path
    m = re.search(r"/novel/([^/]+)", path)
    slug = m.group(1) if m else None
    book_url = f"https://freewebnovel.com/novel/{slug}" if slug else url

    soup = _fetch_html(book_url)
    if not soup:
        return {}

    title = _text(soup, "h1.tit", ".book-name h1", "h1")
    author = _text(soup, ".author a", "[class*='author'] a")
    desc = _text(soup, ".sum p", ".intro p", ".description p")
    cover = ""
    img = soup.select_one(".pic img, .cover img")
    if img:
        cover = img.get("src", "")

    chapter_count = 0
    ch_el = soup.select_one(".chapter strong, .info-chapter strong")
    if ch_el:
        m2 = re.search(r"(\d+)", ch_el.get_text())
        if m2:
            chapter_count = int(m2.group(1))

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": cover,
        "scraper_args": {
            "book_url": book_url,
            "start_chapter": 1,
            "end_chapter": chapter_count or 100,
            "book_name": title,
            "url_file": None,
        },
    }


def extract_metadata_wuxiaworld(url: str) -> dict:
    """wuxiaworld.com — HTML scrape."""
    soup = _fetch_html(url)
    if not soup:
        return {}

    title = _text(soup, "h1.novel-title", ".novel-info h1", "h1")
    author = _text(soup, ".author a", "[class*='author']")
    desc = _text(soup, ".review-content p", ".novel-synopsis p", ".description p")

    # Extract book slug and first chapter URL
    import re as _re
    slug_m = _re.search(r"/novel/([^/?#]+)", url)
    slug = slug_m.group(1) if slug_m else None
    first_chapter_url = ""
    chapter_count = 0
    if soup and slug:
        ch_links = _re.findall(rf'href="(/novel/{slug}/{slug}-chapter-(\d+))"', str(soup))
        if ch_links:
            ch_links_sorted = sorted(set(ch_links), key=lambda x: int(x[1]))
            first_chapter_url = f"https://www.wuxiaworld.com{ch_links_sorted[0][0]}"

    # Chapter count from page text (more reliable than link counting)
    ct_m = _re.search(r"(\d{3,5})\s*[Cc]hap", soup.get_text() if soup else "")
    if ct_m:
        chapter_count = int(ct_m.group(1))

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": "",
        "scraper_args": {
            "url": url,
            "slug": slug,
            "first_chapter_url": first_chapter_url,
        },
    }


def extract_metadata_tapas(url: str) -> dict:
    """tapas.io — HTML scrape of series/info page."""
    # Normalize to /info page
    parsed = urlparse(url)
    path = parsed.path
    m = re.search(r"/series/([^/]+)", path)
    slug = m.group(1) if m else None
    info_url = f"https://tapas.io/series/{slug}/info" if slug else url

    soup = _fetch_html(info_url)
    if not soup:
        return {}

    title = _text(soup, ".series__title", "h1.title", "h1")
    author = _text(soup, ".creator-info__name", ".author-name")
    desc = _text(soup, ".series-info__synopsis p", ".synopsis p")

    return {
        "book_name": title,
        "author": author,
        "chapter_count": 0,
        "description": desc,
        "cover_url": "",
        "scraper_args": {
            "start_url": info_url,
            "book_name": title,
            "max_chapters": 9999,
        },
    }


def extract_metadata_wattpad(url: str) -> dict:
    """wattpad.com — HTML scrape."""
    soup = _fetch_html(url)
    if not soup:
        return {}

    title = _text(soup, ".story-info__title", "h1", "[class*='story-title']")
    author = _text(soup, ".author-info__username a", ".username a")
    desc = _text(soup, ".description-text", "[class*='description'] p")

    return {
        "book_name": title,
        "author": author,
        "chapter_count": 0,
        "description": desc,
        "cover_url": "",
        "scraper_args": {
            "url": url,
        },
    }


def extract_metadata_qidian(url: str) -> dict:
    """qidian.com — HTML scrape (Chinese)."""
    soup = _fetch_html(url)
    if not soup:
        return {}

    title = _text(soup, "h1.book-title", ".book-title", "h1")
    author = _text(soup, ".author-name a", ".book-info .author a")
    desc = _text(soup, ".book-intro p", "#book-intro-detail p")

    chapter_count = 0
    ch_el = soup.select_one(".count-number")
    if ch_el:
        m = re.search(r"(\d+)", ch_el.get_text())
        if m:
            chapter_count = int(m.group(1))

    return {
        "book_name": title,
        "author": author,
        "chapter_count": chapter_count,
        "description": desc,
        "cover_url": "",
        "scraper_args": {
            "book_url": url,
            "book_name": title,
            "start_chapter": 1,
            "end_chapter": chapter_count or 100,
        },
    }


def extract_metadata_generic(url: str) -> dict:
    """Fallback generic metadata extractor."""
    soup = _fetch_html(url)
    if not soup:
        return {}

    title = (
        _text(soup, "h1", "title")
        or (soup.title.get_text(strip=True) if soup.title else "")
    )
    desc = ""
    meta_desc = soup.find("meta", attrs={"name": "description"})
    if meta_desc:
        desc = meta_desc.get("content", "")

    return {
        "book_name": title,
        "author": "",
        "chapter_count": 0,
        "description": desc,
        "cover_url": "",
        "scraper_args": {"url": url},
    }


# Map platform name → extractor function
METADATA_EXTRACTORS = {
    "webnovel":     extract_metadata_webnovel,
    "69shuba":      extract_metadata_69shuba,
    "novelbin":     extract_metadata_novelbin,
    "royalroad":    extract_metadata_royalroad,
    "freewebnovel": extract_metadata_freewebnovel,
    "wuxiaworld":   extract_metadata_wuxiaworld,
    "tapas":        extract_metadata_tapas,
    "wattpad":      extract_metadata_wattpad,
    "qidian":       extract_metadata_qidian,
    "qdmm":         extract_metadata_qidian,  # same structure
    "babelnovel":   extract_metadata_generic,
    "kakao":        extract_metadata_generic,
    "hengyan":      extract_metadata_generic,
}


# ---------------------------------------------------------------------------
# Translation Search (for ZH/KR sources)
# ---------------------------------------------------------------------------

TRANSLATION_SEARCH_SOURCES = [
    "novelupdates.com",
    "novelbin.com",
    "freewebnovel.com",
    "webnovel.com",
    "wuxiaworld.com",
    "scribblehub.com",
]


def search_novelupdates(title: str) -> list:
    """Search novelupdates.com for English translation of a Chinese/Korean title."""
    results = []
    search_url = f"https://www.novelupdates.com/?s={quote_plus(title)}&post_type=seriesplans"
    soup = _fetch_html(search_url)
    if not soup:
        return results

    for item in soup.select(".search_main_box_nu")[:5]:
        link = item.select_one("a.w-blog-entry-title-h")
        if not link:
            link = item.select_one("a")
        if link:
            href = link.get("href", "")
            text = link.get_text(strip=True)
            if href and text:
                results.append({
                    "source": "novelupdates",
                    "title": text,
                    "url": href,
                    "note": "Index page — follow links for translation sites",
                })
    return results


def search_freewebnovel(title: str) -> list:
    """Search freewebnovel.com for an English translation."""
    results = []
    search_url = f"https://freewebnovel.com/search/?searchkey={quote_plus(title)}"
    soup = _fetch_html(search_url)
    if not soup:
        return results

    for item in soup.select(".con-list li, .search-item")[:3]:
        link = item.select_one("a")
        if link:
            href = link.get("href", "")
            text = link.get_text(strip=True)
            if href and text:
                results.append({
                    "source": "freewebnovel",
                    "title": text,
                    "url": href if href.startswith("http") else f"https://freewebnovel.com{href}",
                })
    return results


def search_novelbin(title: str) -> list:
    """Search novelbin.com for an English translation."""
    results = []
    search_url = f"https://novelbin.com/?search={quote_plus(title)}"
    soup = _fetch_html(search_url)
    if not soup:
        return results

    for item in soup.select(".list-truyen .row, .search-result .item")[:3]:
        link = item.select_one("h3.truyen-title a, .title a")
        if link:
            href = link.get("href", "")
            text = link.get_text(strip=True)
            if href and text:
                results.append({
                    "source": "novelbin",
                    "title": text,
                    "url": href if href.startswith("http") else f"https://novelbin.com{href}",
                })
    return results


def find_translation_candidates(title: str, author: str = "") -> list:
    """Search multiple sites for English translation of a Chinese/Korean source."""
    if not title:
        return []

    candidates = []
    query = title

    # Try novelupdates first (most comprehensive index)
    try:
        candidates.extend(search_novelupdates(query))
        time.sleep(0.5)
    except Exception:
        pass

    # Try freewebnovel
    try:
        candidates.extend(search_freewebnovel(query))
        time.sleep(0.5)
    except Exception:
        pass

    # Try novelbin
    try:
        candidates.extend(search_novelbin(query))
        time.sleep(0.5)
    except Exception:
        pass

    return candidates


# ---------------------------------------------------------------------------
# Chapter Numbering Config (for the scraper invocation)
# ---------------------------------------------------------------------------

CHAPTER_BOUNDARY_PATTERNS = [
    r"(?i)^chapter\s*(\d+)",
    r"(?i)^ch(?:\.|\s*)\s*(\d+)",
    r"(?i)^ep(?:isode)?\.?\s*(\d+)",
    r"(?i)^episode\s*(\d+)",
    r"^第\s*(\d+)\s*[章集节话]",   # Chinese
    r"^(\d+)[\.:\-\s]",            # bare number prefix
]


# ---------------------------------------------------------------------------
# Main Router
# ---------------------------------------------------------------------------

def route(url: str, start_chapter: int = 1, end_chapter: Optional[int] = None,
          search_translations: bool = True) -> dict:
    """
    Full Layer 1 routing for a given URL.

    Returns a dict suitable for JSON serialisation and N8N consumption.
    """
    url = url.strip()

    # 1. Detect platform
    platform_config = detect_platform(url)
    if not platform_config:
        return {
            "status": "error",
            "error": f"Unknown platform for URL: {url}",
            "url": url,
            "platform": None,
            "source_language": None,
            "needs_translation": None,
            "translation_candidates": [],
        }

    platform_name = platform_config["platform"]

    # 2. Extract metadata
    extractor = METADATA_EXTRACTORS.get(platform_name, extract_metadata_generic)
    try:
        meta = extractor(url)
    except Exception as e:
        meta = {"book_name": "", "author": "", "chapter_count": 0,
                "description": "", "cover_url": "", "scraper_args": {"url": url}}
        meta["metadata_error"] = str(e)

    # 3. Merge into result
    result = {
        "status": "ok",
        "platform": platform_name,
        "source_language": platform_config["source_language"],
        "needs_translation": platform_config["needs_translation"],
        "needs_browser": platform_config.get("needs_browser", True),
        "needs_login": platform_config.get("needs_login", False),
        "scraper_script": platform_config.get("scraper_script"),
        "canonical_url": url,
        "book_name": meta.get("book_name", ""),
        "author": meta.get("author", ""),
        "chapter_count": meta.get("chapter_count", 0),
        "description": meta.get("description", ""),
        "cover_url": meta.get("cover_url", ""),
        "scraper_args": meta.get("scraper_args", {}),
        "translation_candidates": [],
        "chapter_batch_size": 100,
        "chapter_split_patterns": CHAPTER_BOUNDARY_PATTERNS,
    }

    # Apply user-supplied chapter range override
    if "scraper_args" in result:
        if start_chapter and start_chapter > 1:
            result["scraper_args"]["start_chapter"] = start_chapter
        if end_chapter:
            result["scraper_args"]["end_chapter"] = end_chapter
        elif result["chapter_count"] and not result["scraper_args"].get("end_chapter"):
            result["scraper_args"]["end_chapter"] = result["chapter_count"]

    # 4. If source is not English, search for existing English translation
    if platform_config["needs_translation"] and search_translations and result["book_name"]:
        result["translation_candidates"] = find_translation_candidates(
            result["book_name"], result["author"]
        )
        # Recommend action
        if result["translation_candidates"]:
            result["recommended_action"] = "use_existing_translation"
            result["recommended_url"] = result["translation_candidates"][0]["url"]
        else:
            result["recommended_action"] = "translate_from_source"
            result["recommended_url"] = url
    else:
        result["recommended_action"] = "scrape_direct"
        result["recommended_url"] = url

    return result


# ---------------------------------------------------------------------------
# CLI Entry Point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Layer 1 Router: Detect platform, extract metadata, find translations."
    )
    parser.add_argument("--url", required=True, help="Novel/manhwa URL to process")
    parser.add_argument("--start-chapter", type=int, default=1,
                        help="Starting chapter (default: 1)")
    parser.add_argument("--end-chapter", type=int, default=None,
                        help="Ending chapter (default: auto from metadata)")
    parser.add_argument("--no-translation-search", action="store_true",
                        help="Skip searching for English translations")
    parser.add_argument("--pretty", action="store_true",
                        help="Pretty-print JSON output")
    args = parser.parse_args()

    result = route(
        url=args.url,
        start_chapter=args.start_chapter,
        end_chapter=args.end_chapter,
        search_translations=not args.no_translation_search,
    )

    indent = 2 if args.pretty else None
    print(json.dumps(result, ensure_ascii=False, indent=indent))
    sys.exit(0 if result.get("status") == "ok" else 1)


if __name__ == "__main__":
    main()

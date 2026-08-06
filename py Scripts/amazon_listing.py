#!/usr/bin/env python3
"""
Amazon Kindle & Audiobook Scraper
Scrapes Amazon bestsellers pages and extracts Kindle and Audiobook product details
"""

import csv
import time
import re
import logging
import os
from urllib.parse import urljoin, urlparse, parse_qs
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    InvalidSelectorException,
    WebDriverException,
    NoSuchWindowException
)
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import json

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# ── Bulk run configuration ─────────────────────────────────────────────────────
# URLs are loaded from amazon_uk.csv at startup, grouped by Genre. Pass
# --genre "<Genre Name>" to scrape a single Genre in its own terminal, or
# --list-genres to see the available options.

DEFAULT_BULK_CSV = 'amazon_uk.csv'
# Each Genre/Sub-Genre row appears under several Node Types (category,
# ku-eligible, top-100-paid-kindle, top-100-free-kindle). We use the
# top-100-paid-kindle row as the base URL — the scraper itself derives the
# Free list URL from the same base.
DEFAULT_NODE_TYPE = 'top-100-paid-kindle'


def _genre_slug(name):
    """Convert a Genre/label into a filesystem-safe slug."""
    return re.sub(r'[^A-Za-z0-9]+', '_', name or '').strip('_') or 'unknown'


def _bestseller_slug(url, sub_genre='', genre=''):
    """Build a unique, readable per-URL filename slug.

    The amazon.co.uk bestseller URLs all share the same leading path segment
    (`gp`), so the old `path.split('/')[0]` approach collapsed every URL in a
    genre into the same file. We combine the Sub-Genre hierarchy label (unique
    within a Genre) with the numeric Amazon node ID from the URL — readable
    *and* collision-proof.
    """
    label = _genre_slug(sub_genre or genre or 'list')
    node_match = re.search(r'/digital-text/(\d+)', url or '')
    if node_match:
        return f"{label}_{node_match.group(1)}"
    return label


def _build_sub_genre_label(row):
    """Build a readable sub-genre label from the CSV hierarchy columns."""
    parts = [
        (row.get('Sub-Genre') or '').strip(),
        (row.get('Sub-Sub') or '').strip(),
        (row.get('Sub-Sub-Sub') or '').strip(),
    ]
    parts = [p for p in parts if p]
    return ' > '.join(parts) if parts else (row.get('Genre') or '').strip()


def load_urls_by_genre(csv_path=DEFAULT_BULK_CSV, node_type=DEFAULT_NODE_TYPE):
    """Load bulk URLs from the bestsellers CSV, grouped by Genre.

    Returns:
        dict: { genre_name: [(url, sub_genre_label, genre_name), ...] }
        The genre is repeated inside each tuple so downstream code has a
        single shape regardless of bulk / single-URL mode.
    """
    urls_by_genre = {}
    if not os.path.exists(csv_path):
        logger.warning(f"Bulk CSV not found at '{csv_path}'; bulk mode will be empty.")
        return urls_by_genre

    with open(csv_path, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if (row.get('Node Type') or '').strip() != node_type:
                continue
            genre = (row.get('Genre') or '').strip()
            url = (row.get('Link') or '').strip()
            if not genre or not url:
                continue
            urls_by_genre.setdefault(genre, []).append(
                (url, _build_sub_genre_label(row), genre)
            )
    return urls_by_genre


# Loaded once at import; keyed by Genre exactly as it appears in the CSV.
URLS_BY_GENRE = load_urls_by_genre()
# Flat list preserved for backwards compatibility / no-genre runs.
URLS_TO_SCRAPE = [pair for entries in URLS_BY_GENRE.values() for pair in entries]
# ──────────────────────────────────────────────────────────────────────────────


# ── Enrichment configuration ──────────────────────────────────────────────────
# URL patterns blocked at the network level (Chrome DevTools Protocol) to
# speed up page loads. Adapted from amazon_results/mystery2/enrich.py and
# ultimate_enrich.py — fonts/analytics/ads/videos add nothing to the scrape.
BLOCKED_URLS = [
    "*.woff", "*.woff2", "*.ttf", "*.eot",
    "*google-analytics*", "*googletagmanager*", "*doubleclick*",
    "*amazon-adsystem*", "*fls-na.amazon.com*", "*unagi.amazon.com*",
    "*aax-us-east*", "*aax-us-iad*",
    "*.mp4", "*.webm", "*.ogg",
]

# URL used to prompt manual Amazon sign-in before enrichment begins. Goodreads
# ratings only populate on the book detail page when the request is logged in.
AMAZON_LOGIN_URL = (
    "https://www.amazon.com/ap/signin"
    "?openid.pape.max_auth_age=0"
    "&openid.return_to=https%3A%2F%2Fwww.amazon.com%2F"
    "&openid.identity=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select"
    "&openid.assoc_handle=usflex"
    "&openid.mode=checkid_setup"
    "&openid.claimed_id=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0%2Fidentifier_select"
    "&openid.ns=http%3A%2F%2Fspecs.openid.net%2Fauth%2F2.0"
)
# ──────────────────────────────────────────────────────────────────────────────


class AmazonScraper:
    def __init__(self, headless=False, delay=0.5, enrich=True):
        """
        Initialize the Amazon scraper

        Args:
            headless: Run browser in headless mode
            delay: Delay between requests in seconds (default: 0.5 for faster scraping)
            enrich: If True, also extract Synopsis, Print Length, Publication Date,
                    and Goodreads Rating/Count per book (slower per page).
        """
        self.delay = delay
        self.enrich = enrich
        self.setup_driver(headless)

    def setup_driver(self, headless=False):
        """Setup Selenium WebDriver"""
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument('--headless')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-blink-features=AutomationControlled')
        # Disable images for faster loading (CSS kept for proper page rendering)
        prefs = {
            "profile.managed_default_content_settings.images": 2
        }
        options.add_experimental_option("prefs", prefs)
        options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')

        try:
            # Use webdriver-manager to automatically download and manage ChromeDriver
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.implicitly_wait(5)  # Reduced from 10 to 5 seconds
            # Set page load timeout
            self.driver.set_page_load_timeout(15)
            # Block fonts/analytics/ads at the network level — speeds enrichment
            # passes significantly without affecting the listing/detail extracts.
            try:
                self.driver.execute_cdp_cmd("Network.enable", {})
                self.driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": BLOCKED_URLS})
            except WebDriverException as e:
                logger.warning(f"Could not install CDP URL blocklist: {e}")
            logger.info("WebDriver initialized successfully")
        except Exception as e:
            logger.error(f"Failed to initialize WebDriver: {e}")
            raise

    def amazon_login_flow(self):
        """Open Amazon's sign-in page and pause for the user to log in manually.

        Goodreads rating widgets are gated behind login — without this step the
        Goodreads Rating / Goodreads Rating Count columns stay empty.
        """
        try:
            self.driver.get(AMAZON_LOGIN_URL)
        except WebDriverException as e:
            logger.warning(f"Could not open Amazon sign-in page: {e}")
            return

        print("\n" + "=" * 60)
        print("  ACTION REQUIRED: Log in to Amazon in the browser window.")
        print("  Once you are fully logged in, return here and press ENTER.")
        print("=" * 60 + "\n")
        try:
            input()
        except EOFError:
            # Non-interactive run — skip the wait
            logger.warning("No TTY available for login prompt; continuing without login.")
            return

        try:
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "nav-link-accountList"))
            )
            logger.info("Amazon login confirmed.")
        except TimeoutException:
            logger.warning("Could not verify Amazon login — Goodreads fields may be empty.")

    def extract_enrichment_fields(self, wait_s=10):
        """Extract enrichment fields from the currently loaded book detail page.

        Returns:
            dict with synopsis, print_length, publication_date, language, asin,
            isbn, book_no_in_series, price, goodreads_rating, goodreads_rating_count.

        Adapted from amazon_results/mystery2/enrich.py (synopsis) and
        ultimate_enrich.py (print length, pub date, Goodreads). The
        Language/ASIN/ISBN/Book-No./Price extractors follow the DOM shape
        documented in the project notes.
        """
        empty = {
            'synopsis': '',
            'print_length': '',
            'publication_date': '',
            'language': '',
            'asin': '',
            'isbn': '',
            'book_no_in_series': '',
            'price': '',
            'goodreads_rating': '',
            'goodreads_rating_count': '',
        }

        # ── Synopsis ──────────────────────────────────────────────────────────
        synopsis = ''
        try:
            wait = WebDriverWait(self.driver, wait_s)
            for sel in [
                "#bookDescription_feature_div",
                "#productDescription_feature_div",
                "#productDescription",
            ]:
                try:
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))
                    break
                except TimeoutException:
                    continue

            # Scroll to description; triggers Amazon's lazy-load XHR
            self.driver.execute_script("""
                var el = document.getElementById('bookDescription_feature_div');
                if (el) {
                    window.scrollTo(0, el.getBoundingClientRect().top + window.pageYOffset - 100);
                } else {
                    window.scrollTo(0, Math.floor(document.body.scrollHeight * 0.4));
                }
            """)
            time.sleep(1.5)

            # Click the "Read more" expander if present
            for sel in [
                '[data-a-expander-name="book_description_expander"] .a-expander-prompt',
                "#bookDescription_feature_div .a-expander-prompt",
                "#bookDescription_feature_div a.a-expander-header",
            ]:
                try:
                    el = self.driver.find_element(By.CSS_SELECTOR, sel)
                    self.driver.execute_script("arguments[0].click();", el)
                    time.sleep(0.3)
                    break
                except (NoSuchElementException, WebDriverException):
                    continue

            text = self.driver.execute_script("""
                var selectors = [
                    '[data-a-expander-name="book_description_expander"] .a-expander-content',
                    '#bookDescription_feature_div .a-expander-content',
                    '#bookDescription_feature_div',
                    '#productDescription_feature_div .a-expander-content',
                    '#productDescription'
                ];
                for (var i = 0; i < selectors.length; i++) {
                    var el = document.querySelector(selectors[i]);
                    if (el) {
                        var t = el.textContent.trim();
                        if (t.length > 30) return t;
                    }
                }
                return '';
            """) or ''
            synopsis = re.sub(r"\s+", " ", text).strip()
        except WebDriverException as e:
            logger.warning(f"Synopsis extraction failed: {e}")

        # ── Print Length, Publication Date, Language, ASIN, ISBN,
        # ── Book No. in Series, Price ─────────────────────────────────────────
        print_length = ''
        pub_date = ''
        language = ''
        asin = ''
        isbn = ''
        book_no = ''
        price = ''
        try:
            try:
                WebDriverWait(self.driver, wait_s).until(EC.presence_of_element_located((
                    By.CSS_SELECTOR,
                    "#rpi-attribute-book_details-ebook_pages, #detailBullets_feature_div"
                )))
            except TimeoutException:
                pass

            self.driver.execute_script("""
                var el = document.querySelector(
                    '#rpi-attribute-book_details-ebook_pages, #detailBullets_feature_div'
                );
                if (el) window.scrollTo(0, el.getBoundingClientRect().top + window.pageYOffset - 100);
            """)
            time.sleep(0.5)

            result = self.driver.execute_script("""
                function bulletValue(re) {
                    // Find a .detail-bullet-list / #detailBullets_feature_div bullet
                    // whose .a-text-bold matches `re`, return its sibling span text.
                    var lis = document.querySelectorAll(
                        '#detailBullets_feature_div li, .detail-bullet-list > li'
                    );
                    for (var i = 0; i < lis.length; i++) {
                        var b = lis[i].querySelector('.a-text-bold');
                        if (b && re.test(b.textContent.trim())) {
                            var sib = b.nextElementSibling;
                            if (sib) return sib.textContent.trim();
                        }
                    }
                    return '';
                }

                // ── Print Length ──────────────────────────────────────────
                var pl = '';
                var plA = document.querySelector(
                    '#rpi-attribute-book_details-ebook_pages a[aria-label]'
                );
                if (plA) {
                    var lbl = plA.getAttribute('aria-label') || '';
                    pl = lbl.split(':').pop().trim();
                }
                if (!pl) {
                    var plSpan = document.querySelector(
                        '#rpi-attribute-book_details-ebook_pages .rpi-attribute-value span'
                    );
                    if (plSpan) pl = plSpan.textContent.trim();
                }
                if (!pl) pl = bulletValue(/print\\s+length/i);

                // ── Publication Date ──────────────────────────────────────
                var pd = '';
                var pdSpan = document.querySelector(
                    '#rpi-attribute-book_details-publication_date .rpi-attribute-value span'
                );
                if (pdSpan) pd = pdSpan.textContent.trim();
                if (!pd) pd = bulletValue(/publication\\s+date/i);

                // ── Language ──────────────────────────────────────────────
                var lang = '';
                var langSpan = document.querySelector(
                    '#rpi-attribute-language .rpi-attribute-value span'
                );
                if (langSpan) lang = langSpan.textContent.trim();
                if (!lang) lang = bulletValue(/^language\\b/i);

                // ── ASIN ──────────────────────────────────────────────────
                var asin = bulletValue(/^asin\\b/i);

                // ── ISBN (prefer ISBN-13 over ISBN-10) ────────────────────
                var isbn = bulletValue(/isbn-?13/i);
                if (!isbn) isbn = bulletValue(/isbn-?10/i);
                if (!isbn) isbn = bulletValue(/^isbn\\b/i);

                // ── Book No. in Series — bold text is e.g. "Book 6 of 8" ──
                var bookNo = '';
                var lis2 = document.querySelectorAll(
                    '#detailBullets_feature_div li, .detail-bullet-list > li'
                );
                for (var k = 0; k < lis2.length; k++) {
                    var b3 = lis2[k].querySelector('.a-text-bold');
                    if (b3) {
                        var m = b3.textContent.match(/Book\\s+(\\d+)\\s+of\\s+(\\d+)/i);
                        if (m) { bookNo = m[1] + ' of ' + m[2]; break; }
                    }
                }

                // ── Price (apex "Price to Pay") ───────────────────────────
                var priceStr = '';
                var priceEl = document.querySelector(
                    '.apex-pricetopay-value [aria-hidden="true"], ' +
                    '.priceToPay [aria-hidden="true"], ' +
                    '#corePrice_feature_div [aria-hidden="true"], ' +
                    '.a-price [aria-hidden="true"]'
                );
                if (priceEl) {
                    priceStr = priceEl.textContent.replace(/\\s+/g, ' ').trim();
                }
                if (!priceStr) {
                    var off = document.querySelector(
                        '.apex-pricetopay-value .a-offscreen, .a-price .a-offscreen'
                    );
                    if (off) priceStr = off.textContent.trim();
                }

                return [pl, pd, lang, asin, isbn, bookNo, priceStr];
            """)
            if result and len(result) == 7:
                print_length = result[0] or ''
                pub_date = result[1] or ''
                language = result[2] or ''
                asin = result[3] or ''
                isbn = result[4] or ''
                book_no = result[5] or ''
                price = result[6] or ''
        except WebDriverException as e:
            logger.warning(f"Detail-panel extraction failed: {e}")

        # ── Goodreads Rating & Count ─────────────────────────────────────────
        gr_rating = ''
        gr_count = ''
        try:
            self.driver.execute_script("""
                var el = document.getElementById('goodreadsRatingsWidget_feature_div');
                if (el) {
                    window.scrollTo(0, el.getBoundingClientRect().top + window.pageYOffset - 100);
                } else {
                    window.scrollTo(0, Math.floor(document.body.scrollHeight * 0.65));
                }
            """)
            time.sleep(1.5)
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".gr-review-rating-text"))
                )
            except TimeoutException:
                pass

            result = self.driver.execute_script("""
                var ratingEl = document.querySelector('.gr-review-rating-text .a-size-base');
                var countEl  = document.querySelector('.gr-review-count-text .a-size-base');
                var rating = ratingEl ? ratingEl.textContent.trim() : '';
                var count  = '';
                if (countEl) {
                    count = countEl.textContent.trim()
                        .replace(/ratings?/i, '')
                        .replace(/,/g, '')
                        .trim();
                }
                return [rating, count];
            """)
            if result and len(result) == 2:
                gr_rating = result[0] or ''
                gr_count = result[1] or ''
        except WebDriverException as e:
            logger.warning(f"Goodreads extraction failed: {e}")

        return {
            'synopsis': synopsis,
            'print_length': print_length,
            'publication_date': pub_date,
            'language': language,
            'asin': asin,
            'isbn': isbn,
            'book_no_in_series': book_no,
            'price': price,
            'goodreads_rating': gr_rating,
            'goodreads_rating_count': gr_count,
        }
    
    def close(self):
        """Close the browser"""
        if hasattr(self, 'driver'):
            self.driver.quit()
    
    def safe_get(self, url, max_retries=3):
        """Safely navigate to a URL with retries"""
        for attempt in range(max_retries):
            try:
                self.driver.get(url)
                # Reduced sleep - let page load naturally
                time.sleep(self.delay * 0.5)  # Half the delay for navigation
                return True
            except Exception as e:
                logger.warning(f"Attempt {attempt + 1} failed for {url}: {e}")
                if attempt < max_retries - 1:
                    time.sleep(self.delay)
                else:
                    logger.error(f"Failed to load {url} after {max_retries} attempts")
                    return False
        return False
    
    def extract_book_links_from_bestsellers(self, url):
        """
        Extract book links from a bestsellers page
        
        Args:
            url: URL of the bestsellers page
            
        Returns:
            List of dictionaries with book information
        """
        books = []
        
        if not self.safe_get(url):
            return books
        
        try:
            # Wait for page to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            
            # Optimized scrolling - fewer iterations with shorter delays
            for i in range(3):  # Reduced from 5 to 3
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.3)  # Reduced from 1 to 0.3
                # Scroll back up a bit to trigger more loading
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight - 1000);")
                time.sleep(0.2)  # Reduced from 0.5 to 0.2
            
            # Final scroll to bottom
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(0.5)  # Reduced from 2 to 0.5
            
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            
            # Extract books from page 1 - try multiple selectors
            book_items_page1 = soup.select('div[data-index], li[data-index], div.zg-item, div.p13n-grid-item, div[data-component-type="s-search-result"], div[data-asin]:not([data-asin=""]), ol.zg-grid li, ul.zg-grid li')
            initial_count = len(book_items_page1)
            
            # If no items found, try finding by ASIN attribute
            if initial_count == 0:
                book_items_page1 = soup.find_all('div', {'data-asin': re.compile(r'[A-Z0-9]{10}')})
                initial_count = len(book_items_page1)
            
            # Try to get additional pages - store items from all pages
            all_page_items = []  # List of (page_num, items, soup) tuples
            
            total_found = initial_count
            page_num = 2
            max_pages = 6  # Check up to 6 pages to ensure we get all books
            consecutive_empty_pages = 0  # Track consecutive pages with no items
            
            while page_num <= max_pages:
                # Try to find pagination link for next page
                # First, check if there's a pagination link on the current page
                if page_num > 1:
                    try:
                        # Look for pagination links
                        pagination_links = self.driver.find_elements(By.CSS_SELECTOR, 'a[href*="pg="], li.a-selected + li a, .a-pagination a')
                        page3_link = None
                        for link in pagination_links:
                            href = link.get_attribute('href')
                            if href and f'pg={page_num}' in href:
                                page3_link = href
                                logger.info(f"Found pagination link for page {page_num}: {page3_link[:80]}...")
                                break
                        if page3_link:
                            page_url = page3_link
                        else:
                            page_url = url + ("&pg=" + str(page_num) if "?" in url else "?pg=" + str(page_num))
                    except Exception as e:
                        logger.debug(f"Could not find pagination link, using constructed URL: {e}")
                        page_url = url + ("&pg=" + str(page_num) if "?" in url else "?pg=" + str(page_num))
                else:
                    page_url = url + ("&pg=" + str(page_num) if "?" in url else "?pg=" + str(page_num))
                logger.info(f"Found {total_found} items so far, trying page {page_num} to get more books...")
                
                try:
                    if self.safe_get(page_url):
                        # Reduced wait time for faster loading
                        wait_time = 1.5 if page_num >= 3 else (1.0 if page_num == 2 else 0.5)
                        time.sleep(wait_time)
                        # Optimized scrolling - fewer iterations with shorter delays
                        scroll_iterations = 4 if page_num >= 3 else (3 if page_num == 2 else 2)
                        scroll_delay = 0.4 if page_num >= 3 else (0.3 if page_num == 2 else 0.2)
                        for i in range(scroll_iterations):
                            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                            time.sleep(scroll_delay)
                            # Scroll back up a bit to trigger more loading
                            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight - 1000);")
                            time.sleep(0.2)  # Reduced from 0.8/0.5 to 0.2
                            # Additional scroll down to trigger more lazy loading
                            if page_num >= 2:
                                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                time.sleep(0.2)  # Reduced from 0.5 to 0.2
                        # Final scroll to bottom with reduced wait
                        self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(0.8 if page_num >= 3 else (0.5 if page_num == 2 else 0.3))
                        
                        # For page 2, wait explicitly for items to load and check count
                        if page_num == 2:
                            try:
                                # Wait for at least some items to be present
                                WebDriverWait(self.driver, 10).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-index], li[data-index], div.zg-item, div[data-asin]'))
                                )
                                # Check for "Load more" or "See more" buttons and click them
                                try:
                                    load_more_buttons = self.driver.find_elements(By.CSS_SELECTOR, 'button[aria-label*="more"], a[aria-label*="more"], button:contains("more"), .a-button:contains("more")')
                                    for button in load_more_buttons:
                                        button_text = button.text.lower()
                                        if 'more' in button_text or 'load' in button_text:
                                            logger.info(f"Found 'Load more' button, clicking it...")
                                            self.driver.execute_script("arguments[0].click();", button)
                                            time.sleep(1)  # Reduced from 3 to 1
                                except Exception as e:
                                    logger.debug(f"Could not find/click load more button: {e}")
                                
                                # Check item count using JavaScript - optimized with fewer attempts
                                item_count_js = """
                                var items = document.querySelectorAll('div[data-index], li[data-index], div.zg-item, div[data-asin]:not([data-asin=""])');
                                return items.length;
                                """
                                for attempt in range(3):  # Reduced from 5 to 3
                                    count = self.driver.execute_script(item_count_js)
                                    logger.info(f"Page 2 JavaScript item count check (attempt {attempt+1}): {count}")
                                    if count >= 50:
                                        break
                                    time.sleep(0.8)  # Reduced from 2 to 0.8
                                    # Scroll a bit more and try clicking any expand buttons
                                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                    time.sleep(0.3)  # Reduced from 1 to 0.3
                                    # Try to find and click any expand/show more elements
                                    try:
                                        expand_elems = self.driver.find_elements(By.CSS_SELECTOR, '[aria-expanded="false"], .a-expander-header')
                                        for elem in expand_elems[:2]:  # Reduced from 3 to 2
                                            self.driver.execute_script("arguments[0].click();", elem)
                                            time.sleep(0.3)  # Reduced from 1 to 0.3
                                    except:
                                        pass
                                # Reduced wait for lazy loading
                                time.sleep(1)  # Reduced from 4 to 1
                            except TimeoutException:
                                logger.warning("Timeout waiting for page 2 items to load")
                        
                        page_soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                        page_items = page_soup.select('div[data-index], li[data-index], div.zg-item, div.p13n-grid-item, div[data-component-type="s-search-result"], div[data-asin]:not([data-asin=""]), ol.zg-grid li, ul.zg-grid li')
                        if len(page_items) == 0:
                            page_items = page_soup.find_all('div', {'data-asin': re.compile(r'[A-Z0-9]{10}')})
                        # For page 3, log if we found links via direct extraction
                        if page_num == 3 and len(page_items) == 0:
                            page_links = page_soup.select('a[href*="/dp/"], a[href*="/gp/product/"]')
                            if page_links:
                                logger.info(f"Page 3: Found {len(page_links)} links (will be processed in fallback)")
                        
                        # Store items from this page
                        all_page_items.append((page_num, page_items, page_soup))
                        
                        logger.info(f"Found {len(page_items)} items on page {page_num}")
                        
                        # For page 2, if we got fewer than 50 items, retry with optimized scrolling
                        if page_num == 2 and len(page_items) < 50:
                            logger.warning(f"Page 2 only found {len(page_items)} items (expected 50), retrying with optimized scrolling...")
                            # Optimized scrolling - fewer steps with shorter delays
                            page_height = self.driver.execute_script("return document.body.scrollHeight")
                            scroll_step = page_height // 10  # Reduced from 20 to 10 steps
                            for step in range(10):  # Reduced from 20 to 10
                                scroll_pos = scroll_step * (step + 1)
                                self.driver.execute_script(f"window.scrollTo(0, {scroll_pos});")
                                time.sleep(0.3)  # Reduced from 0.8 to 0.3
                            # Optimized scrolling back and forth
                            for i in range(3):  # Reduced from 5 to 3
                                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                time.sleep(0.5)  # Reduced from 2 to 0.5
                                self.driver.execute_script("window.scrollTo(0, 0);")
                                time.sleep(0.3)  # Reduced from 1 to 0.3
                                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                                time.sleep(0.5)  # Reduced from 2 to 0.5
                            # Reduced wait for lazy loading
                            time.sleep(1.5)  # Reduced from 6 to 1.5
                            # Re-extract
                            page_soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                            page_items_retry = page_soup.select('div[data-index], li[data-index], div.zg-item, div.p13n-grid-item, div[data-component-type="s-search-result"], div[data-asin]:not([data-asin=""]), ol.zg-grid li, ul.zg-grid li')
                            if len(page_items_retry) == 0:
                                page_items_retry = page_soup.find_all('div', {'data-asin': re.compile(r'[A-Z0-9]{10}')})
                            logger.info(f"Retry found {len(page_items_retry)} items (was {len(page_items)})")
                            if len(page_items_retry) >= len(page_items):
                                page_items = page_items_retry
                                # Update the stored items for this page
                                if all_page_items and all_page_items[-1][0] == page_num:
                                    all_page_items[-1] = (page_num, page_items, page_soup)
                                if len(page_items_retry) > len(page_items):
                                    logger.info(f"✅ Retry improved count from {len(page_items)} to {len(page_items_retry)}")
                        
                        total_found += len(page_items)
                        
                        # Track consecutive empty pages
                        if len(page_items) == 0:
                            consecutive_empty_pages += 1
                            logger.info(f"Page {page_num} has 0 items (consecutive empty: {consecutive_empty_pages})")
                            # If we get 2 consecutive empty pages, stop
                            if consecutive_empty_pages >= 1:
                                logger.info(f"Stopping: Found {consecutive_empty_pages} consecutive empty pages. Total books found: {total_found}")
                                break
                        else:
                            consecutive_empty_pages = 0  # Reset counter if we found items
                            logger.info(f"Page {page_num} has {len(page_items)} items, total: {total_found}")
                        
                        # Continue to next page if we haven't hit max pages or consecutive empty pages
                    else:
                        logger.warning(f"Failed to load page {page_num}: {page_url}")
                        break
                except Exception as e:
                    logger.error(f"Error loading page {page_num}: {e}")
                    break
                
                page_num += 1
            
            # Extract books from structured items, deduplicating by ASIN
            seen_asins = set()
            seen_links = set()
            
            # Process page 1 items
            for item in book_items_page1:
                try:
                    link_elem = item.find('a', href=re.compile(r'/dp/|/gp/product/'))
                    if not link_elem:
                        continue
                    
                    href = link_elem.get('href', '')
                    full_url = urljoin('https://www.amazon.com', href)
                    
                    # Extract ASIN
                    asin_match = re.search(r'(?:/dp/|/gp/product/)([A-Z0-9]{10})', full_url)
                    asin = asin_match.group(1) if asin_match else ''
                    
                    # Deduplicate by ASIN and link
                    if asin and asin in seen_asins:
                        logger.debug(f"Skipping duplicate ASIN on page 1: {asin}")
                        continue
                    if full_url in seen_links:
                        logger.debug(f"Skipping duplicate link on page 1: {full_url[:50]}...")
                        continue
                    
                    seen_asins.add(asin)
                    seen_links.add(full_url)
                    
                    # Extract title
                    title_elem = item.find(['span', 'div'], class_=re.compile('text|title', re.I))
                    if not title_elem:
                        title_elem = link_elem
                    title = title_elem.get_text(strip=True) if title_elem else ''
                    
                    books.append({
                        'rank': len(books) + 1,  # Assign rank based on order found
                        'book_link': full_url,
                        'asin': asin,
                        'title': title
                    })
                except Exception as e:
                    logger.warning(f"Error extracting book from page 1: {e}")
                    continue
            
            # Process items from all additional pages (page 2, 3, 4, 5, 6...)
            for page_num, page_items, page_soup in all_page_items:
                for item in page_items:
                    try:
                        link_elem = item.find('a', href=re.compile(r'/dp/|/gp/product/'))
                        if not link_elem:
                            continue
                        
                        href = link_elem.get('href', '')
                        full_url = urljoin('https://www.amazon.com', href)
                        
                        # Extract ASIN
                        asin_match = re.search(r'(?:/dp/|/gp/product/)([A-Z0-9]{10})', full_url)
                        asin = asin_match.group(1) if asin_match else ''
                        
                        # Skip if already seen
                        if asin and asin in seen_asins:
                            logger.debug(f"Skipping duplicate ASIN on page {page_num}: {asin}")
                            continue
                        if full_url in seen_links:
                            logger.debug(f"Skipping duplicate link on page {page_num}: {full_url[:50]}...")
                            continue
                        
                        seen_asins.add(asin)
                        seen_links.add(full_url)
                        
                        # Extract title
                        title_elem = item.find(['span', 'div'], class_=re.compile('text|title', re.I))
                        if not title_elem:
                            title_elem = link_elem
                        title = title_elem.get_text(strip=True) if title_elem else ''
                        
                        books.append({
                            'rank': len(books) + 1,  # Continue ranking from previous pages
                            'book_link': full_url,
                            'asin': asin,
                            'title': title
                        })
                    except Exception as e:
                        logger.warning(f"Error extracting book from page {page_num}: {e}")
                        continue
            
            # Fallback: If no structured items found, try finding links directly from both pages
            if not books:
                logger.info("No books found via structured items, trying direct link extraction...")
                # Try page 1 first
                book_links = soup.select('a[href*="/dp/"], a[href*="/gp/product/"]')
                logger.info(f"Found {len(book_links)} links on page 1")
                
                for link in book_links:
                    if len(books) >= 100:
                        break
                    href = link.get('href', '')
                    # Clean href (remove query params that might cause duplicates)
                    clean_href = re.sub(r'\?.*$', '', href)
                    if '/dp/' in clean_href or '/gp/product/' in clean_href:
                        full_url = urljoin('https://www.amazon.com', clean_href)
                        # Extract ASIN from URL
                        asin_match = re.search(r'(?:/dp/|/gp/product/)([A-Z0-9]{10})', full_url)
                        asin = asin_match.group(1) if asin_match else ''
                        
                        # Deduplicate
                        if asin and asin in seen_asins:
                            continue
                        if full_url in seen_links:
                            continue
                        
                        seen_asins.add(asin)
                        seen_links.add(full_url)
                        
                        # Try to get title
                        title_elem = link.find(['span', 'div'], class_=re.compile('text|title', re.I))
                        if not title_elem:
                            title_elem = link.find_parent(['div', 'li'])
                            if title_elem:
                                title_elem = title_elem.find(['span', 'div'], class_=re.compile('text|title', re.I))
                        title = title_elem.get_text(strip=True) if title_elem else ''
                        
                        books.append({
                            'rank': len(books) + 1,
                            'book_link': full_url,
                            'asin': asin,
                            'title': title
                        })
                
                # Try all additional pages if we have them (fallback extraction)
                for page_num, page_items, page_soup in all_page_items:
                    logger.info(f"Trying page {page_num} links (fallback), currently have {len(books)} books")
                    book_links_page = page_soup.select('a[href*="/dp/"], a[href*="/gp/product/"]')
                    logger.info(f"Found {len(book_links_page)} links on page {page_num}")
                    
                    for link in book_links_page:
                        href = link.get('href', '')
                        clean_href = re.sub(r'\?.*$', '', href)
                        if '/dp/' in clean_href or '/gp/product/' in clean_href:
                            full_url = urljoin('https://www.amazon.com', clean_href)
                            asin_match = re.search(r'(?:/dp/|/gp/product/)([A-Z0-9]{10})', full_url)
                            asin = asin_match.group(1) if asin_match else ''
                            
                            if asin and asin in seen_asins:
                                continue
                            if full_url in seen_links:
                                continue
                            
                            seen_asins.add(asin)
                            seen_links.add(full_url)
                            
                            title_elem = link.find(['span', 'div'], class_=re.compile('text|title', re.I))
                            if not title_elem:
                                title_elem = link.find_parent(['div', 'li'])
                                if title_elem:
                                    title_elem = title_elem.find(['span', 'div'], class_=re.compile('text|title', re.I))
                            title = title_elem.get_text(strip=True) if title_elem else ''
                            
                            books.append({
                                'rank': len(books) + 1,
                                'book_link': full_url,
                                'asin': asin,
                                'title': title
                            })
            
            logger.info(f"Found {len(books)} books on {url}")
            
        except Exception as e:
            logger.error(f"Error extracting books from {url}: {e}")
        
        return books
    
    def extract_kindle_details(self, soup):
        """
        Extract Kindle product details from detail page
        
        Args:
            soup: BeautifulSoup object of the detail page
            
        Returns:
            Dictionary with Kindle product details
        """
        details = {
            'book_name': '',
            'series_name': '',
            'amazon_series_link': '',
            'series_length': '',  # e.g., "16 book series"
            'book_publisher': '',
            'amazon_star_rating': '',
            'amazon_rating_count': '',
            'amazon_book_metrics': '',
            'author': ''  # Add author to Kindle details
        }
        
        try:
            # Extract book name
            title_elem = soup.find('span', {'id': 'productTitle'})
            if not title_elem:
                title_elem = soup.find('h1', class_=re.compile('title', re.I))
            if title_elem:
                details['book_name'] = title_elem.get_text(strip=True)
            
            # Extract series information - try multiple methods
            # Method 1: Look for "Book X of Y" pattern in detail bullets (most reliable)
            detail_bullets_temp = soup.find('div', {'id': 'detailBullets_feature_div'})
            if not detail_bullets_temp:
                detail_bullets_temp = soup.find('div', class_=re.compile('detail-bullets', re.I))
            
            if detail_bullets_temp:
                bullets = detail_bullets_temp.find_all('li')
                for bullet in bullets:
                    text = bullet.get_text(strip=True)
                    # Look for "Book X of Y" pattern
                    book_match = re.search(r'Book\s+\d+\s+of\s+\d+', text, re.I)
                    if book_match:
                        # Find the series link in this bullet
                        series_link_elem = bullet.find('a', href=re.compile(r'/dp/'))
                        if series_link_elem:
                            series_text = series_link_elem.get_text(strip=True)
                            # Clean up the text (remove "Book X of Y" part if present)
                            series_text = re.sub(r'\s*Book\s+\d+\s+of\s+\d+.*$', '', series_text, flags=re.I).strip()
                            # Remove special characters and clean up
                            series_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', series_text).strip()
                            if series_text and len(series_text) > 3:
                                details['series_name'] = series_text
                                series_href = series_link_elem.get('href', '')
                                if series_href:
                                    details['amazon_series_link'] = urljoin('https://www.amazon.com', series_href)
                                    break
                        
                        # Alternative: Extract from text pattern "Book X of Y : Series Name"
                        if not details['series_name']:
                            # Look for pattern like "Book 1 of 8 : Series Name"
                            series_match = re.search(r'Book\s+(\d+)\s+of\s+(\d+)\s*[:\u200E\u200F]*\s*([A-Z][^:\n]+?)(?:\s*See|\s*$)', text, re.I | re.DOTALL)
                            if series_match:
                                book_num = series_match.group(1)
                                total_books = series_match.group(2)
                                series_text = series_match.group(3).strip()
                                series_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', series_text).strip()
                                if series_text and len(series_text) > 3:
                                    details['series_name'] = series_text
                                    details['series_length'] = f"{total_books} book series"
                                    # Try to find the link
                                    series_link_elem = bullet.find('a', href=re.compile(r'/dp/'))
                                    if series_link_elem:
                                        series_href = series_link_elem.get('href', '')
                                        if series_href:
                                            details['amazon_series_link'] = urljoin('https://www.amazon.com', series_href)
                                    break
                        
                        # Extract series length from "Book X of Y" pattern even if we already have series name
                        if not details['series_length']:
                            book_match = re.search(r'Book\s+\d+\s+of\s+(\d+)', text, re.I)
                            if book_match:
                                total_books = book_match.group(1)
                                details['series_length'] = f"{total_books} book series"
            
            # Method 2: Look for series in "Shop this series" section (fallback)
            if not details['series_name']:
                series_elem = soup.find('a', {'id': 'see-full-series'})
                if series_elem:
                    # Find the series name from the h2 heading in the same section
                    series_section = series_elem.find_parent(['div', 'section'])
                    if series_section:
                        h2_elem = series_section.find('h2')
                        if h2_elem:
                            h2_text = h2_elem.get_text(strip=True)
                            # Extract series name and length (e.g., "A Cafe Crimes Cozy Mystery Series (16 book series)")
                            series_name_match = re.search(r'^(.+?)(?:\s*\((\d+)\s*-?\s*book\s+series\))?(?:\s*Shop this series|$)', h2_text, re.I)
                            if series_name_match:
                                details['series_name'] = series_name_match.group(1).strip()
                                if series_name_match.group(2) and not details['series_length']:
                                    details['series_length'] = f"{series_name_match.group(2)} book series"
                        series_href = series_elem.get('href', '')
                        if series_href:
                            details['amazon_series_link'] = urljoin('https://www.amazon.com', series_href)
            
            # Method 3: Look for series link with /dp/.*series pattern (fallback)
            if not details['series_name']:
                series_elem = soup.find('a', href=re.compile(r'/dp/.*series|saga_dp', re.I))
                if series_elem:
                    series_text = series_elem.get_text(strip=True)
                    # Extract series length if present in text (e.g., "Series Name (16 book series)")
                    length_match = re.search(r'\((\d+)\s*-?\s*book\s+series\)', series_text, re.I)
                    if length_match and not details['series_length']:
                        details['series_length'] = f"{length_match.group(1)} book series"
                    # Clean up text
                    series_text = re.sub(r'\s*\(.*book\s+series\).*$', '', series_text, flags=re.I).strip()
                    series_text = re.sub(r'\s*See full series.*$', '', series_text, flags=re.I).strip()
                    if series_text and len(series_text) > 3:
                        details['series_name'] = series_text
                        series_href = series_elem.get('href', '')
                        if series_href:
                            details['amazon_series_link'] = urljoin('https://www.amazon.com', series_href)
            
            # Extract series length from series name if it contains "(X book series)" pattern
            if details['series_name'] and not details['series_length']:
                length_match = re.search(r'\((\d+)\s*-?\s*book\s+series\)', details['series_name'], re.I)
                if length_match:
                    details['series_length'] = f"{length_match.group(1)} book series"
                    # Remove the length from series name
                    details['series_name'] = re.sub(r'\s*\(.*book\s+series\).*$', '', details['series_name'], flags=re.I).strip()
            
            # Extract product details from detail bullets
            detail_bullets = soup.find('div', {'id': 'detailBullets_feature_div'})
            if not detail_bullets:
                detail_bullets = soup.find('div', class_=re.compile('detail-bullets', re.I))
            
            if detail_bullets:
                bullets = detail_bullets.find_all('li')
                for bullet in bullets:
                    text = bullet.get_text(strip=True)
                    
                    # Publisher - look for span after Publisher label
                    if 'Publisher' in text and not details['book_publisher']:
                        # Find all spans in the bullet
                        spans = bullet.find_all('span')
                        for i, span in enumerate(spans):
                            span_text = span.get_text(strip=True)
                            # Check if this span contains "Publisher"
                            if 'Publisher' in span_text and i + 1 < len(spans):
                                # Next span should contain the publisher name
                                next_span = spans[i + 1]
                                publisher_text = next_span.get_text(strip=True)
                                # Remove special characters (RTL marks, etc.)
                                publisher_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', publisher_text).strip()
                                # Remove "Publisher" if it's still in there
                                publisher_text = re.sub(r'^Publisher\s*:?\s*', '', publisher_text, flags=re.I).strip()
                                if publisher_text and len(publisher_text) > 2:
                                    details['book_publisher'] = publisher_text
                                    break
                        # Fallback: Try regex on the full text
                        if not details['book_publisher']:
                            # Try multiple regex patterns
                            patterns = [
                                r'Publisher\s*:?\s*([A-Z][^\n]+?)(?:\s*Accessibility|\s*Publication|\s*Language|\s*ASIN|$)',
                                r'Publisher[:\s]+([A-Z][^;]+)',
                                r'Publisher[:\s]+([A-Za-z][^,]+)',
                            ]
                            for pattern in patterns:
                                publisher_match = re.search(pattern, text, re.I | re.DOTALL)
                                if publisher_match:
                                    pub_text = publisher_match.group(1).strip()
                                    pub_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', pub_text).strip()
                                    # Clean up common suffixes
                                    pub_text = re.sub(r'\s*(?:\(|Accessibility|Publication|Language).*$', '', pub_text, flags=re.I)
                                    if pub_text and len(pub_text) > 2:
                                        details['book_publisher'] = pub_text
                                        break
                    
                    # Best Sellers Rank (Amazon Book Metrics)
                    if 'Best Sellers Rank' in text and not details['amazon_book_metrics']:
                        rank_text = bullet.get_text(separator=' ', strip=True)
                        # Extract the full ranking text
                        rank_match = re.search(r'Best Sellers Rank[:\s]*(.+?)(?:\n|$)', rank_text, re.I | re.DOTALL)
                        if rank_match:
                            details['amazon_book_metrics'] = rank_match.group(1).strip()
                    
                    # Customer Reviews (Star Rating and Count)
                    if 'Customer Reviews' in text:
                        # Try to find star rating
                        stars_elem = bullet.find('span', class_=re.compile('star|rating', re.I))
                        if stars_elem:
                            star_text = stars_elem.get_text(strip=True)
                            star_match = re.search(r'(\d+\.?\d*)\s*(?:out of|stars?)', star_text, re.I)
                            if star_match:
                                details['amazon_star_rating'] = star_match.group(1)
                        
                        # Try to find rating count
                        count_elem = bullet.find('span', {'id': 'acrCustomerReviewText'})
                        if not count_elem:
                            count_elem = bullet.find('a', href=re.compile('customerReviews', re.I))
                        if count_elem:
                            count_text = count_elem.get_text(strip=True)
                            count_match = re.search(r'\(?(\d+[,\d]*)\s*reviews?\)?', count_text, re.I)
                            if count_match:
                                details['amazon_rating_count'] = count_match.group(1).replace(',', '')
            
            # Alternative method: Try to find star rating in separate element
            if not details['amazon_star_rating']:
                rating_elem = soup.find('span', {'id': 'acrPopover'})
                if rating_elem:
                    rating_text = rating_elem.get_text(strip=True)
                    star_match = re.search(r'(\d+\.?\d*)', rating_text)
                    if star_match:
                        details['amazon_star_rating'] = star_match.group(1)
            
            if not details['amazon_rating_count']:
                count_elem = soup.find('span', {'id': 'acrCustomerReviewText'})
                if count_elem:
                    count_text = count_elem.get_text(strip=True)
                    count_match = re.search(r'(\d+[,\d]*)', count_text)
                    if count_match:
                        details['amazon_rating_count'] = count_match.group(1).replace(',', '')
            
            # Extract author - try multiple methods
            # Method 1: Look for author link inside bylineInfo div (most reliable)
            byline_info = soup.find('div', {'id': 'bylineInfo'})
            if byline_info:
                # Find author link inside bylineInfo
                author_link = byline_info.find('a', class_='a-link-normal', href=re.compile(r'/e/|/s/', re.I))
                if author_link:
                    author_text = author_link.get_text(strip=True)
                    if author_text and len(author_text) > 2:
                        details['author'] = author_text
                else:
                    # Fallback: look for author span
                    author_span = byline_info.find('span', class_='author')
                    if author_span:
                        author_link = author_span.find('a', href=re.compile(r'/e/|/s/', re.I))
                        if author_link:
                            author_text = author_link.get_text(strip=True)
                            if author_text and len(author_text) > 2:
                                details['author'] = author_text
            
            # Method 2: Look for author link with /e/ or /s/ pattern (fallback)
            if not details.get('author'):
                author_elem = soup.find('a', href=re.compile(r'/e/|/s/', re.I))
                if author_elem:
                    author_text = author_elem.get_text(strip=True)
                    # Filter out non-author links (follow, etc.)
                    if author_text and len(author_text) > 2 and 'follow' not in author_text.lower():
                        # Make sure it's not a series link
                        href = author_elem.get('href', '')
                        if 'series' not in href.lower() and 'saga' not in href.lower():
                            details['author'] = author_text
            
            # Method 3: Look for "by Author Name" pattern
            if not details.get('author'):
                by_pattern = soup.find(string=re.compile(r'\bby\s+[A-Z]', re.I))
                if by_pattern:
                    parent = by_pattern.parent
                    author_link = parent.find('a', href=re.compile(r'/e/|/s/', re.I))
                    if author_link:
                        author_text = author_link.get_text(strip=True)
                        if author_text and len(author_text) > 2:
                            details['author'] = author_text
            
            # Method 4: Look in contributorNameID data attribute
            if not details.get('author'):
                author_data = soup.find(attrs={'data-contributor-name-id': True})
                if author_data:
                    author_text = author_data.get_text(strip=True)
                    if author_text and len(author_text) > 2:
                        details['author'] = author_text
            
            # Method 5: Look in detail bullets for Author
            if not details.get('author') and detail_bullets:
                bullets = detail_bullets.find_all('li')
                for bullet in bullets:
                    text = bullet.get_text(strip=True)
                    if 'Author' in text or 'Author:' in text:
                        # Try to find author link
                        author_link = bullet.find('a', href=re.compile(r'/e/|/s/', re.I))
                        if author_link:
                            author_text = author_link.get_text(strip=True)
                            if author_text and len(author_text) > 2:
                                details['author'] = author_text
                                break
                        # Fallback: extract from spans
                        spans = bullet.find_all('span')
                        for i, span in enumerate(spans):
                            span_text = span.get_text(strip=True)
                            if 'Author' in span_text and i + 1 < len(spans):
                                next_span = spans[i + 1]
                                author_text = next_span.get_text(strip=True)
                                author_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', author_text).strip()
                                # Remove "(Author)" suffix if present
                                author_text = re.sub(r'\s*\(Author\)\s*$', '', author_text, flags=re.I).strip()
                                if author_text and len(author_text) > 2:
                                    details['author'] = author_text
                                    break
                        if details.get('author'):
                            break
        
        except Exception as e:
            logger.warning(f"Error extracting Kindle details: {e}")
        
        # If we have a series link but no series length, visit series page to get it
        if details.get('amazon_series_link') and not details.get('series_length'):
            series_length = self.extract_series_length(details['amazon_series_link'])
            if series_length:
                details['series_length'] = series_length
        
        return details
    
    def extract_series_length(self, series_url):
        """
        Extract series length from series page
        
        Args:
            series_url: URL of the series page
            
        Returns:
            Series length string (e.g., "16 book series") or empty string
        """
        try:
            if not self.safe_get(series_url):
                return ''
            
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')
            
            # Method 1: Look for "X book series" in the page title or heading
            # Series pages often have text like "A Cafe Crimes Cozy Mystery Series (16 book series)"
            page_text = soup.get_text()
            
            # Pattern 1: "(X book series)" or "(X-book series)"
            series_match = re.search(r'\((\d+)\s*-?\s*book\s+series\)', page_text, re.I)
            if series_match:
                return f"{series_match.group(1)} book series"
            
            # Pattern 2: "X book series" without parentheses
            series_match = re.search(r'(\d+)\s*-?\s*book\s+series', page_text, re.I)
            if series_match:
                return f"{series_match.group(1)} book series"
            
            # Method 2: Look in series title/heading
            series_title = soup.find('h1') or soup.find('h2')
            if series_title:
                title_text = series_title.get_text()
                series_match = re.search(r'\((\d+)\s*-?\s*book\s+series\)', title_text, re.I)
                if series_match:
                    return f"{series_match.group(1)} book series"
            
            # Method 3: Count books in the series list
            # Look for series items or book count indicators
            series_items = soup.find_all(['div', 'li'], class_=re.compile(r'series|book-item', re.I))
            if series_items:
                # Try to find count text
                for item in series_items:
                    item_text = item.get_text()
                    count_match = re.search(r'(\d+)\s*(?:books?|items?)', item_text, re.I)
                    if count_match:
                        return f"{count_match.group(1)} book series"
            
            # Method 4: Look for "There are X books in this series" text
            series_text_elem = soup.find(string=re.compile(r'there are \d+ books? in this series', re.I))
            if series_text_elem:
                count_match = re.search(r'there are (\d+) books? in this series', series_text_elem, re.I)
                if count_match:
                    return f"{count_match.group(1)} book series"
        
        except Exception as e:
            logger.warning(f"Error extracting series length from {series_url}: {e}")
        
        return ''
    
    def extract_audiobook_details(self, soup, kindle_url):
        """
        Extract Audiobook product details from detail page
        
        Args:
            soup: BeautifulSoup object of the detail page
            kindle_url: URL of the Kindle edition
            
        Returns:
            Dictionary with Audiobook product details
        """
        details = {
            'audiobook_link': '',
            'audiobook_name': '',
            'author': '',
            'narrator': '',
            'audiobook_publisher': '',
            'amazon_audiobook_metrics': ''
        }
        
        try:
            # Try to find audiobook link - multiple methods
            audible_links = []
            
            # Method 1: Look for "tmm_aud_swatch_0" which indicates audiobook format (most reliable)
            for link in soup.find_all('a', href=re.compile(r'tmm_aud_swatch')):
                if link not in audible_links:
                    audible_links.append(link)
            
            # Method 2: Look for /gp/product/ links with taud in ref (audiobook format switcher)
            for link in soup.find_all('a', href=re.compile(r'/gp/product/.*taud|taud.*/gp/product/', re.I)):
                if link not in audible_links:
                    audible_links.append(link)
            
            # Method 3: Look for "Audible" or "Audiobook" links with /dp/
            audible_links.extend(soup.find_all('a', href=re.compile(r'/dp/[A-Z0-9]{10}.*audible|audible.*/dp/', re.I)))
            
            # Method 4: Look for links with text containing "Audible" or "Audiobook"
            for link in soup.find_all('a'):
                href = link.get('href', '')
                text = link.get_text(strip=True).lower()
                # Check for both /dp/ and /gp/product/ formats
                if ('audible' in text or 'audiobook' in text) and ('/dp/' in href or '/gp/product/' in href):
                    if link not in audible_links:
                        audible_links.append(link)
            
            # Method 5: Look in "formats" section for audiobook
            formats_section = soup.find('div', {'id': re.compile(r'formats|editions', re.I)})
            if formats_section:
                for link in formats_section.find_all('a', href=re.compile(r'/dp/|/gp/product/')):
                    link_text = link.get_text(strip=True).lower()
                    href = link.get('href', '')
                    # Check if it's an audiobook link
                    if (('audible' in link_text or 'audiobook' in link_text) or 
                        'taud' in href or 'tmm_aud' in href) and link not in audible_links:
                        audible_links.append(link)
            
            # Method 6: Look for format switcher buttons/links
            format_switcher = soup.find('div', {'id': re.compile(r'formatSelector|format-switcher', re.I)})
            if format_switcher:
                for link in format_switcher.find_all('a', href=re.compile(r'/dp/|/gp/product/')):
                    href = link.get('href', '')
                    if ('taud' in href or 'tmm_aud' in href) and link not in audible_links:
                        audible_links.append(link)
            
            # Method 7: Fallback - search all links for taud parameter (last resort)
            if not audible_links:
                for link in soup.find_all('a', href=True):
                    href = link.get('href', '')
                    # Look for links with taud in query params (audiobook format indicator)
                    if 'taud' in href and ('/dp/' in href or '/gp/product/' in href):
                        audible_links.append(link)
                        logger.debug(f"Found audiobook link via taud fallback: {href}")
                        break
            
            if audible_links:
                audiobook_link = audible_links[0].get('href', '')
                # Handle both /dp/ and /gp/product/ formats
                # Amazon accepts both, so we'll use the original format
                details['audiobook_link'] = urljoin('https://www.amazon.com', audiobook_link)
                logger.info(f"Found audiobook link: {details['audiobook_link']}")
                
                # Navigate to audiobook page to get details
                if self.safe_get(details['audiobook_link']):
                    audiobook_soup = BeautifulSoup(self.driver.page_source, 'html.parser')
                    
                    # Extract audiobook name
                    title_elem = audiobook_soup.find('span', {'id': 'productTitle'})
                    if title_elem:
                        details['audiobook_name'] = title_elem.get_text(strip=True)
                    
                    # Extract author from table structure
                    author_row = audiobook_soup.find('tr', {'id': 'detailsauthor'})
                    if author_row:
                        author_link = author_row.find('a', class_='a-link-normal')
                        if author_link:
                            details['author'] = author_link.get_text(strip=True)
                        else:
                            # Fallback: get text from td
                            td = author_row.find('td')
                            if td:
                                details['author'] = td.get_text(strip=True)
                    
                    # Fallback: Extract author from links if table not found
                    if not details['author']:
                        author_elem = audiobook_soup.find('a', href=re.compile(r'/e/|/s/', re.I))
                        if author_elem:
                            author_text = author_elem.get_text(strip=True)
                            if author_text and len(author_text) > 2 and 'follow' not in author_text.lower():
                                details['author'] = author_text
                    
                    # Extract narrator and publisher from table structure (audiobook pages use tables)
                    # Method 1: Look for table rows with specific IDs
                    narrator_row = audiobook_soup.find('tr', {'id': 'detailsnarrator'})
                    if narrator_row:
                        narrator_link = narrator_row.find('a', class_='a-link-normal')
                        if narrator_link:
                            details['narrator'] = narrator_link.get_text(strip=True)
                        else:
                            # Fallback: get text from td
                            td = narrator_row.find('td')
                            if td:
                                details['narrator'] = td.get_text(strip=True)
                    
                    publisher_row = audiobook_soup.find('tr', {'id': 'detailspublisher'})
                    if publisher_row:
                        publisher_link = publisher_row.find('a', class_='a-link-normal')
                        if publisher_link:
                            details['audiobook_publisher'] = publisher_link.get_text(strip=True)
                        else:
                            # Fallback: get text from td
                            td = publisher_row.find('td')
                            if td:
                                details['audiobook_publisher'] = td.get_text(strip=True)
                    
                    # Extract Best Sellers Rank from table
                    rank_row = audiobook_soup.find('tr', string=re.compile(r'Best Sellers Rank', re.I))
                    if not rank_row:
                        # Try finding by th text
                        rank_th = audiobook_soup.find('th', string=re.compile(r'Best Sellers Rank', re.I))
                        if rank_th:
                            rank_row = rank_th.find_parent('tr')
                    
                    if rank_row:
                        rank_td = rank_row.find('td')
                        if rank_td:
                            rank_items = rank_td.find_all('li')
                            rank_texts = []
                            for item in rank_items:
                                text = item.get_text(strip=True)
                                if text:
                                    rank_texts.append(text)
                            if rank_texts:
                                details['amazon_audiobook_metrics'] = ' '.join(rank_texts)
                    
                    # Method 2: Fallback to detail bullets if table structure not found
                    if not details['narrator'] or not details['audiobook_publisher']:
                        detail_bullets = audiobook_soup.find('div', {'id': 'detailBullets_feature_div'})
                        if not detail_bullets:
                            detail_bullets = audiobook_soup.find('div', class_=re.compile('detail-bullets', re.I))
                        
                        if detail_bullets:
                            bullets = detail_bullets.find_all('li')
                            for bullet in bullets:
                                text = bullet.get_text(strip=True)
                                
                                # Extract narrator
                                if ('Narrated by' in text or 'Narrator' in text) and not details['narrator']:
                                    spans = bullet.find_all('span')
                                    for i, span in enumerate(spans):
                                        span_text = span.get_text(strip=True)
                                        if ('Narrated by' in span_text or 'Narrator' in span_text) and i + 1 < len(spans):
                                            next_span = spans[i + 1]
                                            narrator_text = next_span.get_text(strip=True)
                                            narrator_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', narrator_text).strip()
                                            narrator_text = re.sub(r'^(?:Narrated by|Narrator)\s*:?\s*', '', narrator_text, flags=re.I).strip()
                                            if narrator_text and len(narrator_text) > 2:
                                                details['narrator'] = narrator_text
                                                break
                                    
                                    if not details['narrator']:
                                        narrator_match = re.search(r'(?:Narrated by|Narrator)[:\s]+([A-Z][^:\n]+?)(?:\s*Publisher|\s*Release|\s*Language|$)', text, re.I | re.DOTALL)
                                        if narrator_match:
                                            narrator_text = narrator_match.group(1).strip()
                                            narrator_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', narrator_text).strip()
                                            if narrator_text and len(narrator_text) > 2:
                                                details['narrator'] = narrator_text
                                
                                # Extract publisher
                                if 'Publisher' in text and not details['audiobook_publisher']:
                                    spans = bullet.find_all('span')
                                    for i, span in enumerate(spans):
                                        span_text = span.get_text(strip=True)
                                        if 'Publisher' in span_text and i + 1 < len(spans):
                                            next_span = spans[i + 1]
                                            pub_text = next_span.get_text(strip=True)
                                            pub_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', pub_text).strip()
                                            pub_text = re.sub(r'^Publisher\s*:?\s*', '', pub_text, flags=re.I).strip()
                                            if pub_text and len(pub_text) > 2:
                                                details['audiobook_publisher'] = pub_text
                                                break
                                    
                                    if not details['audiobook_publisher']:
                                        publisher_match = re.search(r'Publisher[:\s]+([A-Z][^:\n]+?)(?:\s*Release|\s*Language|\s*ASIN|$)', text, re.I | re.DOTALL)
                                        if publisher_match:
                                            pub_text = publisher_match.group(1).strip()
                                            pub_text = re.sub(r'[\u200E\u200F\u202A-\u202E\s]+', ' ', pub_text).strip()
                                            if pub_text and len(pub_text) > 2:
                                                details['audiobook_publisher'] = pub_text
                                
                                # Extract Best Sellers Rank
                                if 'Best Sellers Rank' in text and not details['amazon_audiobook_metrics']:
                                    rank_text = bullet.get_text(separator=' ', strip=True)
                                    rank_match = re.search(r'Best Sellers Rank[:\s]*(.+?)(?:\n|$)', rank_text, re.I | re.DOTALL)
                                    if rank_match:
                                        details['amazon_audiobook_metrics'] = rank_match.group(1).strip()
            
            # Try to extract author from Kindle page if not found (use same logic as extract_kindle_details)
            if not details['author']:
                # Method 1: Look for author link inside bylineInfo div (most reliable)
                byline_info = soup.find('div', {'id': 'bylineInfo'})
                if byline_info:
                    # Find author link inside bylineInfo
                    author_link = byline_info.find('a', class_='a-link-normal', href=re.compile(r'/e/|/s/', re.I))
                    if author_link:
                        author_text = author_link.get_text(strip=True)
                        if author_text and len(author_text) > 2:
                            details['author'] = author_text
                    else:
                        # Fallback: look for author span
                        author_span = byline_info.find('span', class_='author')
                        if author_span:
                            author_link = author_span.find('a', href=re.compile(r'/e/|/s/', re.I))
                            if author_link:
                                author_text = author_link.get_text(strip=True)
                                if author_text and len(author_text) > 2:
                                    details['author'] = author_text
                
                # Method 2: Look for author link with /e/ or /s/ pattern (fallback)
                if not details['author']:
                    author_elem = soup.find('a', href=re.compile(r'/e/|/s/', re.I))
                    if author_elem:
                        author_text = author_elem.get_text(strip=True)
                        # Filter out non-author links (follow, etc.)
                        if author_text and len(author_text) > 2 and 'follow' not in author_text.lower():
                            # Make sure it's not a series link
                            href = author_elem.get('href', '')
                            if 'series' not in href.lower() and 'saga' not in href.lower():
                                details['author'] = author_text
                
                # Method 3: Look for "by Author Name" pattern
                if not details['author']:
                    by_pattern = soup.find(string=re.compile(r'\bby\s+[A-Z]', re.I))
                    if by_pattern:
                        parent = by_pattern.parent
                        author_link = parent.find('a', href=re.compile(r'/e/|/s/', re.I))
                        if author_link:
                            author_text = author_link.get_text(strip=True)
                            if author_text and len(author_text) > 2:
                                details['author'] = author_text
                
                # Method 4: Look in contributorNameID data attribute
                if not details['author']:
                    author_data = soup.find(attrs={'data-contributor-name-id': True})
                    if author_data:
                        author_text = author_data.get_text(strip=True)
                        if author_text and len(author_text) > 2:
                            details['author'] = author_text
            else:
                # No audiobook link found - log for debugging
                logger.debug(f"No audiobook link found for {kindle_url}")
        
        except Exception as e:
            logger.warning(f"Error extracting Audiobook details: {e}")
        
        return details
    
    def scrape_book_details(self, book_link, rank, genre=''):
        """
        Scrape details for a single book

        Args:
            book_link: URL of the book detail page
            rank: Rank of the book
            genre: Top-level Genre (from amazon_uk.csv / --genre) to copy into output

        Returns:
            Dictionary with all book details
        """
        # ASIN is reliably embedded in the URL — extract once as a baseline so
        # the column is populated even when the detail-bullet block is absent.
        asin_match = re.search(r'(?:/dp/|/gp/product/)([A-Z0-9]{10})', book_link or '')
        asin_from_url = asin_match.group(1) if asin_match else ''

        result = {
            'list_url': '',
            'genre': genre,
            'sub_genre': '',
            'list': '',
            'rank': rank,
            'book_link': book_link,
            'book_type': 'Kindle',
            'book_name': '',
            'series_name': '',
            'amazon_series_link': '',
            'series_length': '',
            'book_publisher': '',
            'amazon_star_rating': '',
            'amazon_rating_count': '',
            'amazon_book_metrics': '',
            'audiobook_link': '',
            'audiobook_name': '',
            'author': '',
            'narrator': '',
            'audiobook_publisher': '',
            'amazon_audiobook_metrics': '',
            # Enrichment fields (populated when self.enrich is True)
            'synopsis': '',
            'print_length': '',
            'publication_date': '',
            'language': '',
            'asin': asin_from_url,
            'isbn': '',
            'book_no_in_series': '',
            'price': '',
            'goodreads_rating': '',
            'goodreads_rating_count': '',
        }
        
        if not self.safe_get(book_link):
            return result
        
        try:
            # Wait for page to load - wait for product title or bylineInfo (optimized)
            try:
                wait = WebDriverWait(self.driver, 10)  # Reduced from 15 to 10
                # Wait for either product title or bylineInfo to be present
                wait.until(EC.any_of(
                    EC.presence_of_element_located((By.ID, 'productTitle')),
                    EC.presence_of_element_located((By.ID, 'bylineInfo')),
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'h1, span.a-size-large'))
                ))
                # Reduced wait for dynamic content
                time.sleep(0.5)  # Reduced from 2 to 0.5
            except TimeoutException:
                logger.warning(f"Timeout waiting for page elements on {book_link}, continuing anyway...")
                # Still try to extract data even if timeout occurred
                time.sleep(0.5)  # Reduced from 2 to 0.5
            
            soup = BeautifulSoup(self.driver.page_source, 'html.parser')

            # Extract Kindle details
            kindle_details = self.extract_kindle_details(soup)
            result.update(kindle_details)

            # Extract enrichment fields (Synopsis, Print Length, Publication Date,
            # Goodreads Rating/Count) BEFORE audiobook extraction, since the
            # audiobook step can navigate away from the Kindle page.
            if self.enrich:
                try:
                    enrichment = self.extract_enrichment_fields()
                    # If the DOM scrape couldn't find an ASIN, keep the one we
                    # already pulled from the URL — never overwrite a real
                    # value with an empty string.
                    if not enrichment.get('asin'):
                        enrichment['asin'] = result.get('asin', '')
                    result.update(enrichment)
                except Exception as e:
                    logger.warning(f"Enrichment failed for {book_link}: {e}")

            # Extract Audiobook details
            audiobook_details = self.extract_audiobook_details(soup, book_link)
            result.update(audiobook_details)
            
            logger.info(f"Scraped details for rank {rank}: {result.get('book_name', 'Unknown')}")
        
        except Exception as e:
            logger.error(f"Error scraping book details for {book_link}: {e}")
        
        return result
    
    def scrape_bestsellers_page(self, base_url, list_type="Paid", sub_genre="",
                                 csv_filename="", checkpoint_file="", resume=False,
                                 genre=""):
        """
        Scrape a bestsellers page (Free or Paid)
        
        Args:
            base_url: Base URL of the bestsellers page
            list_type: "Paid" or "Free"
            sub_genre: Sub-genre name
            csv_filename: CSV filename to check for existing data
            checkpoint_file: Checkpoint file path
            resume: Whether to resume from checkpoint
            
        Returns:
            List of book detail dictionaries
        """
        all_books = []
        
        # Construct URL based on list type
        if list_type.lower() == "free":
            # Try multiple URL patterns for Free list
            if 'ref=pd_zg_hrsr_digital-text' in base_url:
                url = base_url.replace('ref=pd_zg_hrsr_digital-text', 'ref=zg_bs?ie=UTF8&tf=1')
            elif 'ref=zg_bs_nav' in base_url or 'ref=' in base_url:
                # For URLs with ref parameter, replace it with Free list format
                # Pattern: .../ref=zg_bs_nav_digital-text_3_157305011 -> .../ref=zg_bs?ie=UTF8&tf=1
                import re
                url = re.sub(r'ref=[^&]*', 'ref=zg_bs?ie=UTF8&tf=1', base_url)
                # If there was a ? before ref, we might have created ?&, clean it up
                url = url.replace('?&', '&').replace('&&', '&')
            else:
                # Generic Free list URL construction
                if '?' in base_url:
                    url = base_url + '&tf=1'
                else:
                    url = base_url + '?tf=1'
        else:
            url = base_url
        
        logger.info(f"Scraping {list_type} books from {url}")
        
        # Load existing scraped books to avoid duplicates
        scraped_books = set()
        scraped_ranks_by_list = {}
        scraped_book_links_by_list = {}  # Track by book_link and list_type to detect duplicates
        if csv_filename:
            scraped_books, scraped_ranks_by_list, scraped_book_links_by_list = self.load_existing_books(csv_filename)
        
        # Get book links for this specific list type
        scraped_book_links = scraped_book_links_by_list.get(list_type, set())
        
        # Check existing ranks for this list type
        existing_ranks = scraped_ranks_by_list.get(list_type, set())
        if existing_ranks:
            max_existing_rank = max(existing_ranks)
            logger.info(f"Found existing {list_type} entries up to rank {max_existing_rank}")
        logger.info(f"Found {len(scraped_book_links)} unique book links already scraped")
        
        # Load checkpoint if resuming
        start_rank = 0
        if resume and checkpoint_file:
            checkpoint = self.load_checkpoint(checkpoint_file)
            if checkpoint and checkpoint.get('list_url') == url:
                start_rank = checkpoint.get('last_rank', 0)
                logger.info(f"Checkpoint shows last rank: {start_rank}")
        
        # Use the higher of checkpoint rank or max existing rank
        if existing_ranks:
            start_rank = max(start_rank, max(existing_ranks))
            logger.info(f"Starting from rank {start_rank + 1} (skipping ranks 1-{start_rank})")
        
        # Extract book links
        books = self.extract_book_links_from_bestsellers(url)
        
        if not books:
            logger.warning(f"No books found on {url}")
            return all_books
        
        # Check if CSV file exists at the start (for append mode)
        csv_file_exists = os.path.exists(csv_filename) if csv_filename else False
        
        # Scrape details for each book
        last_successful_rank = start_rank
        for book_info in books:
            try:
                book_link = book_info.get('book_link', '')
                rank = book_info.get('rank', 0)
                asin = book_info.get('asin', '')
                
                if not book_link:
                    continue
                
                # Skip if this book_link was already scraped (regardless of rank)
                if book_link in scraped_book_links:
                    logger.info(f"Skipping rank {rank} - book link already scraped: {book_link[:50]}...")
                    last_successful_rank = rank
                    continue
                
                # Skip if rank already exists for this list type
                if rank in existing_ranks:
                    logger.info(f"Skipping rank {rank} - already exists in CSV for {list_type}")
                    last_successful_rank = rank
                    continue
                
                # Skip if already scraped (by URL and book_link)
                if (url, str(rank), book_link) in scraped_books:
                    logger.info(f"Skipping rank {rank} - already scraped")
                    last_successful_rank = rank
                    continue
                
                # Skip if before start rank (when resuming)
                if rank <= start_rank:
                    logger.info(f"Skipping rank {rank} - before start rank {start_rank + 1}")
                    continue
                
                # Scrape book details
                details = self.scrape_book_details(book_link, rank, genre=genre)
                
                # Add metadata
                details['list_url'] = url
                details['genre'] = genre
                details['sub_genre'] = sub_genre
                details['list'] = f"Top 100 {list_type}"
                
                all_books.append(details)
                last_successful_rank = rank
                
                # Mark this book_link as scraped to avoid duplicates
                scraped_book_links.add(book_link)
                
                # Save immediately to CSV
                if csv_filename:
                    self.save_to_csv([details], csv_filename, append=csv_file_exists)
                    csv_file_exists = True  # File exists after first save
                    logger.info(f"Saved rank {rank} to CSV immediately")
                
                # Save checkpoint periodically
                if checkpoint_file and rank % 10 == 0:
                    self.save_checkpoint(checkpoint_file, url, list_type, last_successful_rank)
                
                # Delay between requests
                time.sleep(self.delay)
            
            except KeyboardInterrupt:
                logger.info("Scraping interrupted by user")
                if checkpoint_file:
                    self.save_checkpoint(checkpoint_file, url, list_type, last_successful_rank)
                raise
            except Exception as e:
                logger.error(f"Error processing book {book_info.get('rank', 'unknown')}: {e}")
                # Save checkpoint on error
                if checkpoint_file:
                    self.save_checkpoint(checkpoint_file, url, list_type, last_successful_rank)
                continue
        
        # Final checkpoint save
        if checkpoint_file and last_successful_rank > 0:
            self.save_checkpoint(checkpoint_file, url, list_type, last_successful_rank)
        
        return all_books
    
    def load_existing_books(self, filename):
        """
        Load already scraped books from CSV to avoid duplicates
        
        Args:
            filename: CSV filename
            
        Returns:
            Tuple of (set of (list_url, rank, book_link) tuples, dict of {list_type: set of ranks}, dict of {list_type: set of book_links})
        """
        scraped_books = set()
        scraped_ranks_by_list = {}  # {list_type: set of ranks}
        scraped_book_links_by_list = {}  # {list_type: set of book_links}
        
        if not os.path.exists(filename):
            return scraped_books, scraped_ranks_by_list, scraped_book_links_by_list
        
        try:
            with open(filename, 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    list_url = row.get('List URL', '')
                    rank = row.get('Rank', '')
                    book_link = row.get('Book Link', '')
                    list_type = row.get('List', '')  # e.g., "Top 100 Paid" or "Top 100 Free"
                    
                    if book_link:
                        scraped_books.add((list_url, rank, book_link))
                    
                    # Track ranks and book links by list type
                    if list_type:
                        # Extract "Paid" or "Free" from "Top 100 Paid" or "Top 100 Free"
                        if "Paid" in list_type:
                            list_key = "Paid"
                        elif "Free" in list_type:
                            list_key = "Free"
                        else:
                            list_key = list_type
                        
                        # Track ranks
                        if rank:
                            if list_key not in scraped_ranks_by_list:
                                scraped_ranks_by_list[list_key] = set()
                            try:
                                scraped_ranks_by_list[list_key].add(int(rank))
                            except ValueError:
                                pass
                        
                        # Track book links by list type
                        if book_link:
                            if list_key not in scraped_book_links_by_list:
                                scraped_book_links_by_list[list_key] = set()
                            scraped_book_links_by_list[list_key].add(book_link)
            
            logger.info(f"Loaded {len(scraped_books)} existing books from {filename}")
            for list_type, ranks in scraped_ranks_by_list.items():
                logger.info(f"  {list_type}: ranks {min(ranks) if ranks else 0}-{max(ranks) if ranks else 0}")
        
        except Exception as e:
            logger.warning(f"Error loading existing books: {e}")
        
        return scraped_books, scraped_ranks_by_list, scraped_book_links_by_list
    
    def save_checkpoint(self, checkpoint_file, list_url, list_type, last_rank):
        """
        Save checkpoint to resume later
        
        Args:
            checkpoint_file: Checkpoint JSON file path
            list_url: Current list URL
            list_type: Current list type
            last_rank: Last successfully scraped rank
        """
        try:
            checkpoint = {
                'list_url': list_url,
                'list_type': list_type,
                'last_rank': last_rank,
                'timestamp': time.time()
            }
            with open(checkpoint_file, 'w') as f:
                json.dump(checkpoint, f, indent=2)
            logger.info(f"Checkpoint saved: {list_type} list, last rank: {last_rank}")
        except Exception as e:
            logger.error(f"Error saving checkpoint: {e}")
    
    def load_checkpoint(self, checkpoint_file):
        """
        Load checkpoint to resume scraping
        
        Args:
            checkpoint_file: Checkpoint JSON file path
            
        Returns:
            Dictionary with checkpoint data or None
        """
        if not os.path.exists(checkpoint_file):
            return None
        
        try:
            with open(checkpoint_file, 'r') as f:
                checkpoint = json.load(f)
            logger.info(f"Checkpoint loaded: {checkpoint.get('list_type')} list, last rank: {checkpoint.get('last_rank')}")
            return checkpoint
        except Exception as e:
            logger.warning(f"Error loading checkpoint: {e}")
            return None
    
    def save_to_csv(self, books, filename, append=False):
        """
        Save scraped data to CSV file
        
        Args:
            books: List of book dictionaries
            filename: Output CSV filename
            append: If True, append to existing file; if False, overwrite
        """
        if not books:
            logger.warning("No books to save")
            return
        
        fieldnames = [
            'List URL', 'Genre', 'Sub Genre', 'List', 'Rank', 'Book Link', 'Book Type',
            'Book Name', 'Series Name', 'Amazon Series Link', 'Series Length',
            'Book No. in Series',
            'Book Publisher',
            'Amazon Star Rating', 'Amazon Rating Count', 'Amazon Book Metrics',
            'Audiobook Link', 'Audiobook Name', 'Author', 'Narrator',
            'Audiobook Publisher', 'Amazon Audiobook Metrics',
            'Synopsis', 'Print Length', 'Publication Date',
            'Language', 'ASIN', 'ISBN', 'Price Title (EUR)',
            'Goodreads Rating', 'Goodreads Rating Count',
        ]
        
        file_exists = os.path.exists(filename) and append
        
        try:
            mode = 'a' if append and file_exists else 'w'
            with open(filename, mode, newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                # Write header only if file is new or being overwritten
                if not file_exists:
                    writer.writeheader()
                
                for book in books:
                    row = {
                        'List URL': book.get('list_url', ''),
                        'Genre': book.get('genre', ''),
                        'Sub Genre': book.get('sub_genre', ''),
                        'List': book.get('list', ''),
                        'Rank': book.get('rank', ''),
                        'Book Link': book.get('book_link', ''),
                        'Book Type': book.get('book_type', 'Kindle'),
                        'Book Name': book.get('book_name', ''),
                        'Series Name': book.get('series_name', ''),
                        'Amazon Series Link': book.get('amazon_series_link', ''),
                        'Series Length': book.get('series_length', ''),
                        'Book No. in Series': book.get('book_no_in_series', ''),
                        'Book Publisher': book.get('book_publisher', ''),
                        'Amazon Star Rating': book.get('amazon_star_rating', ''),
                        'Amazon Rating Count': book.get('amazon_rating_count', ''),
                        'Amazon Book Metrics': book.get('amazon_book_metrics', ''),
                        'Audiobook Link': book.get('audiobook_link', ''),
                        'Audiobook Name': book.get('audiobook_name', ''),
                        'Author': book.get('author', ''),
                        'Narrator': book.get('narrator', ''),
                        'Audiobook Publisher': book.get('audiobook_publisher', ''),
                        'Amazon Audiobook Metrics': book.get('amazon_audiobook_metrics', ''),
                        'Synopsis': book.get('synopsis', ''),
                        'Print Length': book.get('print_length', ''),
                        'Publication Date': book.get('publication_date', ''),
                        'Language': book.get('language', ''),
                        'ASIN': book.get('asin', ''),
                        'ISBN': book.get('isbn', ''),
                        'Price Title (EUR)': book.get('price', ''),
                        'Goodreads Rating': book.get('goodreads_rating', ''),
                        'Goodreads Rating Count': book.get('goodreads_rating_count', ''),
                    }
                    writer.writerow(row)
            
            # Only log for multiple books to reduce verbosity (single books are logged in main loop)
            if len(books) > 1:
                action = "Appended" if append and file_exists else "Saved"
                logger.info(f"{action} {len(books)} books to {filename}")
        
        except Exception as e:
            logger.error(f"Error saving to CSV: {e}")


def main():
    """Main function to run the scraper"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Scrape Amazon Kindle and Audiobook data')
    parser.add_argument('--url', type=str, default='',
                       help='Base URL of a single bestsellers page (omit to run from the bulk CSV)')
    parser.add_argument('--sub-genre', type=str, default='',
                       help='Sub-genre name when using --url (e.g., "Cozy mysteries")')
    parser.add_argument('--genre', type=str, default='',
                       help='Run only URLs whose Genre column matches this value. '
                            'Lets you split the workload across terminals — '
                            'each terminal scrapes one Genre. Use --list-genres to see options.')
    parser.add_argument('--list-genres', action='store_true',
                       help='List the Genres available in the bulk CSV and exit')
    parser.add_argument('--bulk-csv', type=str, default=DEFAULT_BULK_CSV,
                       help=f'Path to the bulk-links CSV (default: {DEFAULT_BULK_CSV})')
    parser.add_argument('--output', type=str, default='amazon_results',
                       help='Output folder name for bulk mode, or CSV filename when using --url (default: amazon_results/)')
    parser.add_argument('--list-type', type=str, choices=['Paid', 'Free', 'Both'],
                       default='Both', help='Which list to scrape (default: Both — '
                                            'use Paid or Free to restrict)')
    parser.add_argument('--headless', action='store_true', default=False,
                       help='Run browser in headless mode (default: headful/visible)')
    parser.add_argument('--delay', type=float, default=0.5,
                       help='Delay between requests in seconds (default: 0.5 for faster scraping)')
    parser.add_argument('--resume', action='store_true',
                       help='Resume from checkpoint if available')
    parser.add_argument('--checkpoint', type=str, default='',
                       help='Checkpoint file path (default: <output>.checkpoint.json)')
    parser.add_argument('--no-enrich', dest='enrich', action='store_false', default=True,
                       help='Skip per-book enrichment (Synopsis, Print Length, '
                            'Publication Date, Goodreads Rating/Count). Faster but fewer columns.')
    parser.add_argument('--no-login', dest='login', action='store_false', default=True,
                       help='Skip the manual Amazon login prompt when enrichment is on. '
                            'Goodreads ratings will likely be empty without login.')

    args = parser.parse_args()

    # If the user passed a custom CSV path, reload the genre map from it.
    if args.bulk_csv != DEFAULT_BULK_CSV:
        urls_by_genre = load_urls_by_genre(args.bulk_csv)
    else:
        urls_by_genre = URLS_BY_GENRE

    if args.list_genres:
        if not urls_by_genre:
            print(f"No genres found in {args.bulk_csv}.")
            return
        print(f"Genres available in {args.bulk_csv}:")
        for g in sorted(urls_by_genre.keys()):
            print(f"  {g}  ({len(urls_by_genre[g])} URLs)")
        print("\nRun a single genre with:")
        print('  python amazon_listing.py --genre "<Genre name>"')
        return

    # Build list of (url, sub_genre, genre) triples to process.
    if args.url:
        # In single-URL mode, --genre (if any) labels the Genre column.
        urls_to_process = [(args.url, args.sub_genre, args.genre)]
    elif args.genre:
        if args.genre not in urls_by_genre:
            available = ', '.join(sorted(urls_by_genre.keys())) or '(none)'
            logger.error(f"Genre '{args.genre}' not found in {args.bulk_csv}. Available: {available}")
            return
        urls_to_process = urls_by_genre[args.genre]
        logger.info(f"Selected genre '{args.genre}' — {len(urls_to_process)} URLs")
    else:
        urls_to_process = [triple for entries in urls_by_genre.values() for triple in entries]

    # In bulk mode create the output folder; in single-URL mode use --output as the CSV path directly
    bulk_mode = not args.url
    if bulk_mode:
        # Put each genre's results in its own subfolder so concurrent terminal
        # runs don't trample each other's checkpoints.
        if args.genre:
            output_folder = os.path.join(args.output, _genre_slug(args.genre))
        else:
            output_folder = args.output
        os.makedirs(output_folder, exist_ok=True)
        logger.info(f"Output folder: {output_folder}/")

    scraper = AmazonScraper(headless=args.headless, delay=args.delay, enrich=args.enrich)

    # When enrichment is on, prompt the user to log in to Amazon so the
    # Goodreads rating widget renders on each book detail page.
    if args.enrich and args.login:
        scraper.amazon_login_flow()

    try:
        total_scraped = 0

        for idx, (url, sub_genre, genre) in enumerate(urls_to_process, start=1):
            logger.info(f"[{idx}/{len(urls_to_process)}] Processing: {genre} / {sub_genre} — {url}")

            # Derive per-URL CSV path
            if bulk_mode:
                slug = _bestseller_slug(url, sub_genre=sub_genre, genre=genre)
                csv_filename = os.path.join(output_folder, f"{slug}.csv")
                checkpoint_file = os.path.join(output_folder, f"{slug}.checkpoint.json")
            else:
                csv_filename = args.output if args.output.endswith('.csv') else args.output + '.csv'
                checkpoint_file = args.checkpoint or csv_filename.replace('.csv', '.checkpoint.json')

            if args.list_type in ['Paid', 'Both']:
                logger.info("Scraping Paid books...")
                paid_books = scraper.scrape_bestsellers_page(
                    url, list_type="Paid", sub_genre=sub_genre, genre=genre,
                    csv_filename=csv_filename, checkpoint_file=checkpoint_file, resume=args.resume
                )
                total_scraped += len(paid_books) if paid_books else 0

            if args.list_type in ['Free', 'Both']:
                logger.info("Scraping Free books...")
                free_books = scraper.scrape_bestsellers_page(
                    url, list_type="Free", sub_genre=sub_genre, genre=genre,
                    csv_filename=csv_filename, checkpoint_file=checkpoint_file, resume=args.resume
                )
                total_scraped += len(free_books) if free_books else 0

        # Final summary
        if total_scraped:
            logger.info(f"Scraping completed! Total books scraped in this session: {total_scraped}")
        else:
            logger.warning("No new books were scraped")

    except KeyboardInterrupt:
        logger.info("Scraping interrupted by user")
    except Exception as e:
        logger.error(f"Error during scraping: {e}")
    finally:
        scraper.close()


if __name__ == '__main__':
    main()


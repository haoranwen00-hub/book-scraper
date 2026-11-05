"""
Book Scraper - scrape product data from books.toscrape.com and export to Excel / CSV.

Usage:
    python scraper.py                      # scrape all 50 list pages (1000 books)
    python scraper.py --pages 3            # only the first 3 pages
    python scraper.py --details            # also visit each book page (category, UPC, stock, description)
    python scraper.py --output result      # custom output file name (result.xlsx / result.csv)

Requires Python 3.10+.
"""

import argparse
import csv
import logging
import re
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
START_URL = urljoin(BASE_URL, "catalogue/page-1.html")
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

# Fields added by --details. Every book gets these defaults first, so all rows
# keep the same columns even when a detail page fails.
DETAIL_DEFAULTS = {"category": "", "upc": "", "stock_count": 0, "description": ""}

# Cell values starting with these characters can be run as formulas by Excel.
FORMULA_PREFIXES = ("=", "+", "-", "@")

log = logging.getLogger("book_scraper")


# ---------------------------------------------------------------- networking
def make_session(pool_size: int) -> requests.Session:
    """HTTP session with a browser-like User-Agent and automatic retries."""
    session = requests.Session()
    session.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
    )
    retry = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    # pool size matches the number of worker threads, so connections are reused
    session.mount("https://", HTTPAdapter(max_retries=retry, pool_maxsize=pool_size))
    return session


def get_soup(session: requests.Session, url: str) -> BeautifulSoup:
    resp = session.get(url, timeout=15)
    resp.raise_for_status()
    resp.encoding = "utf-8"  # the site declares no charset; force UTF-8 so "£" decodes correctly
    return BeautifulSoup(resp.text, "html.parser")


# ------------------------------------------------------------------- parsing
def parse_price(text: str) -> float:
    match = re.search(r"[\d.]+", text)
    return float(match.group()) if match else 0.0


def parse_rating(card) -> int:
    """Read the star rating from a class like "star-rating Three"; 0 if missing."""
    element = card.select_one("p.star-rating")
    classes = element.get("class", []) if element else []
    return next((RATING_MAP[c] for c in classes if c in RATING_MAP), 0)


def parse_card(card, page_url: str) -> dict:
    link = card.select_one("h3 a")
    if link is None:
        raise ValueError("card has no title link")
    price = card.select_one("p.price_color")
    availability = card.select_one("p.availability")
    return {
        "title": (link.get("title") or link.text).strip(),
        "price_gbp": parse_price(price.text) if price else 0.0,
        "rating": parse_rating(card),
        "in_stock": bool(availability and "In stock" in availability.text),
        "url": urljoin(page_url, link["href"]),
    }


def parse_list_page(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extract every book card on a list page. A broken card is skipped, not fatal."""
    books = []
    for card in soup.select("article.product_pod"):
        try:
            books.append(parse_card(card, page_url))
        except (ValueError, KeyError) as exc:
            log.warning("Skipped a card on %s: %s", page_url, exc)
    return books


def parse_detail_page(soup: BeautifulSoup) -> dict:
    """Extract extra fields from a single book page. Missing parts become defaults."""
    table = {
        row.th.text.strip(): row.td.text.strip()
        for row in soup.select("table.table tr")
        if row.th and row.td
    }
    stock = re.search(r"\d+", table.get("Availability", ""))
    desc = soup.select_one("#product_description + p")
    crumbs = soup.select("ul.breadcrumb li a")
    return {
        "category": crumbs[-1].text.strip() if len(crumbs) >= 3 else "",
        "upc": table.get("UPC", ""),
        "stock_count": int(stock.group()) if stock else 0,
        "description": desc.text.strip() if desc else "",
    }


# ------------------------------------------------------------------ scraping
def scrape_list(session: requests.Session, max_pages: int | None, delay: float) -> list[dict]:
    """Follow the "next" button page by page.

    If a page fails after all retries, stop and keep the books collected so far.
    """
    books, url, page = [], START_URL, 0
    while url and (max_pages is None or page < max_pages):
        page += 1
        try:
            soup = get_soup(session, url)
        except requests.RequestException as exc:
            log.error("Page %d failed, stopping here and keeping %d books: %s", page, len(books), exc)
            break
        items = parse_list_page(soup, url)
        books.extend(items)
        log.info("Page %d: %d books (total %d)", page, len(items), len(books))
        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None
        if url:
            time.sleep(delay)  # be polite to the server
    return books


def enrich_with_details(session: requests.Session, books: list[dict], workers: int) -> int:
    """Visit every book page concurrently and merge detail fields into the records.

    Returns the number of books whose detail page failed (they keep default values).
    """
    for book in books:
        book.update(DETAIL_DEFAULTS)

    failed = 0
    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(get_soup, session, b["url"]): b for b in books}
        for i, future in enumerate(as_completed(futures), 1):
            book = futures[future]
            try:
                book.update(parse_detail_page(future.result()))
            except Exception as exc:  # network or parsing error: skip this book, keep the rest
                failed += 1
                log.warning("Detail page failed (%s): %s", book["url"], exc)
            if i % 100 == 0 or i == len(books):
                log.info("Details: %d/%d", i, len(books))
    return failed


# ------------------------------------------------------------------- export
def clean_value(value):
    """Make scraped text safe for Excel/CSV.

    Removes control characters that Excel cannot store, and prefixes text that
    starts with = + - @ with an apostrophe, so it is shown as text and never run as a formula.
    """
    if not isinstance(value, str):
        return value
    value = ILLEGAL_CHARACTERS_RE.sub("", value)
    if value.startswith(FORMULA_PREFIXES):
        value = "'" + value
    return value


def export_csv(books: list[dict], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: opens cleanly in Excel
        writer = csv.DictWriter(f, fieldnames=list(books[0].keys()))
        writer.writeheader()
        for b in books:
            writer.writerow({k: clean_value(v) for k, v in b.items()})


def style_sheet(ws, col_widths: dict[str, int]) -> None:
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for idx, cell in enumerate(ws[1], 1):
        ws.column_dimensions[get_column_letter(idx)].width = col_widths.get(cell.value, 14)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def build_summary(books: list[dict]) -> tuple[str, list[list]]:
    """Group by category if any book has one, otherwise by star rating."""
    key = "category" if any(b.get("category") for b in books) else "rating"
    groups = defaultdict(list)
    for b in books:
        groups[b.get(key) or "Unknown"].append(b)
    rows = []
    for name, items in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        prices = [b["price_gbp"] for b in items]
        rows.append([
            name,
            len(items),
            round(sum(prices) / len(prices), 2),
            min(prices),
            max(prices),
            round(sum(b["rating"] for b in items) / len(items), 2),
        ])
    return key, rows


def export_excel(books: list[dict], path: Path) -> None:
    wb = Workbook()

    # Sheet 1: raw data
    ws = wb.active
    ws.title = "Books"
    headers = list(books[0].keys())
    ws.append(headers)
    for b in books:
        ws.append([clean_value(b[h]) for h in headers])
    for row in ws.iter_rows(min_row=2, min_col=headers.index("price_gbp") + 1,
                            max_col=headers.index("price_gbp") + 1):
        row[0].number_format = "£#,##0.00"
    style_sheet(ws, {"title": 50, "url": 60, "description": 80, "upc": 20})

    # Sheet 2: summary
    key, rows = build_summary(books)
    ss = wb.create_sheet("Summary")
    ss.append([key.capitalize(), "Books", "Avg Price (£)", "Min Price (£)", "Max Price (£)", "Avg Rating"])
    for r in rows:
        ss.append([clean_value(v) for v in r])
    style_sheet(ss, {"Category": 22})

    wb.save(path)


# ---------------------------------------------------------------------- main
def output_paths(output: str) -> tuple[Path, Path]:
    """Add .xlsx/.csv to the output name without cutting names like "run.v2"."""
    out = Path(output)
    if out.suffix.lower() in (".xlsx", ".csv"):  # user typed "books.xlsx" -> use "books"
        out = out.with_suffix("")
    return out.parent / f"{out.name}.xlsx", out.parent / f"{out.name}.csv"


def main() -> None:
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com to Excel/CSV.")
    parser.add_argument("--pages", type=int, default=None, help="max list pages to scrape (default: all)")
    parser.add_argument("--details", action="store_true", help="also scrape each book's detail page")
    parser.add_argument("--workers", type=int, default=8, help="concurrent requests for detail pages")
    parser.add_argument("--delay", type=float, default=0.3, help="seconds to wait between list pages")
    parser.add_argument("--output", default="output/books", help="output path without extension")
    args = parser.parse_args()

    if args.pages is not None and args.pages < 1:
        parser.error("--pages must be at least 1")
    if args.workers < 1:
        parser.error("--workers must be at least 1")
    if args.delay < 0:
        parser.error("--delay cannot be negative")

    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s",
                        datefmt="%H:%M:%S")

    start = time.time()
    session = make_session(pool_size=args.workers)
    books = scrape_list(session, args.pages, args.delay)
    if not books:
        log.error("No books found - check the network connection or whether the site layout changed.")
        return
    failed = 0
    if args.details:
        failed = enrich_with_details(session, books, args.workers)
        # put the long URL column last for readability
        books = [{k: v for k, v in b.items() if k != "url"} | {"url": b["url"]} for b in books]

    xlsx_path, csv_path = output_paths(args.output)
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)
    export_excel(books, xlsx_path)
    export_csv(books, csv_path)
    log.info("Done: %d books in %.1fs -> %s, %s", len(books), time.time() - start, xlsx_path, csv_path)
    if failed:
        log.warning("%d detail pages failed; those rows have empty detail fields.", failed)


if __name__ == "__main__":
    main()

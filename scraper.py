"""
Book Scraper - scrape product data from books.toscrape.com and export to Excel / CSV.

Usage:
    python scraper.py                      # scrape all 50 list pages (1000 books)
    python scraper.py --pages 3            # only the first 3 pages
    python scraper.py --details            # also visit each book page (category, UPC, stock, description)
    python scraper.py --output result      # custom output file name (result.xlsx / result.csv)
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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
START_URL = urljoin(BASE_URL, "catalogue/page-1.html")
RATING_MAP = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

log = logging.getLogger("book_scraper")


# ---------------------------------------------------------------- networking
def make_session() -> requests.Session:
    """HTTP session with a browser-like User-Agent and automatic retries."""
    session = requests.Session()
    session.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
    )
    retry = Retry(total=3, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    session.mount("https://", HTTPAdapter(max_retries=retry, pool_maxsize=10))
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


def parse_list_page(soup: BeautifulSoup, page_url: str) -> list[dict]:
    """Extract the basic info of every book card on a list page."""
    books = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        rating_word = card.select_one("p.star-rating")["class"][1]
        books.append({
            "title": link["title"].strip(),
            "price_gbp": parse_price(card.select_one("p.price_color").text),
            "rating": RATING_MAP.get(rating_word, 0),
            "in_stock": "In stock" in card.select_one("p.availability").text,
            "url": urljoin(page_url, link["href"]),
        })
    return books


def parse_detail_page(soup: BeautifulSoup) -> dict:
    """Extract extra fields from a single book page."""
    table = {row.th.text.strip(): row.td.text.strip() for row in soup.select("table.table tr")}
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
    """Follow the "next" button page by page."""
    books, url, page = [], START_URL, 0
    while url and (max_pages is None or page < max_pages):
        page += 1
        soup = get_soup(session, url)
        items = parse_list_page(soup, url)
        books.extend(items)
        log.info("Page %d: %d books (total %d)", page, len(items), len(books))
        next_link = soup.select_one("li.next a")
        url = urljoin(url, next_link["href"]) if next_link else None
        time.sleep(delay)  # be polite to the server
    return books


def enrich_with_details(session: requests.Session, books: list[dict], workers: int) -> None:
    """Visit every book page concurrently and merge detail fields into the records."""
    def fetch(book):
        return book, parse_detail_page(get_soup(session, book["url"]))

    with ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(fetch, b) for b in books]
        for i, future in enumerate(as_completed(futures), 1):
            try:
                book, details = future.result()
                book.update(details)
            except requests.RequestException as exc:
                log.warning("Detail page failed: %s", exc)
            if i % 100 == 0 or i == len(books):
                log.info("Details: %d/%d", i, len(books))


# ------------------------------------------------------------------- export
def export_csv(books: list[dict], path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as f:  # utf-8-sig: opens cleanly in Excel
        writer = csv.DictWriter(f, fieldnames=list(books[0].keys()))
        writer.writeheader()
        writer.writerows(books)


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
    """Group by category if available, otherwise by star rating."""
    key = "category" if books[0].get("category") else "rating"
    groups = defaultdict(list)
    for b in books:
        groups[b[key]].append(b)
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
        ws.append([b[h] for h in headers])
    for row in ws.iter_rows(min_row=2, min_col=headers.index("price_gbp") + 1,
                            max_col=headers.index("price_gbp") + 1):
        row[0].number_format = "£#,##0.00"
    style_sheet(ws, {"title": 50, "url": 60, "description": 80, "upc": 20})

    # Sheet 2: summary
    key, rows = build_summary(books)
    ss = wb.create_sheet("Summary")
    ss.append([key.capitalize(), "Books", "Avg Price (£)", "Min Price (£)", "Max Price (£)", "Avg Rating"])
    for r in rows:
        ss.append(r)
    style_sheet(ss, {"Category": 22})

    wb.save(path)


# ---------------------------------------------------------------------- main
def main() -> None:
    parser = argparse.ArgumentParser(description="Scrape books.toscrape.com to Excel/CSV.")
    parser.add_argument("--pages", type=int, default=None, help="max list pages to scrape (default: all)")
    parser.add_argument("--details", action="store_true", help="also scrape each book's detail page")
    parser.add_argument("--workers", type=int, default=8, help="concurrent requests for detail pages")
    parser.add_argument("--delay", type=float, default=0.3, help="seconds to wait between list pages")
    parser.add_argument("--output", default="output/books", help="output path without extension")
    args = parser.parse_args()

    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)s  %(message)s",
                        datefmt="%H:%M:%S")

    start = time.time()
    session = make_session()
    books = scrape_list(session, args.pages, args.delay)
    if not books:
        log.error("No books found - has the site layout changed?")
        return
    if args.details:
        enrich_with_details(session, books, args.workers)
        # put the long URL column last for readability
        books = [{k: v for k, v in b.items() if k != "url"} | {"url": b["url"]} for b in books]

    out = Path(args.output)
    out.parent.mkdir(parents=True, exist_ok=True)
    export_excel(books, out.with_suffix(".xlsx"))
    export_csv(books, out.with_suffix(".csv"))
    log.info("Done: %d books in %.1fs -> %s(.xlsx/.csv)", len(books), time.time() - start, out)


if __name__ == "__main__":
    main()

# Book Scraper — Web Data to Excel

A Python web scraper that collects product data from an e-commerce catalogue
([books.toscrape.com](https://books.toscrape.com), a sandbox site built for scraping practice)
and exports clean, ready-to-use **Excel** and **CSV** files.

It is a template for typical freelance data-extraction tasks: product lists, prices,
stock levels, categories, and descriptions, delivered as a formatted spreadsheet.

## Features

- **Automatic pagination**: follows the "next" button until the last page (50 pages, 1,000 products)
- **Detail-page scraping** (`--details`): category, UPC, exact stock count, full description
- **Concurrent requests** with a thread pool: all 1,000 detail pages in about 75 seconds
- **Reliable**: automatic retries with back-off on network errors and 429/5xx responses, request timeouts, and a delay between list pages
- **Fault-tolerant**: a failed page or a malformed product is logged and skipped, so one error never loses the data already collected
- **Safe Excel output**: control characters are removed and formula-like text (`=`, `+`, `-`, `@`) is escaped, which prevents CSV/Excel formula injection
- **Data cleaning**: prices parsed to numbers, star ratings converted to 1–5, stock text converted to integers
- **Formatted Excel output**:
  - `Books` sheet: styled header, frozen top row, auto-filter, currency format
  - `Summary` sheet: count, average/min/max price, and average rating per category
- **CSV output** in UTF-8 with BOM, so it opens correctly in Excel
- **Command-line options** for page limit, worker count, delay, and output path

## Sample Output

The full result is in [`sample_output/`](sample_output/): `books.xlsx` and `books.csv`, 1,000 rows.

| title | price_gbp | rating | in_stock | category | upc | stock_count |
|---|---|---|---|---|---|---|
| A Light in the Attic | 51.77 | 3 | True | Poetry | a897fe39b1053632 | 22 |
| Tipping the Velvet | 53.74 | 1 | True | Historical Fiction | 90fa61229261140a | 20 |

Summary sheet (top categories by book count):

| Category | Books | Avg Price (£) | Min | Max | Avg Rating |
|---|---|---|---|---|---|
| Default | 152 | 34.39 | 10.23 | 58.87 | 2.84 |
| Nonfiction | 110 | 34.26 | 11.68 | 59.90 | 2.88 |
| Sequential Art | 75 | 34.57 | 10.16 | 57.62 | 2.97 |
| Fiction | 65 | 36.07 | 10.60 | 59.98 | 3.18 |

## Quick Start

```bash
pip install -r requirements.txt

python scraper.py                    # all pages, basic fields
python scraper.py --details          # all pages + detail pages
python scraper.py --pages 3          # first 3 pages only
python scraper.py --output data/run1 # writes data/run1.xlsx and data/run1.csv
```

| Option | Default | Description |
|---|---|---|
| `--pages N` | all | Maximum number of list pages |
| `--details` | off | Visit each product page for extra fields |
| `--workers N` | 8 | Concurrent requests for detail pages |
| `--delay S` | 0.3 | Seconds to wait between list pages |
| `--output PATH` | `output/books` | Output path without extension |

## How It Works

1. `scrape_list()` downloads each list page, parses every product card with BeautifulSoup CSS selectors, and follows the "next" link.
2. `enrich_with_details()` fetches product pages in parallel with `ThreadPoolExecutor` and merges the extra fields into each record.
3. `export_excel()` / `export_csv()` write the data, and `build_summary()` creates the per-category statistics sheet.

## Tech Stack

Python 3.10+ · requests · BeautifulSoup4 · openpyxl · concurrent.futures

## Adapting to Other Websites

To scrape a different static website, change `START_URL` and the CSS selectors in
`parse_list_page()` / `parse_detail_page()`. Pagination, retries, concurrency, and
Excel export work without further changes. Always check a site's terms of service
and `robots.txt` before scraping it.

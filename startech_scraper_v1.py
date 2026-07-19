#!/usr/bin/env python3
"""
Star Tech Product Scraper
=========================
Scrapes product specification and description HTML from startech.com.bd.

DEPENDENCIES:
    pip install requests beautifulsoup4 openpyxl pandas

USAGE:
    1. Create an Excel file (e.g., urls.xlsx) with URLs in the FIRST column.
       The first row should be a header (e.g., "URL").
    2. Run: python scraper.py
    3. Results are saved to an auto-timestamped Excel file in the output folder.

OUTPUT COLUMNS:
    - URL
    - Product Name (from <title> tag)
    - Specification HTML (full outer HTML of #specification section)
    - Description HTML (full outer HTML of #description section)
    - Status (success / error)
    - Error Message
"""

import os
import sys
import time
import re
from datetime import datetime

import requests
from bs4 import BeautifulSoup
import pandas as pd

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
INPUT_FILE = "urls.xlsx"          # Excel file with URLs in the first column
OUTPUT_DIR = "output"             # Directory for output files
REQUEST_DELAY = 2                 # Seconds between requests (polite scraping)
TIMEOUT = 30                      # Request timeout in seconds
MAX_RETRIES = 2                   # Number of retries per URL

# Realistic browser headers to avoid blocks
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/126.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "Cache-Control": "max-age=0",
}

# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def ensure_dir(path: str) -> str:
    """Create directory if it doesn't exist."""
    os.makedirs(path, exist_ok=True)
    return path


def fetch_page(url: str, retries: int = MAX_RETRIES) -> str:
    """
    Fetch the HTML content of a URL with retries.
    Raises Exception on failure.
    """
    last_exception = None
    for attempt in range(1, retries + 1):
        try:
            response = requests.get(
                url,
                headers=HEADERS,
                timeout=TIMEOUT,
                allow_redirects=True,
            )
            response.raise_for_status()
            return response.text
        except requests.exceptions.HTTPError as e:
            last_exception = f"HTTP {e.response.status_code}: {e.response.reason}"
        except requests.exceptions.ConnectionError as e:
            last_exception = f"Connection error: {str(e)}"
        except requests.exceptions.Timeout:
            last_exception = "Request timed out"
        except requests.exceptions.RequestException as e:
            last_exception = f"Request failed: {str(e)}"

        if attempt < retries:
            time.sleep(REQUEST_DELAY * attempt)  # Exponential-ish backoff

    raise Exception(last_exception)


def extract_title(soup: BeautifulSoup) -> str:
    """Extract product name from <title> tag."""
    title_tag = soup.find("title")
    if title_tag:
        return title_tag.get_text(strip=True)
    return ""


def extract_specification_html(soup: BeautifulSoup) -> str:
    """
    Extract the full outer HTML of the specification section.
    Target: <section class="specification-tab m-tb-10" id="specification">
    """
    spec = soup.find("section", {"id": "specification", "class": "specification-tab"})
    if spec:
        return str(spec)
    # Fallback: try by id only
    spec = soup.find("section", {"id": "specification"})
    if spec:
        return str(spec)
    return ""


def extract_description_html(soup: BeautifulSoup) -> str:
    """
    Extract the full outer HTML of the description section.
    Target: <section class="description bg-white m-tb-15" id="description">
    """
    desc = soup.find("section", {"id": "description", "class": "description"})
    if desc:
        return str(desc)
    # Fallback: try by id only
    desc = soup.find("section", {"id": "description"})
    if desc:
        return str(desc)
    return ""


def read_urls_from_excel(filepath: str) -> list:
    """
    Read URLs from the first column of an Excel file.
    Skips the header row and empty cells.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Input file not found: {filepath}")

    df = pd.read_excel(filepath, header=0)
    if df.empty:
        raise ValueError("Input Excel file is empty.")

    # First column
    urls = df.iloc[:, 0].dropna().astype(str).tolist()
    # Basic URL validation
    urls = [u.strip() for u in urls if u.strip().lower().startswith("http")]
    return urls


# ---------------------------------------------------------------------------
# MAIN SCRAPER
# ---------------------------------------------------------------------------

def scrape_urls(urls: list) -> list:
    """
    Scrape each URL and return a list of result dictionaries.
    """
    results = []
    total = len(urls)

    for idx, url in enumerate(urls, start=1):
        print(f"[{idx}/{total}] Scraping: {url}")
        row = {
            "URL": url,
            "Product Name": "",
            "Specification HTML": "",
            "Description HTML": "",
            "Status": "error",
            "Error Message": "",
        }

        try:
            html = fetch_page(url)
            soup = BeautifulSoup(html, "html.parser")

            row["Product Name"] = extract_title(soup)
            row["Specification HTML"] = extract_specification_html(soup)
            row["Description HTML"] = extract_description_html(soup)

            # Validate extraction
            if not row["Specification HTML"] and not row["Description HTML"]:
                row["Error Message"] = (
                    "Both specification and description sections not found. "
                    "Page structure may have changed."
                )
            elif not row["Specification HTML"]:
                row["Error Message"] = (
                    "Specification section not found. "
                    "Partial data extracted (description only)."
                )
            elif not row["Description HTML"]:
                row["Error Message"] = (
                    "Description section not found. "
                    "Partial data extracted (specification only)."
                )
            else:
                row["Status"] = "success"
                row["Error Message"] = ""

        except Exception as e:
            row["Error Message"] = str(e)
            print(f"    ERROR: {e}")

        results.append(row)

        # Polite delay between requests (skip after last item)
        if idx < total:
            time.sleep(REQUEST_DELAY)

    return results


def save_results(results: list) -> str:
    """
    Save results to an auto-generated Excel file with timestamp.
    Returns the output file path.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"scraped_results_{timestamp}.xlsx"
    out_path = os.path.join(ensure_dir(OUTPUT_DIR), filename)

    df = pd.DataFrame(results)
    # Reorder columns to match desired output
    column_order = [
        "URL",
        "Product Name",
        "Specification HTML",
        "Description HTML",
        "Status",
        "Error Message",
    ]
    df = df[column_order]

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
        worksheet = writer.sheets["Results"]

        # Auto-adjust column widths for readability
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            adjusted_width = min(max_length + 2, 100)  # Cap at 100
            worksheet.column_dimensions[column_letter].width = adjusted_width

    return out_path


# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------

def main():
    print("=" * 60)
    print("Star Tech Product Scraper")
    print("=" * 60)
    print(f"Input file : {INPUT_FILE}")
    print(f"Output dir : {OUTPUT_DIR}")
    print(f"Delay      : {REQUEST_DELAY}s")
    print("=" * 60)

    # Read URLs
    try:
        urls = read_urls_from_excel(INPUT_FILE)
    except Exception as e:
        print(f"\nFailed to read input file: {e}")
        sys.exit(1)

    if not urls:
        print("\nNo valid URLs found in the input file.")
        sys.exit(1)

    print(f"\nFound {len(urls)} URL(s) to scrape.\n")

    # Scrape
    results = scrape_urls(urls)

    # Save
    out_file = save_results(results)

    # Summary
    success_count = sum(1 for r in results if r["Status"] == "success")
    error_count = len(results) - success_count

    print("\n" + "=" * 60)
    print("SCRAPING COMPLETE")
    print("=" * 60)
    print(f"Total URLs : {len(results)}")
    print(f"Success    : {success_count}")
    print(f"Errors     : {error_count}")
    print(f"Output     : {out_file}")
    print("=" * 60)


if __name__ == "__main__":
    main()

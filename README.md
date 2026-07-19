# Startech_scraper
A scraper to scrap specification and description data from Startech website  Scrapes product specification tables from startech.com.bd product pages.

# DEPENDENCIES:
  pip install requests beautifulsoup4 openpyxl pandas

# USAGE:
  1. Create an Excel file named "urls.xlsx" with URLs in the FIRST column.
       The first row should be a header (e.g., "URL").
  2. Run: python startech_scraper.py
  3. Results are saved to output/scraped_results_YYYYMMDD_HHMMSS.xlsx

# OUTPUT COLUMNS:
  - URL
    - Product Name (from <title>)
    - Specification HTML (full outer HTML of #specification section)
    - Description HTML (full outer HTML of #description section)
    - Status (success / error)
    - Error Message
# CONFIGURATION
# ---------------------------------------------------------------------------
INPUT_FILE = "urls.xlsx"          # Excel file with URLs in the first column
OUTPUT_DIR = "output"             # Directory for output files
REQUEST_DELAY = 2                 # Seconds between requests (polite scraping)
TIMEOUT = 30                      # Request timeout in seconds
MAX_RETRIES = 2                   # Number of retries per URL
# IMPORTANT: Do NOT include "br" in Accept-Encoding. The server sends Brotli (Error of V1 fixed in V2)
# which Python requests cannot decode without the extra "brotli" package.
# Keeping it to "gzip, deflate" ensures reliable decoding.



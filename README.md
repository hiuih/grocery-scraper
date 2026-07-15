# Grocery Price Scraper

Scrapes **every product** from Fresh St. Market and Save On Foods and exports them to two Excel files with product name, product number, regular price, and promo price.

## Output

| File | Contents |
|------|----------|
| `Fresh_St_Market_Products.xlsx` | All Fresh St. Market products |
| `Save_On_Foods_Products.xlsx` | All Save On Foods products (if accessible) |

Each Excel file has four columns:

| Product Name | Product Number | Regular Price | Promo Price |
|---|---|---|---|
| Activia Probiotic Yogurt - Blueberry 650g | 5680034596 | $6.29 | $4.29 |
| Dan-D Pak - Almonds - Salted 350g | 77079570079 | $12.59 | |

Promo prices appear in **bold red** when a product is on sale.

---

## Setup (brand new computer)

### Requirements
- Windows 10 or 11
- Internet connection

### Steps

1. Download or clone this repository
2. Double-click **`setup.bat`**
3. Follow the on-screen instructions — it handles everything automatically:
   - Installs Python (if not already installed)
   - Installs required packages (`playwright`, `openpyxl`)
   - Downloads the Chromium browser (~130 MB, one time only)
   - Creates a **"Run Grocery Scraper"** shortcut on your Desktop

> You only need to run `setup.bat` **once** per computer.

---

## Running the scraper

After setup, just double-click **`Run Grocery Scraper.bat`** on your Desktop. It runs both scrapers back to back.

- Fresh St. Market takes **30–40 minutes** (~213 categories, ~8,500 products)
- Save-On-Foods takes **1–3 hours** (~815 categories, ~22,000 products) — it's slower because it has to work around Cloudflare bot protection
- Two Excel files are saved to your Desktop when done
- Fresh St. Market runs fully in the background (no window)
- Save-On-Foods runs headless too, but may briefly **open a visible browser window** once or twice per run if Cloudflare needs re-verifying. Usually it clears itself in a few seconds; if it ever shows a checkbox or puzzle, just click it and the scraper carries on automatically.

---

## How it works

| Site | Status |
|------|--------|
| Fresh St. Market | Fully supported — scrapes all ~213 categories |
| Save On Foods | Fully supported — scrapes all ~815 categories, works around Cloudflare automatically |

Both scripts use **Playwright**-based headless Chromium to load each category page, paginate through results, and extract:
- Product name
- Product number (UPC/barcode)
- Regular price
- Promo/sale price (if currently on sale)

Save-On-Foods sits behind Cloudflare bot protection, so `save_on_foods_scraper.py` uses `patchright` (a stealth Playwright build) instead of plain Playwright, and falls back to a real, visible browser window to clear the challenge if headless attempts are rejected — then reuses that session's cookies for the rest of the run.

---

## Files

```
grocery-scraper/
├── grocery_scraper.py         ← Fresh St. Market scraper
├── save_on_foods_scraper.py   ← Save-On-Foods scraper
├── setup.bat                  ← one-click installer for new computers
├── Run Grocery Scraper.bat    ← runs both scrapers
└── README.md
```

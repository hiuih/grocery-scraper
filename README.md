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

After setup, just double-click **`Run Grocery Scraper.bat`** on your Desktop.

- Takes **15–30 minutes** to complete
- Two Excel files are saved to your Desktop when done
- The browser runs in the background (you won't see it)

---

## How it works

| Site | Status |
|------|--------|
| Fresh St. Market | Fully supported — scrapes all 22 categories |
| Save On Foods | Attempted — may be blocked by Cloudflare bot protection |

The scraper uses **Playwright** (a headless Chromium browser) to load each page, scroll to reveal lazy-loaded products, and extract:
- Product name
- Product number (UPC/barcode)
- Regular price
- Promo/sale price (if currently on sale)

---

## Files

```
grocery-scraper/
├── grocery_scraper.py   ← main scraper script
├── setup.bat            ← one-click installer for new computers
└── README.md
```

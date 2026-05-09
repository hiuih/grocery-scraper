#!/usr/bin/env python3
"""
GROCERY PRICE SCRAPER
======================
Scrapes every product from:
  - Fresh St. Market  ->  Fresh_St_Market_Products.xlsx
  - Save On Foods     ->  Save_On_Foods_Products.xlsx

HOW TO RUN:
  1. Make sure Python is installed (python.org)
  2. Double-click this file  OR  open a terminal and run:
         python grocery_scraper.py
  3. The first run downloads a small browser (~130 MB) one time only
  4. Wait 15-30 minutes for it to finish
  5. Two Excel files will appear on your Desktop
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# ─────────────────────────────────────────────────────────────────
#  STEP 1: Auto-install packages + browser
# ─────────────────────────────────────────────────────────────────
import subprocess

def pip_install(pkg):
    r = subprocess.run([sys.executable, "-m", "pip", "install", pkg, "-q"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        print(f"  ERROR installing {pkg}: {r.stderr.strip()}")
        sys.exit(1)

print("=" * 60)
print("  GROCERY PRICE SCRAPER")
print("=" * 60)
print()
print("Checking dependencies...")
for pkg in ["playwright", "openpyxl"]:
    try:
        __import__(pkg)
        print(f"  [OK] {pkg}")
    except ImportError:
        print(f"  Installing {pkg}...")
        pip_install(pkg)
        print(f"  [OK] {pkg} installed.")

print("  Checking browser (one-time ~130 MB download if needed)...")
r = subprocess.run([sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
                   capture_output=True, text=True)
print("  [OK] Browser ready.\n")

# ─────────────────────────────────────────────────────────────────
#  STEP 2: Imports
# ─────────────────────────────────────────────────────────────────
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re, time, os
from datetime import datetime

# ─────────────────────────────────────────────────────────────────
#  STEP 3: Settings
# ─────────────────────────────────────────────────────────────────
OUTPUT_DIR    = os.path.join(os.path.expanduser("~"), "Desktop")
FRESH_ST_FILE = os.path.join(OUTPUT_DIR, "Fresh_St_Market_Products.xlsx")
SAVE_ON_FILE  = os.path.join(OUTPUT_DIR, "Save_On_Foods_Products.xlsx")

FRESH_ST_BASE = "https://www.freshstmarket.com/sm/pickup/rsid/055"
SAVE_ON_BASE  = "https://www.saveonfoods.com"

FRESH_ST_CATEGORIES = [
    ("deli-prepared-foods-id-3568",  "Deli & Prepared Foods"),
    ("baby-id-3571",                 "Baby"),
    ("personal-care-items-id-3605",  "Personal Care Items"),
    ("dairy-eggs-cheese-id-3622",    "Dairy, Eggs & Cheese"),
    ("condiments-sauces-id-3657",    "Condiments & Sauces"),
    ("crackers-cookies-id-3659",     "Crackers & Cookies"),
    ("grains-pasta-sides-id-3664",   "Grains, Pasta & Sides"),
    ("international-id-3671",        "International"),
    ("bakery-id-3673",               "Bakery"),
    ("bulk-id-3675",                 "Bulk"),
    ("beverages-id-3691",            "Beverages"),
    ("frozen-food-id-3708",          "Frozen Food"),
    ("household-goods-id-3717",      "Household Goods"),
    ("breakfast-cereals-id-3752",    "Breakfast & Cereals"),
    ("canned-goods-id-3756",         "Canned Goods"),
    ("snacks-candy-id-3777",         "Snacks & Candy"),
    ("pet-care-id-3788",             "Pet Care"),
    ("wellness-products-id-3795",    "Wellness Products"),
    ("meat-id-3796",                 "Meat"),
    ("produce-id-3803",              "Produce"),
    ("seafood-id-3804",              "Seafood"),
    ("spices-baking-id-3806",        "Spices & Baking"),
]

# ─────────────────────────────────────────────────────────────────
#  STEP 4: Helpers
# ─────────────────────────────────────────────────────────────────
CARD_SEL  = "[data-testid^='ProductCardWrapper']"
NAME_SEL  = "[class*='ProductCardNameWrapper--']"
PRICE_SEL = "[class*='ProductPrice--']"
WAS_SEL   = "[class*='ProductWasPrice--']"

UNIT_PATTERN = re.compile(
    r",\s*\d*\.?\d+\s*(Each|Gram|Kilogram|ml|L|kg|g|lb|oz|pk|pack)\b.*$",
    flags=re.IGNORECASE
)

def scroll_to_load_all(page, max_wait=25):
    """Scroll until no new product cards appear."""
    deadline = time.time() + max_wait
    prev = 0
    stall = 0
    while time.time() < deadline:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(1.2)
        count = page.evaluate(
            f"document.querySelectorAll(\"{CARD_SEL}\").length"
        )
        if count == prev:
            stall += 1
            if stall >= 3:
                break
        else:
            stall = 0
        prev = count
    page.evaluate("window.scrollTo(0, 0)")   # scroll back to top


def extract_products(page):
    """Pull all product cards from the current loaded page."""
    scroll_to_load_all(page)
    cards   = page.query_selector_all(CARD_SEL)
    results = []
    for card in cards:
        try:
            raw_id = card.get_attribute("data-testid").replace("ProductCardWrapper-", "")
            pnum   = raw_id.lstrip("0") or raw_id

            name_el  = card.query_selector(NAME_SEL)
            price_el = card.query_selector(PRICE_SEL)
            was_el   = card.query_selector(WAS_SEL)

            if not name_el:
                continue

            name = name_el.inner_text().strip()
            name = name.replace("Open Product Description", "").strip()
            name = UNIT_PATTERN.sub("", name).strip()

            current  = price_el.inner_text().strip() if price_el else ""
            was_raw  = was_el.inner_text().strip()   if was_el   else ""
            was      = re.sub(r"^was\s*", "", was_raw, flags=re.IGNORECASE).strip()

            regular = was     if was else current
            promo   = current if was else ""

            results.append({
                "Product Name":   name,
                "Product Number": pnum,
                "Regular Price":  regular,
                "Promo Price":    promo,
            })
        except Exception:
            continue
    return results


def has_next_page(page, current_page):
    """Return True if a next page exists beyond current_page."""
    try:
        # Look for a pagination link with number higher than current_page
        links = page.query_selector_all("a[class*='Pagination'], button[class*='Pagination']")
        for link in links:
            txt = link.inner_text().strip()
            if txt.isdigit() and int(txt) > current_page:
                return True
        # Also check for a "Next" button
        nxt = page.query_selector("a[aria-label*='Next'], button[aria-label*='Next']")
        if nxt and nxt.is_enabled():
            return True
    except Exception:
        pass
    return False


def open_browser(playwright):
    browser = playwright.chromium.launch(headless=True)
    ctx = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        viewport={"width": 1280, "height": 900},
        locale="en-CA",
    )
    # Only block heavy media — keep JS/CSS/fonts so React renders
    ctx.route("**/*.{mp4,mp3,avi,wmv}", lambda r: r.abort())
    return browser, ctx


# ─────────────────────────────────────────────────────────────────
#  STEP 5: Fresh St. Market scraper
# ─────────────────────────────────────────────────────────────────
def scrape_fresh_st(playwright):
    print("=" * 60)
    print("  Scraping FRESH ST. MARKET")
    print("=" * 60)

    browser, ctx = open_browser(playwright)
    page = ctx.new_page()
    seen = {}   # product_number -> dict (deduplication)

    for cat_i, (cat_slug, cat_name) in enumerate(FRESH_ST_CATEGORIES, 1):
        print(f"\n  [{cat_i}/{len(FRESH_ST_CATEGORIES)}] {cat_name}")
        page_num = 1
        cat_new  = 0

        while True:
            skip = (page_num - 1) * 30
            url  = f"{FRESH_ST_BASE}/categories/{cat_slug}?page={page_num}&skip={skip}"

            try:
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                time.sleep(2.5)   # let React hydrate
            except Exception as e:
                print(f"    [!] Load error p{page_num}: {e}")
                break

            products = extract_products(page)
            if not products:
                break

            for p in products:
                pn = p["Product Number"]
                if pn not in seen:
                    seen[pn] = p
                    cat_new += 1
                else:
                    # Update promo price if we now have one
                    if p["Promo Price"] and not seen[pn]["Promo Price"]:
                        seen[pn]["Promo Price"]   = p["Promo Price"]
                        seen[pn]["Regular Price"] = p["Regular Price"]

            print(f"    Page {page_num}: {len(products):>3} items  "
                  f"(running total: {len(seen)})", end="\r")

            if not has_next_page(page, page_num):
                break
            page_num += 1

        print(f"    [OK] {cat_new:>4} new products  (total: {len(seen)})          ")

    browser.close()
    print(f"\n  Done! {len(seen)} unique products.\n")
    return list(seen.values())


# ─────────────────────────────────────────────────────────────────
#  STEP 6: Save On Foods scraper
# ─────────────────────────────────────────────────────────────────
SAVE_ON_DEPTS = [
    ("Bakery",          "/store/browse/Bakery/_/N-1b6s"),
    ("Beverages",       "/store/browse/Beverages/_/N-1b6u"),
    ("Dairy / Eggs",    "/store/browse/Dairy-Eggs-Cheese/_/N-1b6v"),
    ("Deli",            "/store/browse/Deli/_/N-1b70"),
    ("Frozen Foods",    "/store/browse/Frozen-Foods/_/N-1b6z"),
    ("Grocery",         "/store/browse/Grocery/_/N-1b71"),
    ("Health & Beauty", "/store/browse/Health-Beauty/_/N-1b72"),
    ("Household",       "/store/browse/Household/_/N-1b73"),
    ("Meat & Seafood",  "/store/browse/Meat-Seafood/_/N-1b74"),
    ("Natural/Organic", "/store/browse/Natural-Organic/_/N-1b75"),
    ("Produce",         "/store/browse/Produce/_/N-1b76"),
    ("Specialty",       "/store/browse/Specialty-Foods/_/N-1b77"),
]

def extract_save_on(page):
    """Extract products from a Save On Foods page."""
    try:
        page.wait_for_selector("[data-product-id], .product-tile", timeout=12000)
    except PWTimeout:
        return []

    cards = (page.query_selector_all("[data-product-id]") or
             page.query_selector_all(".product-tile") or [])

    products = []
    for card in cards:
        try:
            pnum = (card.get_attribute("data-product-id") or
                    card.get_attribute("data-sku") or "").strip()
            name_el = (card.query_selector("[class*='name']") or
                       card.query_selector("[class*='title']") or
                       card.query_selector("h3") or card.query_selector("h2"))
            name    = name_el.inner_text().strip() if name_el else ""

            reg_el  = (card.query_selector("[class*='regular']") or
                       card.query_selector("[class*='was']") or
                       card.query_selector("[class*='original']"))
            regular = reg_el.inner_text().strip() if reg_el else ""

            promo_el = (card.query_selector("[class*='sale']") or
                        card.query_selector("[class*='promo']") or
                        card.query_selector("[class*='now']"))
            promo    = promo_el.inner_text().strip() if promo_el else ""

            if not regular and not promo:
                pe = card.query_selector("[class*='price']")
                regular = pe.inner_text().strip() if pe else ""

            if name or pnum:
                products.append({"Product Name": name, "Product Number": pnum,
                                  "Regular Price": regular, "Promo Price": promo})
        except Exception:
            continue
    return products


def scrape_save_on(playwright):
    print("=" * 60)
    print("  Scraping SAVE ON FOODS")
    print("=" * 60)
    print("  Opening browser...")

    browser, ctx = open_browser(playwright)
    page = ctx.new_page()

    try:
        page.goto(SAVE_ON_BASE, wait_until="domcontentloaded", timeout=20000)
    except Exception as e:
        print(f"  [!] Could not connect to Save On Foods: {e}")
        browser.close()
        return []

    body = page.content()
    if "blocked" in body.lower() or "ray id" in body.lower():
        print("  [!] Save On Foods is blocking automated access.")
        print("      Their Cloudflare settings prevent scraping from any tool.")
        browser.close()
        return []

    print("  [OK] Connected.\n")
    all_products = {}

    for dept_name, dept_path in SAVE_ON_DEPTS:
        print(f"  Department: {dept_name}")
        page_num = 1

        while True:
            url = f"{SAVE_ON_BASE}{dept_path}?Nrpp=48&No={(page_num-1)*48}"
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=25000)
                time.sleep(2)
            except Exception as e:
                print(f"    [!] Error: {e}")
                break

            if "blocked" in page.content().lower():
                print("    [!] Blocked mid-scrape.")
                break

            products = extract_save_on(page)
            if not products:
                break

            for p in products:
                key = p["Product Number"] or p["Product Name"]
                if key and key not in all_products:
                    all_products[key] = p

            print(f"    Page {page_num}: {len(products)} items  "
                  f"(total: {len(all_products)})", end="\r")

            if len(products) < 48:
                break
            page_num += 1

        print(f"    [OK] {dept_name}  (total: {len(all_products)})          ")

    browser.close()
    print(f"\n  Done! {len(all_products)} unique products.\n")
    return list(all_products.values())


# ─────────────────────────────────────────────────────────────────
#  STEP 7: Excel builder
# ─────────────────────────────────────────────────────────────────
def build_excel(products, filepath, store_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = store_name[:31]

    def bdr():
        s = Side(style="thin", color="BDD7EE")
        return Border(left=s, right=s, top=s, bottom=s)

    # Title
    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value     = f"{store_name}  -  Product Price List"
    c.font      = Font(bold=True, size=14, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Timestamp
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value     = f"Scraped on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
    c.font      = Font(italic=True, size=9, color="888888")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 15

    # Count
    ws.merge_cells("A3:D3")
    c = ws["A3"]
    c.value     = f"Total products: {len(products):,}"
    c.font      = Font(italic=True, size=9, color="888888")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[3].height = 14

    # Headers
    for i, col in enumerate(["Product Name", "Product Number", "Regular Price", "Promo Price"], 1):
        c = ws.cell(row=4, column=i, value=col)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
    ws.row_dimensions[4].height = 22

    alt = PatternFill("solid", fgColor="DEEAF1")
    for row_i, p in enumerate(products, 5):
        is_alt = (row_i % 2 == 0)
        cells = [
            ws.cell(row=row_i, column=1, value=p.get("Product Name",   "")),
            ws.cell(row=row_i, column=2, value=p.get("Product Number", "")),
            ws.cell(row=row_i, column=3, value=p.get("Regular Price",  "")),
            ws.cell(row=row_i, column=4, value=p.get("Promo Price",    "")),
        ]
        for cell in cells:
            cell.border    = bdr()
            cell.alignment = Alignment(vertical="center")
            if is_alt:
                cell.fill = alt

        cells[0].alignment = Alignment(horizontal="left",   vertical="center")
        cells[1].alignment = Alignment(horizontal="center", vertical="center")
        cells[1].font      = Font(name="Courier New", size=9)
        cells[2].alignment = Alignment(horizontal="right",  vertical="center")
        cells[3].alignment = Alignment(horizontal="right",  vertical="center")
        if p.get("Promo Price"):
            cells[3].font = Font(bold=True, color="C00000")

    for i, w in enumerate([52, 18, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:D{4 + len(products)}"

    wb.save(filepath)
    print(f"  [SAVED] {filepath}")


# ─────────────────────────────────────────────────────────────────
#  STEP 8: Main
# ─────────────────────────────────────────────────────────────────
def main():
    start = time.time()
    print(f"  Output folder: {OUTPUT_DIR}")
    print(f"  Started at:    {datetime.now().strftime('%I:%M %p')}\n")

    with sync_playwright() as pw:
        fresh = scrape_fresh_st(pw)
        if fresh:
            print("  Building Fresh St. Market Excel file...")
            build_excel(fresh, FRESH_ST_FILE, "Fresh St. Market")
        else:
            print("  [!] No Fresh St. Market products collected.")

        print()

        save_on = scrape_save_on(pw)
        if save_on:
            print("  Building Save On Foods Excel file...")
            build_excel(save_on, SAVE_ON_FILE, "Save On Foods")
        else:
            print("  [!] No Save On Foods products - file not created.")

    mins, secs = divmod(int(time.time() - start), 60)
    print()
    print("=" * 60)
    print("  ALL DONE!")
    print("=" * 60)
    if fresh:
        print(f"  Fresh St. Market : {len(fresh):>6,} products  ->  Excel saved")
    if save_on:
        print(f"  Save On Foods    : {len(save_on):>6,} products  ->  Excel saved")
    print(f"  Time taken       : {mins}m {secs}s")
    print(f"  Files on Desktop : {OUTPUT_DIR}")
    print()

    try:
        input("  Press ENTER to close this window...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()

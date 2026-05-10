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
  3. The first run downloads a browser (~130 MB) one time only
  4. Wait 15-30 minutes for it to finish
  5. Two Excel files will appear on your Desktop
"""

# ── Fix console encoding without replacing stdout (avoids blank-window bug) ──
import sys, os

try:
    if hasattr(sys.stdout, "reconfigure"):          # Python 3.7+
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        import io
        sys.stdout = io.TextIOWrapper(
            sys.stdout.buffer, encoding="utf-8",
            errors="replace", line_buffering=True
        )
        sys.stderr = io.TextIOWrapper(
            sys.stderr.buffer, encoding="utf-8",
            errors="replace", line_buffering=True
        )
except Exception:
    pass


def p(msg="", flush=True):
    """Print with guaranteed flush so the window never looks blank."""
    print(msg, flush=flush)


# ─────────────────────────────────────────────────────────────────
#  STEP 1: Auto-install packages + browser
# ─────────────────────────────────────────────────────────────────
import subprocess

p("=" * 60)
p("  GROCERY PRICE SCRAPER")
p("=" * 60)
p()
p("Checking dependencies...")


def pip_install(pkg):
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", pkg, "-q"],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        subprocess.run([sys.executable, "-m", "pip", "install", pkg])


for pkg, import_name in [("playwright", "playwright"), ("openpyxl", "openpyxl")]:
    try:
        __import__(import_name)
        p(f"  [OK] {pkg}")
    except ImportError:
        p(f"  Installing {pkg}...")
        pip_install(pkg)
        p(f"  [OK] {pkg} installed.")

p("  Checking browser (downloads ~130 MB first time if needed)...")
subprocess.run(
    [sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
    capture_output=True, text=True
)
p("  [OK] Browser ready.")
p()

# ─────────────────────────────────────────────────────────────────
#  STEP 2: Imports
# ─────────────────────────────────────────────────────────────────
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import re, time, json
    from datetime import datetime
    from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
except ImportError as e:
    p(f"  ERROR: Could not import required module: {e}")
    p("  Try running setup.bat again, or run this in a terminal:")
    p("    pip install playwright openpyxl")
    p("    python -m playwright install chromium")
    p()
    try:
        input("  Press ENTER to close...")
    except EOFError:
        pass
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────
#  STEP 3: Find the real Desktop (OneDrive-aware)
# ─────────────────────────────────────────────────────────────────
def find_desktop():
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
        )
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        if os.path.isdir(desktop):
            return desktop
    except Exception:
        pass

    home = os.path.expanduser("~")
    for path in [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive - Personal", "Desktop"),
        os.path.join(home, "OneDrive - Commercial", "Desktop"),
    ]:
        if os.path.isdir(path):
            return path

    fallback = os.path.join(home, "Desktop")
    os.makedirs(fallback, exist_ok=True)
    return fallback


OUTPUT_DIR    = find_desktop()
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
#  STEP 4: CSS selectors for DOM fallback
# ─────────────────────────────────────────────────────────────────
CARD_SEL  = "[data-testid^='ProductCardWrapper']"
NAME_SEL  = "[class*='ProductCardNameWrapper--']"
PRICE_SEL = "[class*='ProductPrice--']"
WAS_SEL   = "[class*='ProductWasPrice--']"
UNIT_RE   = re.compile(
    r",\s*\d*\.?\d+\s*(Each|Gram|Kilogram|ml|L|kg|g|lb|oz|pk|pack)\b.*$",
    flags=re.IGNORECASE
)


def make_browser(playwright):
    browser = playwright.chromium.launch(headless=True)
    ctx = browser.new_context(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        viewport={"width": 1280, "height": 900},
        locale="en-CA",
    )
    # Only block heavy media — JS/CSS/fonts must load for React
    ctx.route("**/*.{mp4,mp3,avi,wmv}", lambda r: r.abort())
    return browser, ctx


# ─────────────────────────────────────────────────────────────────
#  STEP 5: API interception helpers
# ─────────────────────────────────────────────────────────────────
def _looks_like_products(lst):
    """Return True if a list looks like a product array."""
    if not lst or not isinstance(lst, list):
        return False
    s = lst[0]
    if not isinstance(s, dict):
        return False
    keys_lower = {k.lower() for k in s.keys()}
    product_signals = {"sku", "upc", "name", "price", "productid", "displayname",
                       "itemid", "regularpricestring", "pricestring"}
    return bool(keys_lower & product_signals)


def _fmt_price(v):
    """Format a raw price value as '$X.XX' or empty string."""
    if v is None:
        return ""
    v = str(v).strip().replace(",", "")
    if not v or v.lower() in ("0", "0.0", "none", "null", ""):
        return ""
    if v.startswith("$"):
        return v
    try:
        return f"${float(v):.2f}"
    except ValueError:
        return v


def _parse_api_item(item):
    """Convert one raw API product dict to our standard dict."""
    pnum = str(
        item.get("sku") or item.get("upc") or item.get("productId") or
        item.get("id") or item.get("itemId") or ""
    ).strip().lstrip("0") or str(item.get("id", ""))

    name = str(
        item.get("name") or item.get("displayName") or item.get("title") or
        item.get("productName") or item.get("description") or ""
    ).strip()

    if not name and not pnum:
        return None

    # Price – try several nested structures
    regular = promo = ""
    pd = item.get("price") or item.get("prices") or item.get("pricing") or {}
    if isinstance(pd, dict):
        reg_raw = (pd.get("regular") or pd.get("originalPrice") or
                   pd.get("listPrice") or pd.get("was") or pd.get("regularPrice") or "")
        sale_raw = (pd.get("sale") or pd.get("promo") or pd.get("salePrice") or
                    pd.get("now") or pd.get("promotionalPrice") or "")
        if not reg_raw:
            reg_raw = pd.get("current") or pd.get("value") or pd.get("amount") or ""
        regular = _fmt_price(reg_raw)
        promo   = _fmt_price(sale_raw)
    elif isinstance(pd, (int, float)):
        regular = _fmt_price(pd)
    else:
        # Try flat fields on the item itself
        regular = _fmt_price(
            item.get("regularPrice") or item.get("listPrice") or
            item.get("basePrice") or item.get("normalPrice") or ""
        )
        promo = _fmt_price(
            item.get("salePrice") or item.get("promoPrice") or
            item.get("discountedPrice") or ""
        )

    return {
        "Product Name":   name,
        "Product Number": pnum,
        "Regular Price":  regular,
        "Promo Price":    promo,
    }


def fetch_all_via_api(page, first_url, first_data, products_key):
    """
    Given the first API response, paginate through all remaining pages
    by modifying the skip/offset parameter and using page.request.get().
    Returns a list of parsed product dicts.
    """
    raw = list(first_data[products_key])
    page_size = len(raw)
    if page_size == 0:
        return []

    # Try to get total count
    total = None
    for tk in ("totalCount", "total", "totalItems", "count", "numFound",
               "totalResults", "itemCount", "recordCount"):
        if tk in first_data and isinstance(first_data[tk], int):
            total = first_data[tk]
            break

    p(f"    API: {page_size} items on first call, total={total}")

    if not total or total <= page_size:
        results = [r for r in (_parse_api_item(i) for i in raw) if r]
        return results

    parsed = urlparse(first_url)
    params = {k: v[0] if len(v) == 1 else v
              for k, v in parse_qs(parsed.query, keep_blank_values=True).items()}

    # Detect which parameter controls offset
    skip_key = None
    for sk in ("skip", "Skip", "offset", "Offset", "start", "from"):
        if sk in params:
            skip_key = sk
            break
    if not skip_key:
        skip_key = "skip"

    page_key = None
    for pk in ("page", "Page", "pageNumber", "pageNum", "currentPage"):
        if pk in params:
            page_key = pk
            break

    skip = page_size
    page_n = 2
    consecutive_empty = 0

    while skip < total and consecutive_empty < 3:
        params[skip_key] = str(skip)
        if page_key:
            params[page_key] = str(page_n)

        new_url = urlunparse(parsed._replace(query=urlencode(params, doseq=True)))
        try:
            resp = page.request.get(new_url, timeout=20000)
            data = resp.json()
            items = data.get(products_key, [])
            if not items:
                consecutive_empty += 1
                skip += page_size
                page_n += 1
                continue
            consecutive_empty = 0
            raw.extend(items)
            skip += len(items)
            page_n += 1
            p(f"    API page {page_n - 1}: +{len(items)} items  "
              f"(collected: {len(raw)}/{total})")
        except Exception as e:
            p(f"    API page error at skip={skip}: {e}")
            consecutive_empty += 1
            skip += page_size
            page_n += 1

    results = [r for r in (_parse_api_item(i) for i in raw) if r]
    return results


# ─────────────────────────────────────────────────────────────────
#  STEP 6: DOM fallback — incremental scroll extraction
# ─────────────────────────────────────────────────────────────────
def extract_dom_incremental(page, max_stall=10):
    """
    Scroll in small steps and extract product cards at each position.
    Works for both virtual lists (react-window) and regular infinite scroll.
    Deduplicates by product number so no item is counted twice.
    """
    seen = {}

    try:
        page.wait_for_selector(CARD_SEL, timeout=15000)
    except Exception:
        return []

    time.sleep(1)

    scroll_pos = 0
    step = 400          # pixels per scroll tick
    stall = 0
    last_height = 0

    while stall < max_stall:
        prev = len(seen)

        for card in page.query_selector_all(CARD_SEL):
            try:
                raw_id = card.get_attribute("data-testid") or ""
                pnum   = raw_id.replace("ProductCardWrapper-", "").lstrip("0") or raw_id
                if not pnum or pnum in seen:
                    continue

                ne = card.query_selector(NAME_SEL)
                if not ne:
                    continue
                name = ne.inner_text().strip()
                name = name.replace("Open Product Description", "").strip()
                name = UNIT_RE.sub("", name).strip()

                pe  = card.query_selector(PRICE_SEL)
                we  = card.query_selector(WAS_SEL)
                cur = pe.inner_text().strip() if pe else ""
                was_raw = we.inner_text().strip() if we else ""
                was = re.sub(r"^was\s*", "", was_raw, flags=re.IGNORECASE).strip()

                seen[pnum] = {
                    "Product Name":   name,
                    "Product Number": pnum,
                    "Regular Price":  was if was else cur,
                    "Promo Price":    cur if was else "",
                }
            except Exception:
                continue

        stall = 0 if len(seen) > prev else stall + 1

        scroll_pos += step
        page.evaluate(f"window.scrollTo(0, {scroll_pos})")
        time.sleep(0.6)

        cur_height = page.evaluate("document.body.scrollHeight")
        if scroll_pos >= cur_height and cur_height == last_height:
            # Reached the true bottom — do one final extraction pass
            for card in page.query_selector_all(CARD_SEL):
                try:
                    raw_id = card.get_attribute("data-testid") or ""
                    pnum   = raw_id.replace("ProductCardWrapper-", "").lstrip("0") or raw_id
                    if not pnum or pnum in seen:
                        continue
                    ne = card.query_selector(NAME_SEL)
                    if not ne:
                        continue
                    name = ne.inner_text().strip()
                    name = name.replace("Open Product Description", "").strip()
                    name = UNIT_RE.sub("", name).strip()
                    pe  = card.query_selector(PRICE_SEL)
                    we  = card.query_selector(WAS_SEL)
                    cur = pe.inner_text().strip() if pe else ""
                    was_raw = we.inner_text().strip() if we else ""
                    was = re.sub(r"^was\s*", "", was_raw, flags=re.IGNORECASE).strip()
                    seen[pnum] = {
                        "Product Name":   name,
                        "Product Number": pnum,
                        "Regular Price":  was if was else cur,
                        "Promo Price":    cur if was else "",
                    }
                except Exception:
                    continue
            break
        last_height = cur_height

    return list(seen.values())


# ─────────────────────────────────────────────────────────────────
#  STEP 7: Fresh St. Market scraper
# ─────────────────────────────────────────────────────────────────
def scrape_fresh_st(playwright):
    p("=" * 60)
    p("  Scraping FRESH ST. MARKET")
    p("=" * 60)

    browser, ctx = make_browser(playwright)
    page = ctx.new_page()
    all_products = {}   # pnum -> product dict

    for cat_i, (cat_slug, cat_name) in enumerate(FRESH_ST_CATEGORIES, 1):
        p(f"\n  [{cat_i}/{len(FRESH_ST_CATEGORIES)}] {cat_name}")

        # ── Set up API response interception ──────────────────────
        intercepted = []   # will hold (url, data, products_key) once found

        def on_response(resp, _intercepted=intercepted):
            if _intercepted:           # already found one, stop looking
                return
            try:
                if resp.status != 200:
                    return
                if "json" not in resp.headers.get("content-type", ""):
                    return
                data = resp.json()
                if not isinstance(data, dict):
                    return
                for key in ("products", "items", "data", "results", "records", "catalogItems"):
                    val = data.get(key)
                    if _looks_like_products(val):
                        _intercepted.append((resp.url, data, key))
                        return
            except Exception:
                pass

        page.on("response", on_response)

        url = f"{FRESH_ST_BASE}/categories/{cat_slug}"
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=35000)
            time.sleep(3.5)
        except Exception as e:
            p(f"    [!] Load error: {e}")
            page.remove_listener("response", on_response)
            continue

        # Scroll once to trigger any lazy API calls
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(1)

        page.remove_listener("response", on_response)

        # ── Choose extraction method ──────────────────────────────
        if intercepted:
            first_url, first_data, products_key = intercepted[0]
            p(f"    Using API (intercepted from {first_url[:80]}...)")
            cat_prods = fetch_all_via_api(page, first_url, first_data, products_key)
        else:
            p("    API not intercepted — using DOM scroll fallback")
            cat_prods = extract_dom_incremental(page)

            # DOM fallback: also try URL pagination (skip by 30)
            if cat_prods:
                page_num = 2
                consecutive_empty = 0
                while consecutive_empty < 2:
                    skip = (page_num - 1) * 30
                    paged_url = f"{FRESH_ST_BASE}/categories/{cat_slug}?page={page_num}&skip={skip}"
                    try:
                        page.goto(paged_url, wait_until="domcontentloaded", timeout=30000)
                        time.sleep(2.5)
                    except Exception:
                        break
                    extra = extract_dom_incremental(page)
                    known_pnums = {pr["Product Number"] for pr in cat_prods}
                    new_extra = [pr for pr in extra if pr["Product Number"] not in known_pnums]
                    if not new_extra:
                        consecutive_empty += 1
                    else:
                        consecutive_empty = 0
                        cat_prods.extend(new_extra)
                        p(f"    DOM page {page_num}: +{len(new_extra)} items  "
                          f"(cat total: {len(cat_prods)})")
                    page_num += 1

        # ── Merge into global dict ────────────────────────────────
        cat_new = 0
        for pr in cat_prods:
            pn = pr["Product Number"]
            if pn not in all_products:
                all_products[pn] = pr
                cat_new += 1
            elif pr.get("Promo Price") and not all_products[pn].get("Promo Price"):
                all_products[pn]["Promo Price"]   = pr["Promo Price"]
                all_products[pn]["Regular Price"]  = pr["Regular Price"]

        p(f"    Done: {cat_new} new  |  {len(cat_prods)} from this category  "
          f"|  running total: {len(all_products)}")

    browser.close()
    p(f"\n  Fresh St. finished. {len(all_products)} unique products.\n")
    return list(all_products.values())


# ─────────────────────────────────────────────────────────────────
#  STEP 8: Save On Foods scraper
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
    try:
        page.wait_for_selector("[data-product-id], .product-tile", timeout=12000)
    except PWTimeout:
        return []

    cards = page.query_selector_all("[data-product-id]") or \
            page.query_selector_all(".product-tile") or []
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
            promo   = promo_el.inner_text().strip() if promo_el else ""
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
    p("=" * 60)
    p("  Scraping SAVE ON FOODS")
    p("=" * 60)
    p("  Opening browser...")

    browser, ctx = make_browser(playwright)
    page = ctx.new_page()

    try:
        page.goto(SAVE_ON_BASE, wait_until="domcontentloaded", timeout=20000)
    except Exception as e:
        p(f"  [!] Could not connect: {e}")
        browser.close()
        return []

    content = page.content().lower()
    if "blocked" in content or "ray id" in content or "cloudflare" in content:
        p("  [!] Save On Foods is blocking automated access.")
        p("      Their Cloudflare protection prevents scraping.")
        browser.close()
        return []

    p("  Connected. Starting...\n")
    all_products = {}

    for dept_name, dept_path in SAVE_ON_DEPTS:
        p(f"  Department: {dept_name}")
        page_num = 1

        while True:
            url = f"{SAVE_ON_BASE}{dept_path}?Nrpp=48&No={(page_num-1)*48}"
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=25000)
                time.sleep(2)
            except Exception as e:
                p(f"    [!] Error: {e}")
                break

            if "blocked" in page.content().lower():
                p("    [!] Blocked.")
                break

            products = extract_save_on(page)
            if not products:
                break

            for pr in products:
                key = pr["Product Number"] or pr["Product Name"]
                if key and key not in all_products:
                    all_products[key] = pr

            p(f"    Page {page_num}: {len(products)} items  (total: {len(all_products)})")

            if len(products) < 48:
                break
            page_num += 1

        p(f"    Done: {dept_name}  (total: {len(all_products)})")

    browser.close()
    p(f"\n  Save On Foods finished. {len(all_products)} unique products.\n")
    return list(all_products.values())


# ─────────────────────────────────────────────────────────────────
#  STEP 9: Excel builder
# ─────────────────────────────────────────────────────────────────
def build_excel(products, filepath, store_name):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = store_name[:31]

    def bdr():
        s = Side(style="thin", color="BDD7EE")
        return Border(left=s, right=s, top=s, bottom=s)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value     = f"{store_name}  -  Product Price List"
    c.font      = Font(bold=True, size=14, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value     = f"Scraped on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}"
    c.font      = Font(italic=True, size=9, color="888888")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 15

    ws.merge_cells("A3:D3")
    c = ws["A3"]
    c.value     = f"Total products: {len(products):,}"
    c.font      = Font(italic=True, size=9, color="888888")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[3].height = 14

    for i, col in enumerate(["Product Name", "Product Number",
                              "Regular Price", "Promo Price"], 1):
        c = ws.cell(row=4, column=i, value=col)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
    ws.row_dimensions[4].height = 22

    alt = PatternFill("solid", fgColor="DEEAF1")
    for row_i, pr in enumerate(products, 5):
        cells = [
            ws.cell(row=row_i, column=1, value=pr.get("Product Name",   "")),
            ws.cell(row=row_i, column=2, value=pr.get("Product Number", "")),
            ws.cell(row=row_i, column=3, value=pr.get("Regular Price",  "")),
            ws.cell(row=row_i, column=4, value=pr.get("Promo Price",    "")),
        ]
        for cell in cells:
            cell.border    = bdr()
            cell.alignment = Alignment(vertical="center")
            if row_i % 2 == 0:
                cell.fill = alt

        cells[0].alignment = Alignment(horizontal="left",   vertical="center")
        cells[1].alignment = Alignment(horizontal="center", vertical="center")
        cells[1].font      = Font(name="Courier New", size=9)
        cells[2].alignment = Alignment(horizontal="right",  vertical="center")
        cells[3].alignment = Alignment(horizontal="right",  vertical="center")
        if pr.get("Promo Price"):
            cells[3].font = Font(bold=True, color="C00000")

    for i, w in enumerate([52, 18, 16, 16], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:D{4 + len(products)}"

    wb.save(filepath)
    p(f"  [SAVED] {filepath}")


# ─────────────────────────────────────────────────────────────────
#  STEP 10: Main
# ─────────────────────────────────────────────────────────────────
def main():
    start = time.time()
    p(f"  Output folder: {OUTPUT_DIR}")
    p(f"  Started at:    {datetime.now().strftime('%I:%M %p')}")
    p()

    fresh = []
    save_on = []

    try:
        with sync_playwright() as pw:
            fresh = scrape_fresh_st(pw)
            if fresh:
                p("  Building Fresh St. Market Excel file...")
                build_excel(fresh, FRESH_ST_FILE, "Fresh St. Market")
            else:
                p("  [!] No Fresh St. Market products collected.")

            p()

            save_on = scrape_save_on(pw)
            if save_on:
                p("  Building Save On Foods Excel file...")
                build_excel(save_on, SAVE_ON_FILE, "Save On Foods")
            else:
                p("  [!] No Save On Foods products - file not created.")

    except KeyboardInterrupt:
        p("\n  Stopped by user.")
    except Exception as e:
        p(f"\n  UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
        p()
        p("  Please send the error above to whoever set this up.")

    mins, secs = divmod(int(time.time() - start), 60)
    p()
    p("=" * 60)
    p("  ALL DONE!")
    p("=" * 60)
    if fresh:
        p(f"  Fresh St. Market : {len(fresh):>6,} products  ->  Excel saved")
    if save_on:
        p(f"  Save On Foods    : {len(save_on):>6,} products  ->  Excel saved")
    p(f"  Time taken       : {mins}m {secs}s")
    p(f"  Files saved to   : {OUTPUT_DIR}")
    p()

    try:
        input("  Press ENTER to close this window...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()

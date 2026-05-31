#!/usr/bin/env python3
"""
GROCERY PRICE SCRAPER
======================
Scrapes every product from:
  - Fresh St. Market  ->  Fresh_St_Market_Products.xlsx
  - Save On Foods     ->  Save_On_Foods_Products.xlsx

HOW TO RUN:
  Double-click "Run Grocery Scraper.bat"  OR  open a terminal and run:
      python grocery_scraper.py
  Results appear on your Desktop when finished (~30-90 min).
"""

# ── Fix console encoding ──────────────────────────────────────────
import sys, os

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    elif hasattr(sys.stdout, "buffer"):
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8",
                                      errors="replace", line_buffering=True)
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8",
                                      errors="replace", line_buffering=True)
except Exception:
    pass


def p(msg="", flush=True):
    print(msg, flush=flush)


# ─────────────────────────────────────────────────────────────────
#  Auto-install packages + browser
# ─────────────────────────────────────────────────────────────────
import subprocess

p("=" * 60)
p("  GROCERY PRICE SCRAPER")
p("=" * 60)
p()
p("Checking dependencies...")


def pip_install(pkg):
    r = subprocess.run([sys.executable, "-m", "pip", "install", pkg, "-q"],
                       capture_output=True, text=True)
    if r.returncode != 0:
        subprocess.run([sys.executable, "-m", "pip", "install", pkg])


for pkg in ["playwright", "openpyxl", "requests"]:
    try:
        __import__(pkg)
        p(f"  [OK] {pkg}")
    except ImportError:
        p(f"  Installing {pkg}...")
        pip_install(pkg)
        p(f"  [OK] {pkg} installed.")

p("  Checking browser...")
subprocess.run([sys.executable, "-m", "playwright", "install", "chromium", "--with-deps"],
               capture_output=True, text=True)
p("  [OK] Browser ready.")
p()

# ─────────────────────────────────────────────────────────────────
#  Imports
# ─────────────────────────────────────────────────────────────────
try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    import re, time
    from datetime import datetime
except ImportError as e:
    p(f"  ERROR: {e}")
    try:
        input("  Press ENTER to close...")
    except EOFError:
        pass
    sys.exit(1)

# ─────────────────────────────────────────────────────────────────
#  Desktop path (OneDrive-aware)
# ─────────────────────────────────────────────────────────────────
def find_desktop():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        d, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        if os.path.isdir(d):
            return d
    except Exception:
        pass
    home = os.path.expanduser("~")
    for candidate in [
        os.path.join(home, "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "OneDrive - Personal", "Desktop"),
    ]:
        if os.path.isdir(candidate):
            return candidate
    fb = os.path.join(home, "Desktop")
    os.makedirs(fb, exist_ok=True)
    return fb


OUTPUT_DIR    = find_desktop()
FRESH_ST_FILE = os.path.join(OUTPUT_DIR, "Fresh_St_Market_Products.xlsx")
SAVE_ON_FILE  = os.path.join(OUTPUT_DIR, "Save_On_Foods_Products.xlsx")

# ─────────────────────────────────────────────────────────────────
#  Site configs  (both run on the same Wynshop platform)
# ─────────────────────────────────────────────────────────────────
SITES = {
    "Save On Foods": {
        "base":  "https://www.saveonfoods.com/sm/planning/rsid/1982",
        "host":  "saveonfoods.com",
        "file":  SAVE_ON_FILE,
    },
    "Fresh St. Market": {
        "base":  "https://www.freshstmarket.com/sm/pickup/rsid/055",
        "host":  "freshstmarket.com",
        "file":  FRESH_ST_FILE,
    },
}

CDP_PORT = 9222

# Auto-detect Chrome or Edge — tries common install locations
def _find_browser():
    candidates = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.join(os.path.expanduser("~"),
                     r"AppData\Local\Google\Chrome\Application\chrome.exe"),
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None

CHROME_EXE = _find_browser()

# ─────────────────────────────────────────────────────────────────
#  DOM selectors  (same on both sites)
# ─────────────────────────────────────────────────────────────────
CARD_SEL  = "[data-testid^='ProductCardWrapper']"
PRICE_SEL = "[class*='ProductPrice--']"
WAS_SEL   = "[class*='ProductWasPrice--']"

# ─────────────────────────────────────────────────────────────────
#  Browser helpers
# ─────────────────────────────────────────────────────────────────
def launch_chrome_cdp():
    """
    Launch a fresh Chrome instance with --remote-debugging-port.
    Uses a temp profile so it doesn't conflict with any existing Chrome window.
    Returns (proc, temp_dir) — caller should delete temp_dir when done.
    """
    import subprocess as sp
    import tempfile
    temp_dir = tempfile.mkdtemp(prefix="scraper_chrome_")
    proc = sp.Popen([
        CHROME_EXE,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={temp_dir}",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-popup-blocking",
        "--window-size=1280,900",
        "about:blank",
    ])
    time.sleep(5)   # give Chrome time to start and open the debug port
    return proc, temp_dir


def make_browser(playwright, use_real_chrome=False):
    """
    use_real_chrome=True  → attach to Chrome via CDP on port 9222 (real fingerprint,
                            bypasses Cloudflare). Call launch_chrome_cdp() first.
    use_real_chrome=False → headless Chromium (fast, works for Fresh St.).
    Returns (browser, context).
    """
    if use_real_chrome:
        # Retry a few times — Chrome needs a moment to start its debug port
        last_err = None
        for attempt in range(6):
            try:
                browser = playwright.chromium.connect_over_cdp(
                    f"http://127.0.0.1:{CDP_PORT}", timeout=8000)
                ctx = (browser.contexts[0]
                       if browser.contexts
                       else browser.new_context(viewport={"width": 1280, "height": 900}))
                return browser, ctx
            except Exception as e:
                last_err = e
                time.sleep(2)
        raise RuntimeError(f"Could not connect to Chrome on port {CDP_PORT}: {last_err}")

    browser = playwright.chromium.launch(headless=True)
    ctx = browser.new_context(
        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"),
        viewport={"width": 1280, "height": 900},
        locale="en-CA",
    )
    ctx.route("**/*.{mp4,mp3,avi,wmv,woff2,woff}", lambda r: r.abort())
    return browser, ctx


def is_blocked(page):
    c = page.content().lower()
    # Only flag as blocked if it's actually a Cloudflare challenge page,
    # not just any page that references Cloudflare as a CDN
    return ("just a moment" in c and "checking" in c) or \
           "cf-browser-verification" in c or \
           "__cf_chl_" in c or \
           ("ray id" in c and "enable javascript" in c)


def safe_goto(page, url, retries=2):
    for attempt in range(retries + 1):
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=40000)
            time.sleep(2.5)
            # Wait up to 30s if Cloudflare challenge appears
            if is_blocked(page):
                p("    [cf] Cloudflare challenge — waiting for it to pass...")
                for _ in range(12):
                    time.sleep(5)
                    if not is_blocked(page):
                        p("    [cf] Passed.")
                        break
            return True
        except Exception as e:
            if attempt == retries:
                p(f"    [!] Failed after {retries+1} attempts: {e}")
                return False
            time.sleep(3)
    return False


# ─────────────────────────────────────────────────────────────────
#  Category discovery
# ─────────────────────────────────────────────────────────────────
def discover_categories(page, base_url, host):
    """
    Load /departments and return list of (slug, name) for every LEAF
    subcategory (those that have a parent/child slug with a '/').
    Leaf categories are the ones that contain actual products.
    """
    p(f"  Discovering categories from /departments...")
    if not safe_goto(page, f"{base_url}/departments"):
        return []

    js = f"""
    (() => {{
        const prefix = 'https://www.{host}';
        const links = Array.from(document.querySelectorAll('a[href*="/categories/"]'));
        const seen = new Set();
        const results = [];
        for (const a of links) {{
            const slug = a.href
                .replace(prefix, '')
                .replace(/\\/sm\\/[^/]+\\/rsid\\/\\d+/, '')
                .replace('/categories/', '')
                .trim();
            if (!slug || seen.has(slug) || slug.startsWith('category/')) continue;
            if (!slug.includes('-id-')) continue;
            seen.add(slug);
            results.push({{ slug, name: a.innerText.trim() }});
        }}
        return results;
    }})()
    """
    try:
        all_cats = page.evaluate(js)
    except Exception as e:
        p(f"  [!] Category discovery error: {e}")
        return []

    # Only leaf categories (slug contains '/' = has a parent/child path)
    leafs = [c for c in all_cats if '/' in c['slug']]
    p(f"  Found {len(leafs)} leaf categories  ({len(all_cats)} total incl. parents)")
    return [(c['slug'], c['name']) for c in leafs]


# ─────────────────────────────────────────────────────────────────
#  Product extraction from one page
# ─────────────────────────────────────────────────────────────────
def extract_products(page, category=""):
    cards = page.query_selector_all(CARD_SEL)
    results = []
    for card in cards:
        try:
            raw_id = card.get_attribute("data-testid") or ""
            pnum   = raw_id.replace("ProductCardWrapper-", "").lstrip("0") or raw_id

            # Second <p> is the clean product name; first has price appended in aria form
            ps   = card.query_selector_all("p")
            name = ""
            if len(ps) >= 2:
                name = ps[1].inner_text().strip()
            elif ps:
                name = re.sub(r",\s*\$[\d.]+.*$", "", ps[0].inner_text()).strip()

            name = name.replace("Open Product Description", "").strip()

            pe      = card.query_selector(PRICE_SEL)
            we      = card.query_selector(WAS_SEL)
            current = pe.inner_text().strip() if pe else ""
            was_raw = we.inner_text().strip() if we else ""
            was     = re.sub(r"^was\s*", "", was_raw, flags=re.IGNORECASE).strip()

            if not name and not pnum:
                continue

            results.append({
                "Category":       category,
                "Product Name":   name,
                "Product Number": pnum,
                "Regular Price":  was if was else current,
                "Promo Price":    current if was else "",
            })
        except Exception:
            continue
    return results


# ─────────────────────────────────────────────────────────────────
#  Paginate through all pages of a URL
# ─────────────────────────────────────────────────────────────────
def scrape_all_pages(page, base_url, category, page_size=30):
    """
    Scrape page 1, read total count, then paginate with ?page=N&skip=M.
    Returns list of product dicts.
    """
    if not safe_goto(page, base_url):
        return []

    if is_blocked(page):
        p(f"    [!] Blocked.")
        return []

    # Read total from pagination widget
    total_items = 0
    try:
        pag_el = page.query_selector('[class*="Pagination"]')
        if pag_el:
            m = re.search(r"of\s+([\d,]+)", pag_el.inner_text())
            if m:
                total_items = int(m.group(1).replace(",", ""))
    except Exception:
        pass

    prods_p1 = extract_products(page, category)
    if not prods_p1 and total_items == 0:
        return []

    if total_items == 0:
        total_items = len(prods_p1)

    total_pages = max(1, -(-total_items // page_size))
    all_prods   = {pr["Product Number"]: pr for pr in prods_p1 if pr["Product Number"]}

    for pg in range(2, total_pages + 1):
        url = f"{base_url}?page={pg}&skip={(pg-1)*page_size}"
        if not safe_goto(page, url):
            break
        if is_blocked(page):
            p(f"    [!] Blocked at page {pg}.")
            break
        prods = extract_products(page, category)
        if not prods:
            break
        for pr in prods:
            pn = pr["Product Number"]
            if pn and pn not in all_prods:
                all_prods[pn] = pr

    return list(all_prods.values())


# ─────────────────────────────────────────────────────────────────
#  Sitemap gap-fill  (catches any products not in any category page)
# ─────────────────────────────────────────────────────────────────
def scrape_sitemap_gaps(page, base_url, scraped_pnums_set):
    """
    Fetch /sitemap.xml, compare product SKUs with what was scraped,
    navigate to each missing product page and read __PRELOADED_STATE__.product.
    """
    import requests as _req

    host        = base_url.split('/sm/')[0]          # https://www.freshstmarket.com
    sitemap_url = f"{host}/sitemap.xml"

    p(f"  Fetching sitemap: {sitemap_url}")
    try:
        r = _req.get(sitemap_url, timeout=20,
                     headers={"User-Agent": "Mozilla/5.0"})
        text = r.text
    except Exception as e:
        p(f"  [!] Sitemap fetch failed: {e}")
        return []

    product_urls = re.findall(
        r'https://www\.freshstmarket\.com/product/[^\s<>"]+', text)

    missing = []
    for url in product_urls:
        m = re.search(r'-id-(\d+)$', url)
        if not m:
            continue
        raw_sku = m.group(1)
        sku     = raw_sku.lstrip("0") or raw_sku
        if sku not in scraped_pnums_set:
            slug = url.replace(f"{host}/product/", "")
            missing.append((sku, slug))

    p(f"  Sitemap: {len(product_urls)} products total, "
      f"{len(missing)} not yet in scraped set")

    if not missing:
        p("  All sitemap products already captured!")
        return []

    results = []
    for i, (sku, slug) in enumerate(missing, 1):
        url = f"{base_url}/product/{slug}"
        try:
            if not safe_goto(page, url):
                continue
            prod = page.evaluate("window.__PRELOADED_STATE__?.product || {}")
            if not isinstance(prod, dict) or not prod.get("name"):
                continue

            name     = str(prod.get("name", "")).strip()
            price    = str(prod.get("price", "") or "").strip()
            was      = str(prod.get("wasPrice", "") or "").strip()
            is_disc  = bool(prod.get("isDiscounted", False))
            cats     = prod.get("categories") or []
            cat_name = cats[0].get("name", "") if cats else "Uncategorized"

            results.append({
                "Category":       cat_name,
                "Product Name":   name,
                "Product Number": sku,
                "Regular Price":  was if (is_disc and was) else price,
                "Promo Price":    price if (is_disc and was) else "",
            })
        except Exception:
            pass

        if i % 10 == 0 or i == len(missing):
            p(f"  Gap fill: {i}/{len(missing)} done  ({len(results)} found)")

    p(f"  Gap fill complete: {len(results)} additional products added.")
    return results


# ─────────────────────────────────────────────────────────────────
#  Main site scraper
# ─────────────────────────────────────────────────────────────────
def scrape_site(playwright, site_name, cfg, use_real_chrome=False):
    p("=" * 60)
    p(f"  Scraping {site_name.upper()}")
    p("=" * 60)

    base     = cfg["base"]
    host     = cfg["host"]
    all_prods = {}   # pnum -> dict

    browser, ctx = make_browser(playwright, use_real_chrome)
    page = ctx.new_page()

    # ── Step 1: Discover categories ───────────────────────────────
    categories = discover_categories(page, base, host)
    if not categories:
        p("  [!] Could not discover categories. Trying promotions page only.")
        categories = []

    p()
    total_cats = len(categories)

    # ── Step 2: Scrape each leaf category ─────────────────────────
    for i, (slug, cat_name) in enumerate(categories, 1):
        url = f"{base}/categories/{slug}"
        p(f"  [{i}/{total_cats}] {cat_name}", flush=True)

        prods = scrape_all_pages(page, url, cat_name)

        new = 0
        for pr in prods:
            pn = pr["Product Number"]
            if not pn:
                continue
            if pn not in all_prods:
                all_prods[pn] = pr
                new += 1

        p(f"    {len(prods)} items  (+{new} new, running total: {len(all_prods)})")

    # ── Step 3: Overlay promo prices from /promotions ─────────────
    p()
    p("  Overlaying promo prices from /promotions...")
    promo_url = f"{base}/promotions"
    promo_total = 0
    if safe_goto(page, promo_url) and not is_blocked(page):
        try:
            pag_el = page.query_selector('[class*="Pagination"]')
            if pag_el:
                m = re.search(r"of\s+([\d,]+)", pag_el.inner_text())
                if m:
                    promo_total = int(m.group(1).replace(",", ""))
        except Exception:
            pass

        promo_pages = max(1, -(-promo_total // 30)) if promo_total else 1
        p(f"  Promotions: {promo_total:,} items across {promo_pages} pages")

        for pg in range(1, promo_pages + 1):
            url = f"{promo_url}?page={pg}&skip={(pg-1)*30}"
            if pg > 1 and not safe_goto(page, url):
                break
            if pg > 1 and is_blocked(page):
                break
            prods = extract_products(page, "Sale")
            for pr in prods:
                pn = pr["Product Number"]
                if not pn:
                    continue
                if pn in all_prods:
                    # Update promo price on existing product
                    if pr.get("Promo Price"):
                        all_prods[pn]["Promo Price"]   = pr["Promo Price"]
                        all_prods[pn]["Regular Price"]  = pr["Regular Price"]
                else:
                    # Product only found in promotions (not in any category)
                    all_prods[pn] = pr

            if pg % 10 == 0 or pg == promo_pages:
                p(f"  Promotions page {pg}/{promo_pages}  "
                  f"(total products: {len(all_prods):,})")

    try:
        page.close()
    except Exception:
        pass
    p(f"\n  {site_name} finished — {len(all_prods):,} unique products.\n")
    return list(all_prods.values())


# ─────────────────────────────────────────────────────────────────
#  Excel builder
# ─────────────────────────────────────────────────────────────────
COLS       = ["Category", "Product Name", "Product Number", "Regular Price", "Promo Price"]
COL_WIDTHS = [28, 50, 18, 16, 16]

CAT_COLOURS = [
    "D9E1F2", "E2EFDA", "FCE4D6", "FFF2CC", "EDEDED",
    "D6E4F7", "F2D7EE", "D7F2EE", "F7ECD6", "E8D7F2",
]


def build_excel(products, filepath, store_name):
    products = sorted(products, key=lambda x: (x.get("Category", ""),
                                               x.get("Product Name", "")))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = store_name[:31]

    def bdr(color="BDD7EE"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    last_col = get_column_letter(len(COLS))

    ws.merge_cells(f"A1:{last_col}1")
    c = ws["A1"]
    c.value     = f"{store_name}  -  Product Price List"
    c.font      = Font(bold=True, size=14, color="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill      = PatternFill("solid", fgColor="D6E4F7")
    ws.row_dimensions[1].height = 30

    ws.merge_cells(f"A2:{last_col}2")
    c = ws["A2"]
    c.value     = (f"Scraped on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}   "
                   f"|   Total products: {len(products):,}")
    c.font      = Font(italic=True, size=9, color="555555")
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14

    for i, col in enumerate(COLS, 1):
        c = ws.cell(row=3, column=i, value=col)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
    ws.row_dimensions[3].height = 22

    cat_colour_idx = {}
    colour_counter = 0
    prev_cat = None

    for row_i, pr in enumerate(products, 4):
        cat = pr.get("Category", "")
        if cat not in cat_colour_idx:
            cat_colour_idx[cat] = colour_counter % len(CAT_COLOURS)
            colour_counter += 1

        is_cat_start = (cat != prev_cat)
        prev_cat = cat
        fill = PatternFill("solid", fgColor=CAT_COLOURS[cat_colour_idx[cat]])

        values = [cat, pr.get("Product Name",""), pr.get("Product Number",""),
                  pr.get("Regular Price",""), pr.get("Promo Price","")]
        for col_i, val in enumerate(values, 1):
            c = ws.cell(row=row_i, column=col_i, value=val)
            c.border    = bdr("C9C9C9")
            c.fill      = fill
            c.alignment = Alignment(vertical="center")

        ws.cell(row=row_i, column=1).font      = Font(color="444444", size=9, italic=True)
        ws.cell(row=row_i, column=2).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=row_i, column=3).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_i, column=3).font      = Font(name="Courier New", size=9)
        ws.cell(row=row_i, column=4).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=row_i, column=5).alignment = Alignment(horizontal="right", vertical="center")
        if pr.get("Promo Price"):
            ws.cell(row=row_i, column=5).font = Font(bold=True, color="C00000")

        if is_cat_start and row_i > 4:
            for col_i in range(1, len(COLS) + 1):
                old = ws.cell(row=row_i, column=col_i).border
                ws.cell(row=row_i, column=col_i).border = Border(
                    left=old.left, right=old.right, bottom=old.bottom,
                    top=Side(style="medium", color="888888"))

    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:{last_col}{3 + len(products)}"

    wb.save(filepath)
    p(f"  [SAVED] {filepath}")


# ─────────────────────────────────────────────────────────────────
#  Main
# ─────────────────────────────────────────────────────────────────
def main():
    start = time.time()
    p(f"  Output folder : {OUTPUT_DIR}")
    p(f"  Started at    : {datetime.now().strftime('%I:%M %p')}")
    p()

    results = {}

    try:
        with sync_playwright() as pw:
            for site_name, cfg in SITES.items():
                # Save On Foods needs the real Chrome browser to bypass Cloudflare.
                # Fresh St. runs headlessly (no Chrome needed).
                use_real_chrome = (site_name == "Save On Foods")

                chrome_proc = None
                chrome_tmp  = None
                if use_real_chrome:
                    if not CHROME_EXE:
                        p("─" * 60)
                        p("  ERROR: Chrome or Edge not found on this computer.")
                        p("  Save On Foods requires Chrome or Edge to bypass")
                        p("  Cloudflare protection.")
                        p("  Install Chrome from: https://www.google.com/chrome")
                        p("  Then run the scraper again.")
                        p("─" * 60)
                        results[site_name] = []
                        continue
                    p("─" * 60)
                    p("  SAVE ON FOODS needs your real Chrome/Edge browser.")
                    p("  Chrome will be closed and reopened automatically.")
                    p("  If a security check appears, click it and wait —")
                    p("  the script will continue on its own after ~5 seconds.")
                    p("─" * 60)
                    p()
                    try:
                        input("  Press ENTER to open Chrome and start Save On Foods... ")
                    except (EOFError, KeyboardInterrupt):
                        p("  Skipping Save On Foods.")
                        results[site_name] = []
                        continue
                    chrome_proc, chrome_tmp = launch_chrome_cdp()
                    p("  Chrome launched with remote debugging. Connecting...")

                prods = scrape_site(pw, site_name, cfg, use_real_chrome=use_real_chrome)
                results[site_name] = prods

                # Close the Chrome process we launched for Save On Foods
                if chrome_proc:
                    try:
                        chrome_proc.terminate()
                    except Exception:
                        pass
                    time.sleep(1)
                    # Clean up the temp profile directory
                    try:
                        import shutil
                        shutil.rmtree(chrome_tmp, ignore_errors=True)
                    except Exception:
                        pass

                if prods:
                    p(f"  Building {site_name} Excel file...")
                    build_excel(prods, cfg["file"], site_name)
                    p()
                else:
                    p(f"  [!] No products collected for {site_name}.")
                    p()

    except KeyboardInterrupt:
        p("\n  Stopped by user.")
    except Exception as e:
        p(f"\n  UNEXPECTED ERROR: {e}")
        import traceback
        traceback.print_exc()
        p()
        p("  Please send the above error to whoever set this up.")

    mins, secs = divmod(int(time.time() - start), 60)
    p()
    p("=" * 60)
    p("  ALL DONE!")
    p("=" * 60)
    for site_name, prods in results.items():
        if prods:
            p(f"  {site_name:<20}: {len(prods):>6,} products  ->  Excel saved")
    p(f"  Time taken           : {mins}m {secs}s")
    p(f"  Files saved to       : {OUTPUT_DIR}")
    p()

    try:
        input("  Press ENTER to close this window...")
    except EOFError:
        pass


if __name__ == "__main__":
    main()

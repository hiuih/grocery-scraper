#!/usr/bin/env python3
"""
GROCERY PRICE SCRAPER
======================
Scrapes every product from:
  - Fresh St. Market  ->  Fresh_St_Market_Products.xlsx
  - Save On Foods     ->  Save_On_Foods_Products.xlsx

HOW TO RUN:
  Double-click this file, or run:  python grocery_scraper.py
  Results appear on your Desktop when finished.
"""

# ── Fix console encoding (avoids blank-window bug) ────────────────
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


for pkg, iname in [("playwright", "playwright"), ("openpyxl", "openpyxl")]:
    try:
        __import__(iname)
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
    import re, time, json
    from datetime import datetime
    from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
except ImportError as e:
    p(f"  ERROR: {e}")
    p("  Run:  pip install playwright openpyxl && python -m playwright install chromium")
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
    for p_ in [os.path.join(home, "Desktop"),
               os.path.join(home, "OneDrive", "Desktop"),
               os.path.join(home, "OneDrive - Personal", "Desktop"),
               os.path.join(home, "OneDrive - Commercial", "Desktop")]:
        if os.path.isdir(p_):
            return p_
    fb = os.path.join(home, "Desktop")
    os.makedirs(fb, exist_ok=True)
    return fb


OUTPUT_DIR    = find_desktop()
FRESH_ST_FILE = os.path.join(OUTPUT_DIR, "Fresh_St_Market_Products.xlsx")
SAVE_ON_FILE  = os.path.join(OUTPUT_DIR, "Save_On_Foods_Products.xlsx")

# ─────────────────────────────────────────────────────────────────
#  Fresh St. Market — top-level categories only.
#  Subcategories are discovered dynamically at runtime by visiting
#  each parent category page and following "View all" tile links.
#  This avoids hardcoded IDs that go stale when the store updates.
# ─────────────────────────────────────────────────────────────────
FRESH_ST_BASE = "https://www.freshstmarket.com/sm/pickup/rsid/055"

FRESH_ST_TOP_CATS = [
    ("deli-prepared-foods-id-3568",    "Deli & Prepared Foods"),
    ("baby-id-3571",                   "Baby"),
    ("personal-care-items-id-3605",    "Personal Care"),
    ("dairy-eggs-cheese-id-3622",      "Dairy, Eggs & Cheese"),
    ("condiments-sauces-id-3657",      "Condiments & Sauces"),
    ("crackers-cookies-id-3659",       "Crackers & Cookies"),
    ("grains-pasta-sides-id-3664",     "Grains, Pasta & Sides"),
    ("international-id-3671",          "International"),
    ("bakery-id-3673",                 "Bakery"),
    ("bulk-id-3675",                   "Bulk"),
    ("beverages-id-3691",              "Beverages"),
    ("frozen-food-id-3708",            "Frozen Food"),
    ("household-goods-id-3717",        "Household Goods"),
    ("breakfast-cereals-id-3752",      "Breakfast & Cereals"),
    ("canned-goods-id-3756",           "Canned Goods"),
    ("snacks-candy-id-3777",           "Snacks & Candy"),
    ("pet-care-id-3788",               "Pet Care"),
    ("wellness-products-id-3795",      "Wellness"),
    ("meat-id-3796",                   "Meat"),
    ("produce-id-3803",                "Produce"),
    ("seafood-id-3804",                "Seafood"),
    ("spices-baking-id-3806",          "Spices & Baking"),
]

# Regex to identify category URLs (ends with -id-NNNN)
_CAT_ID_RE = re.compile(r'/categories/[^?#\s]+-id-\d+', re.IGNORECASE)

# ─────────────────────────────────────────────────────────────────
#  Browser setup
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
        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"),
        viewport={"width": 1280, "height": 900},
        locale="en-CA",
    )
    ctx.route("**/*.{mp4,mp3,avi,wmv}", lambda r: r.abort())
    return browser, ctx


def make_stealth_browser(playwright):
    """
    Launch a visible (non-headless) browser with anti-bot measures.
    Used for Cloudflare-protected sites like Save On Foods.
    """
    launch_kwargs = dict(
        headless=False,
        args=["--start-maximized", "--disable-blink-features=AutomationControlled"],
    )
    try:
        browser = playwright.chromium.launch(channel="chrome", **launch_kwargs)
    except Exception:
        browser = playwright.chromium.launch(**launch_kwargs)

    ctx = browser.new_context(
        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"),
        viewport={"width": 1366, "height": 768},
        locale="en-CA",
        timezone_id="America/Vancouver",
    )
    # Remove automation fingerprints
    ctx.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
        Object.defineProperty(navigator, 'plugins',   { get: () => [
            { name: 'Chrome PDF Plugin' },
            { name: 'Chrome PDF Viewer' },
            { name: 'Native Client' }
        ]});
        window.chrome = { runtime: {} };
    """)
    ctx.route("**/*.{mp4,mp3,avi,wmv}", lambda r: r.abort())
    return browser, ctx


# ─────────────────────────────────────────────────────────────────
#  Wynshop direct API helpers
#  Discovered endpoint: storefrontgateway.freshstmarket.com
#    /api/stores/055/categories/{id}/products  → product list
#    /api/stores/055/categories/{id}/groupby   → subcategory groups
# ─────────────────────────────────────────────────────────────────
WYNSHOP_STORE_ID = "055"
WYNSHOP_GW       = "https://storefrontgateway.freshstmarket.com"

def _extract_cat_id(url):
    m = re.search(r'-id-(\d+)', url)
    return m.group(1) if m else None


def _pw_get(page, url):
    """
    HTTP GET via Playwright's own request API — no CORS restrictions,
    shares browser cookies automatically.
    Returns parsed JSON dict/list, or {"_error": ...} on failure.
    """
    try:
        resp = page.request.get(url, headers={"Accept": "application/json"})
        if not resp.ok:
            return {"_error": resp.status}
        return resp.json()
    except Exception as e:
        return {"_error": str(e)}


def _wynshop_api_products(page, cat_id, cat_name):
    """
    Fetch all products for a category from the Wynshop storefrontgateway API.
    Uses Playwright request (no CORS). Returns parsed list or [].
    """
    # Only try the /groupby-style products path — /api/2.0/products returns 404
    endpoints = [
        f"{WYNSHOP_GW}/api/stores/{WYNSHOP_STORE_ID}/categories/{cat_id}/products",
    ]
    for ep in endpoints:
        data = _pw_get(page, ep)
        if data.get("_error"):
            p(f"      [API-ERR] {ep.split('?')[0][-50:]} → {data['_error']}")
            continue
        for key, val in data.items():
            if _looks_like_products(val):
                p(f"      [API-HIT] {ep.split('?')[0][-50:]} key={key} n={len(val)}")
                return fetch_all_via_api(page, ep, data, key, cat_name)
        for key, val in data.items():
            if isinstance(val, dict):
                for ik, iv in val.items():
                    if _looks_like_products(iv):
                        p(f"      [API-HIT-NESTED] key={key}.{ik}")
                        return fetch_all_via_api(page, ep, val, ik, cat_name)
        p(f"      [API-KEYS] {ep.split('?')[0][-50:]} → {list(data.keys())[:6]}")
    return []


def _wynshop_groupby_subcats(page, cat_id):
    """
    Call /groupby (no pagination params — the API doesn't support them)
    to get all subcategory groups the DOM may not show.
    Returns [(sub_id_str, sub_name), ...].
    """
    url  = f"{WYNSHOP_GW}/api/stores/{WYNSHOP_STORE_ID}/categories/{cat_id}/groupby"
    data = _pw_get(page, url)
    if data.get("_error"):
        p(f"      [GROUPBY-ERR] {data['_error']}")
        return []
    items = data.get("items") or []
    total = data.get("total") or len(items)
    if not items:
        p(f"      [GROUPBY] empty — keys: {list(data.keys())[:6]}")
        return []
    first = items[0]
    p(f"      [GROUPBY] {total} groups, item keys: {list(first.keys())[:10]}")
    # Print full first item so we can see its structure
    p(f"      [GROUPBY-ITEM1] { {k: str(v)[:60] for k, v in list(first.items())[:8]} }")
    result = []
    for item in items:
        sid = (item.get("id") or item.get("categoryId") or
               item.get("departmentId") or item.get("groupId") or
               item.get("nodeId"))
        if not sid:
            continue
        name = item.get("name") or item.get("title") or str(sid)
        result.append((str(sid), name))
    return result


# ─────────────────────────────────────────────────────────────────
#  Fresh St. — dynamic subcategory discovery helpers
# ─────────────────────────────────────────────────────────────────
def _find_subcat_tiles(page):
    """
    Find subcategory links on a Fresh St. category page.
    Primary pass  : "View all" tile links (the standard pattern).
    Fallback pass : any category-ID link outside nav/header/footer chrome,
                    catching sub-subcategory pages that list categories
                    without "View all" tiles.
    Returns a list of (full_url, label) tuples.
    """
    results = []
    seen    = set()
    base_parsed = urlparse(FRESH_ST_BASE)

    def _resolve(href):
        if href.startswith("/"):
            return (f"{base_parsed.scheme}://{base_parsed.netloc}"
                    + href.split("?")[0])
        if href.startswith("http"):
            return href.split("?")[0]
        return None

    def _label_from_link(link, href):
        try:
            label = link.evaluate("""
                el => {
                    let node = el.parentElement;
                    for (let i = 0; i < 6; i++) {
                        if (!node) break;
                        const h = node.querySelector('h2,h3,h4,h5,h6');
                        if (h && h !== el) return h.innerText.trim();
                        node = node.parentElement;
                    }
                    return null;
                }
            """)
        except Exception:
            label = None
        if not label:
            slug  = href.rstrip("/").split("/")[-1].split("?")[0]
            label = slug.split("-id-")[0].replace("-", " ").title()
        else:
            label = re.sub(r"\s*\(\d+\)\s*$", "", label).strip()
        return label

    # ── Primary pass: "View all" tile links ──────────────────────
    for link in page.query_selector_all("a[href]"):
        try:
            text = link.inner_text().strip()
            href = link.get_attribute("href") or ""
        except Exception:
            continue

        if "view all" not in text.lower():
            continue
        if not _CAT_ID_RE.search(href):
            continue

        full_url = _resolve(href)
        if not full_url or full_url in seen:
            continue
        seen.add(full_url)
        results.append((full_url, _label_from_link(link, href)))

    if results:
        return results

    # ── Fallback pass: any category-ID link outside page chrome ──
    # Triggered when sub-subcategory pages don't use "View all" tiles.
    for link in page.query_selector_all("a[href]"):
        try:
            href = link.get_attribute("href") or ""
            text = (link.inner_text() or "").strip()
        except Exception:
            continue

        if not _CAT_ID_RE.search(href):
            continue

        try:
            in_chrome = link.evaluate("""
                el => {
                    let n = el;
                    while (n) {
                        const tag = (n.tagName || '').toLowerCase();
                        const cls = (n.className || '').toString().toLowerCase();
                        if (['nav','header','footer'].includes(tag)) return true;
                        if (cls.includes('breadcrumb') || cls.includes('sidebar') ||
                            cls.includes('topbar')     || cls.includes('top-bar') ||
                            cls.includes('globalheader') ||
                            cls.includes('global-header')) return true;
                        n = n.parentElement;
                    }
                    return false;
                }
            """)
            if in_chrome:
                continue
        except Exception:
            pass

        full_url = _resolve(href)
        if not full_url or full_url in seen:
            continue
        seen.add(full_url)

        slug  = href.rstrip("/").split("/")[-1].split("?")[0]
        label = text or slug.split("-id-")[0].replace("-", " ").title()
        label = re.sub(r"\s*\(\d+\)\s*$", "", label).strip()
        results.append((full_url, label))

    return results


def _scrape_category(page, url, cat_name, all_products, visited):
    """
    Visit a Fresh St. category URL.
      • If the page has product cards (or an intercepted API response):
        extract products and add to all_products.
      • Otherwise: find subcategory tile links and recurse into each.
    Uses a 'visited' set so no URL is fetched twice.
    """
    if url in visited:
        return
    visited.add(url)

    p(f"\n  → {cat_name}")

    # ── Set up API interception ────────────────────────────────
    intercepted = []

    def _on_resp(resp, _buf=intercepted):
        if _buf:
            return
        try:
            if resp.status != 200:
                return
            ct = resp.headers.get("content-type", "")
            if "json" not in ct:
                return
            data = resp.json()
            if _looks_like_products(data):
                _buf.append((resp.url, {"products": data}, "products"))
                return
            if not isinstance(data, dict):
                return
            # Log every JSON API response so we can see what the page calls
            if any(x in resp.url for x in ("api", "storefront", "product", "catalog", "search")):
                p(f"      [NET] {resp.url.split('?')[0][-70:]} → {list(data.keys())[:5]}")
            for key, val in data.items():
                if _looks_like_products(val):
                    _buf.append((resp.url, data, key))
                    return
            for key, val in data.items():
                if isinstance(val, dict):
                    for inner_key, inner_val in val.items():
                        if _looks_like_products(inner_val):
                            _buf.append((resp.url, val, inner_key))
                            return
        except Exception:
            pass

    page.on("response", _on_resp)
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=35000)
        time.sleep(2.5)
    except Exception as e:
        p(f"    [!] Load error: {e}")
        page.remove_listener("response", _on_resp)
        return

    # Scroll to trigger any lazy-load API calls
    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(1.5)
    page.evaluate("window.scrollTo(0, 0)")
    time.sleep(0.5)
    page.remove_listener("response", _on_resp)

    # ── Extract products if this page has any ────────────────
    has_cards = bool(page.query_selector_all(CARD_SEL))

    cat_id = _extract_cat_id(url)

    if has_cards or intercepted:
        if intercepted:
            first_url, first_data, products_key = intercepted[0]
            p("    API intercepted — paginating...")
            cat_prods = fetch_all_via_api(
                page, first_url, first_data, products_key, cat_name)
        else:
            # Try Wynshop products API directly before falling back to DOM
            cat_prods = (_wynshop_api_products(page, cat_id, cat_name)
                         if cat_id else [])
            if not cat_prods:
                p("    No API hit — using DOM scroll fallback")
                cat_prods = extract_dom_incremental(page, cat_name)

        cat_new = 0
        for pr in cat_prods:
            pn = pr["Product Number"]
            if not pn:
                continue
            if pn not in all_products:
                all_products[pn] = pr
                cat_new += 1
            elif pr.get("Promo Price") and not all_products[pn].get("Promo Price"):
                all_products[pn]["Promo Price"]  = pr["Promo Price"]
                all_products[pn]["Regular Price"] = pr["Regular Price"]

        p(f"    +{cat_new} new  |  {len(cat_prods)} fetched  "
          f"|  running total: {len(all_products)}")

    # ── Subcategory discovery: DOM tiles + Wynshop groupby API ────
    subcats = _find_subcat_tiles(page)

    # Also ask the groupby API — it may know subcategories not in the DOM
    if cat_id:
        api_subcats = _wynshop_groupby_subcats(page, cat_id)
        dom_ids = {_extract_cat_id(u) for u, _ in subcats}
        visited_api_ids = set()
        for sid, sname in api_subcats:
            if sid in dom_ids or sid in visited_api_ids:
                continue
            visited_api_ids.add(sid)
            sub_label = f"{cat_name} / {sname}"
            # Fake a "visited" URL so we never revisit this ID via DOM path
            fake_url = f"{FRESH_ST_BASE}/categories/api-id-{sid}"
            if fake_url in visited:
                continue
            visited.add(fake_url)
            api_prods = _wynshop_api_products(page, sid, sub_label)
            if api_prods:
                new_count = 0
                for pr in api_prods:
                    pn = pr["Product Number"]
                    if pn and pn not in all_products:
                        all_products[pn] = pr
                        new_count += 1
                p(f"    [API-sub] {sname}: +{new_count} new ({len(api_prods)} fetched) "
                  f"| total: {len(all_products)}")

    if subcats:
        p(f"    {len(subcats)} DOM subcategories — drilling in...")
        for subcat_url, subcat_label in subcats:
            _scrape_category(
                page, subcat_url,
                f"{cat_name} / {subcat_label}",
                all_products, visited,
            )
    elif not has_cards and not intercepted:
        p(f"    [!] No products or subcategory tiles — skipping: {url}")


# ─────────────────────────────────────────────────────────────────
#  API interception helpers
# ─────────────────────────────────────────────────────────────────
def _looks_like_products(lst):
    if not isinstance(lst, list) or not lst:
        return False
    s = lst[0]
    if not isinstance(s, dict):
        return False
    keys_lc = {k.lower() for k in s.keys()}
    return bool(keys_lc & {
        "sku", "upc", "name", "price", "productid",
        "displayname", "itemid", "wasprize", "waspricestring",
        "title", "productname", "productsku", "barcode",
        "regularprice", "currentprice", "listprice",
    })


def _fmt_price(v):
    if v is None:
        return ""
    s = str(v).strip().replace(",", "")
    if not s or s.lower() in ("0", "0.0", "none", "null", ""):
        return ""
    if s.startswith("$"):
        return s
    try:
        return f"${float(s):.2f}"
    except ValueError:
        return s


def _parse_api_item(item, category=""):
    """
    Parse one product dict from the Wynshop/storefrontgateway API.
    Field names confirmed from live HTML inspection:
      sku, name, price (string, current price), wasPrice (string, original),
      isDiscounted (bool)
    """
    pnum = str(
        item.get("sku") or item.get("upc") or item.get("productId") or
        item.get("id") or item.get("itemId") or ""
    ).strip().lstrip("0")

    name = str(
        item.get("name") or item.get("displayName") or item.get("title") or
        item.get("productName") or item.get("description") or ""
    ).strip()

    if not name and not pnum:
        return None

    # ── Wynshop flat format ──────────────────────────────────────
    # price    = current price (may be sale price)  e.g. "$10.49"
    # wasPrice = original price (only present if on sale) e.g. "$13.99"
    # isDiscounted = True/False
    is_disc   = bool(item.get("isDiscounted", False))
    price_str = str(item.get("price")    or "").strip()
    was_str   = str(item.get("wasPrice") or "").strip()

    if is_disc and was_str:
        regular = _fmt_price(was_str)
        promo   = _fmt_price(price_str)
    elif price_str:
        regular = _fmt_price(price_str)
        promo   = ""
    else:
        # ── Nested price object (fallback for other platforms) ───
        pd = item.get("price") or item.get("prices") or item.get("pricing") or {}
        if isinstance(pd, dict):
            reg_raw  = (pd.get("regular") or pd.get("originalPrice") or
                        pd.get("listPrice") or pd.get("was") or
                        pd.get("regularPrice") or "")
            sale_raw = (pd.get("sale") or pd.get("promo") or
                        pd.get("salePrice") or pd.get("now") or
                        pd.get("promotionalPrice") or "")
            if not reg_raw:
                reg_raw = pd.get("current") or pd.get("value") or pd.get("amount") or ""
            regular = _fmt_price(reg_raw)
            promo   = _fmt_price(sale_raw)
        elif isinstance(pd, (int, float)):
            regular = _fmt_price(pd)
            promo   = ""
        else:
            regular = _fmt_price(
                item.get("regularPrice") or item.get("listPrice") or
                item.get("basePrice") or item.get("normalPrice") or "")
            promo   = _fmt_price(
                item.get("salePrice") or item.get("promoPrice") or
                item.get("discountedPrice") or "")

    return {
        "Category":       category,
        "Product Name":   name,
        "Product Number": pnum,
        "Regular Price":  regular,
        "Promo Price":    promo,
    }


def page_fetch(page, url):
    """
    Fetch a URL using the browser's own fetch() — includes session cookies
    and auth headers automatically. Much more reliable than page.request.get().
    """
    result = page.evaluate("""
    async (url) => {
        try {
            const r = await fetch(url, {
                credentials: 'include',
                headers: { 'Accept': 'application/json, */*' }
            });
            if (!r.ok) return { _error: r.status };
            return await r.json();
        } catch(e) {
            return { _error: String(e) };
        }
    }
    """, url)
    return result if result else {}


def fetch_all_via_api(page, first_url, first_data, products_key, category):
    """
    Paginate the storefront API completely.
    Uses page.evaluate(fetch()) so cookies/auth are handled automatically.
    Tries take=9999 first; falls back to skip-based pagination.
    """
    raw       = list(first_data[products_key])
    page_size = max(len(raw), 1)

    # Total count — handle int or string
    total = None
    for tk in ("totalCount", "total", "totalItems", "count", "numFound",
               "totalResults", "itemCount", "recordCount"):
        v = first_data.get(tk)
        if v is not None:
            try:
                total = int(v)
                break
            except (ValueError, TypeError):
                pass

    p(f"    API: {page_size} items, totalCount={total}")

    if total and total <= page_size:
        return [r for r in (_parse_api_item(i, category) for i in raw) if r]

    # Parse URL params
    parsed = urlparse(first_url)
    params = {k: (v[0] if len(v) == 1 else v)
              for k, v in parse_qs(parsed.query, keep_blank_values=True).items()}

    skip_key = next((s for s in ("skip", "Skip", "offset", "Offset", "start", "from")
                     if s in params), "skip")
    take_key = next((s for s in ("take", "Take", "limit", "Limit", "pageSize", "count")
                     if s in params), "take")

    # ── Try fetching everything at once (take=9999) ────────────
    params[skip_key] = "0"
    params[take_key] = "9999"
    big_url  = urlunparse(parsed._replace(query=urlencode(params, doseq=True)))
    big_data = page_fetch(page, big_url)

    if (not big_data.get("_error") and
            isinstance(big_data.get(products_key), list) and
            len(big_data[products_key]) > page_size):
        big_list = big_data[products_key]
        p(f"    Got all {len(big_list)} products in one request.")
        return [r for r in (_parse_api_item(i, category) for i in big_list) if r]

    # ── Fall back to skip-based pagination ─────────────────────
    params[take_key] = str(page_size)
    if not total:
        total = 99999   # keep going until we get empty responses

    skip     = page_size
    page_num = 2
    no_new   = 0

    while skip < total and no_new < 3:
        params[skip_key] = str(skip)
        next_url = urlunparse(parsed._replace(query=urlencode(params, doseq=True)))
        data = page_fetch(page, next_url)

        if data.get("_error"):
            p(f"    API error at skip={skip}: {data['_error']}")
            no_new += 1
            skip += page_size
            page_num += 1
            continue

        items = data.get(products_key, [])
        if not items:
            no_new += 1
        else:
            no_new = 0
            raw.extend(items)
            skip += len(items)
            p(f"    Page {page_num}: +{len(items)} items  "
              f"(collected: {len(raw)}{f'/{total}' if total < 99999 else ''})")
        page_num += 1

    return [r for r in (_parse_api_item(i, category) for i in raw) if r]


# ─────────────────────────────────────────────────────────────────
#  DOM fallback — incremental scroll for virtual lists
# ─────────────────────────────────────────────────────────────────
def extract_dom_incremental(page, category, max_stall=10):
    seen = {}
    try:
        page.wait_for_selector(CARD_SEL, timeout=15000)
    except Exception:
        return []

    time.sleep(1)
    scroll_pos, step, stall, last_h = 0, 400, 0, 0

    def _grab():
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
                cur     = pe.inner_text().strip() if pe else ""
                was_raw = we.inner_text().strip() if we else ""
                was     = re.sub(r"^was\s*", "", was_raw, flags=re.IGNORECASE).strip()
                seen[pnum] = {
                    "Category":       category,
                    "Product Name":   name,
                    "Product Number": pnum,
                    "Regular Price":  was if was else cur,
                    "Promo Price":    cur if was else "",
                }
            except Exception:
                continue

    while stall < max_stall:
        prev = len(seen)
        _grab()
        stall = 0 if len(seen) > prev else stall + 1
        scroll_pos += step
        page.evaluate(f"window.scrollTo(0, {scroll_pos})")
        time.sleep(0.6)
        cur_h = page.evaluate("document.body.scrollHeight")
        if scroll_pos >= cur_h and cur_h == last_h:
            _grab()
            break
        last_h = cur_h

    return list(seen.values())


# ─────────────────────────────────────────────────────────────────
#  Fresh St. Market scraper  (dynamic subcategory discovery)
# ─────────────────────────────────────────────────────────────────
def scrape_fresh_st(playwright):
    p("=" * 60)
    p("  Scraping FRESH ST. MARKET  (dynamic category discovery)")
    p("=" * 60)
    p("  Subcategories are discovered at runtime — no stale IDs.")
    p()

    browser, ctx = make_browser(playwright)
    page         = ctx.new_page()
    all_products = {}    # pnum -> dict
    visited      = set() # track visited URLs to avoid duplicates

    total = len(FRESH_ST_TOP_CATS)
    for i, (slug, name) in enumerate(FRESH_ST_TOP_CATS, 1):
        url = f"{FRESH_ST_BASE}/categories/{slug}"
        p(f"\n[{i}/{total}] {name}")
        _scrape_category(page, url, name, all_products, visited)

    browser.close()
    p(f"\n  Fresh St. finished — {len(all_products):,} unique products.\n")
    return list(all_products.values())


# ─────────────────────────────────────────────────────────────────
#  Save On Foods scraper
# ─────────────────────────────────────────────────────────────────
SAVE_ON_BASE  = "https://www.saveonfoods.com"
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


def scrape_save_on(playwright):
    p("=" * 60)
    p("  Scraping SAVE ON FOODS")
    p("=" * 60)
    p("  Opening a real (visible) browser to bypass Cloudflare.")
    p("  If a CAPTCHA or challenge appears, solve it and the")
    p("  scraper will continue automatically.")
    p()

    browser, ctx = make_stealth_browser(playwright)
    page = ctx.new_page()

    # ── Land on homepage first to establish a clean session ───
    try:
        page.goto(SAVE_ON_BASE, wait_until="domcontentloaded", timeout=30000)
        # Give Cloudflare up to 8 s to auto-resolve its JS challenge
        time.sleep(8)
    except Exception as e:
        p(f"  [!] Could not connect to Save On Foods: {e}")
        browser.close()
        return []

    content = page.content().lower()
    if "ray id" in content and "cloudflare" in content and "blocked" in content:
        p("  [!] Cloudflare hard-blocked this IP.")
        p("  [!] Try running the scraper from a different network,")
        p("  [!] or wait a few hours and try again.")
        browser.close()
        return []

    p("  Connected to Save On Foods. Starting departments...\n")
    all_products = {}

    for dept_name, dept_path in SAVE_ON_DEPTS:
        p(f"  Department: {dept_name}")
        page_num = 1

        while True:
            url = f"{SAVE_ON_BASE}{dept_path}?Nrpp=48&No={(page_num-1)*48}"
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=30000)
                time.sleep(3)   # extra wait for CF-protected pages
            except Exception as e:
                p(f"    [!] Error loading page: {e}")
                break

            content_lc = page.content().lower()
            if "ray id" in content_lc and "cloudflare" in content_lc:
                p("    [!] Blocked on this department. Skipping.")
                break

            try:
                page.wait_for_selector(
                    "[data-product-id], .product-tile, [class*='ProductCard']",
                    timeout=15000)
            except PWTimeout:
                p(f"    No products found on page {page_num}.")
                break

            cards = (page.query_selector_all("[data-product-id]") or
                     page.query_selector_all(".product-tile") or
                     page.query_selector_all("[class*='ProductCard']") or [])
            if not cards:
                break

            products = []
            for card in cards:
                try:
                    pnum = (card.get_attribute("data-product-id") or
                            card.get_attribute("data-sku") or "").strip()
                    ne   = (card.query_selector("[class*='name']") or
                            card.query_selector("[class*='title']") or
                            card.query_selector("h3") or card.query_selector("h2"))
                    name = ne.inner_text().strip() if ne else ""
                    re_  = (card.query_selector("[class*='regular']") or
                            card.query_selector("[class*='was']") or
                            card.query_selector("[class*='original']"))
                    reg  = re_.inner_text().strip() if re_ else ""
                    pe   = (card.query_selector("[class*='sale']") or
                            card.query_selector("[class*='promo']") or
                            card.query_selector("[class*='now']"))
                    pro  = pe.inner_text().strip() if pe else ""
                    if not reg and not pro:
                        px = card.query_selector("[class*='price']")
                        reg = px.inner_text().strip() if px else ""
                    if name or pnum:
                        products.append({
                            "Category":       dept_name,
                            "Product Name":   name,
                            "Product Number": pnum,
                            "Regular Price":  reg,
                            "Promo Price":    pro,
                        })
                except Exception:
                    continue

            for pr in products:
                key = pr["Product Number"] or pr["Product Name"]
                if key and key not in all_products:
                    all_products[key] = pr

            p(f"    Page {page_num}: {len(products)} items  (total: {len(all_products)})")
            if len(products) < 48:
                break
            page_num += 1

        p(f"    Done: {dept_name}")

    browser.close()
    p(f"\n  Save On Foods finished — {len(all_products):,} unique products.\n")
    return list(all_products.values())


# ─────────────────────────────────────────────────────────────────
#  Excel builder  (5 columns: Category, Name, Number, Regular, Promo)
# ─────────────────────────────────────────────────────────────────
COLS = ["Category", "Product Name", "Product Number", "Regular Price", "Promo Price"]
COL_WIDTHS = [28, 48, 18, 16, 16]

# Category → colour map (cycles through palette for new categories)
CAT_COLOURS = [
    "D9E1F2", "E2EFDA", "FCE4D6", "FFF2CC", "EDEDED",
    "D6E4F7", "F2D7EE", "D7F2EE", "F7ECD6", "E8D7F2",
]


def build_excel(products, filepath, store_name):
    # Sort by category then product name
    products = sorted(products, key=lambda x: (x.get("Category", ""),
                                               x.get("Product Name", "")))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = store_name[:31]

    def bdr(color="BDD7EE"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Title rows ──────────────────────────────────────────────
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

    # ── Header row ──────────────────────────────────────────────
    for i, col in enumerate(COLS, 1):
        c = ws.cell(row=3, column=i, value=col)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = PatternFill("solid", fgColor="1F4E79")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
    ws.row_dimensions[3].height = 22

    # ── Data rows ───────────────────────────────────────────────
    cat_colour_idx = {}
    colour_counter = 0
    prev_cat = None

    for row_i, pr in enumerate(products, 4):
        cat = pr.get("Category", "")
        if cat not in cat_colour_idx:
            cat_colour_idx[cat] = colour_counter % len(CAT_COLOURS)
            colour_counter += 1

        # Slightly darker shade for first row of a new category
        is_cat_start = (cat != prev_cat)
        prev_cat = cat

        fill_hex = CAT_COLOURS[cat_colour_idx[cat]]
        fill     = PatternFill("solid", fgColor=fill_hex)

        values = [
            cat,
            pr.get("Product Name",   ""),
            pr.get("Product Number", ""),
            pr.get("Regular Price",  ""),
            pr.get("Promo Price",    ""),
        ]
        for col_i, val in enumerate(values, 1):
            c = ws.cell(row=row_i, column=col_i, value=val)
            c.border    = bdr("C9C9C9")
            c.fill      = fill
            c.alignment = Alignment(vertical="center")

        # Per-column alignment/font overrides
        ws.cell(row=row_i, column=1).font      = Font(color="444444", size=9, italic=True)
        ws.cell(row=row_i, column=2).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(row=row_i, column=3).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=row_i, column=3).font      = Font(name="Courier New", size=9)
        ws.cell(row=row_i, column=4).alignment = Alignment(horizontal="right", vertical="center")
        ws.cell(row=row_i, column=5).alignment = Alignment(horizontal="right", vertical="center")
        if pr.get("Promo Price"):
            ws.cell(row=row_i, column=5).font = Font(bold=True, color="C00000")

        # Bold top border on category start
        if is_cat_start and row_i > 4:
            for col_i in range(1, len(COLS) + 1):
                old = ws.cell(row=row_i, column=col_i).border
                ws.cell(row=row_i, column=col_i).border = Border(
                    left=old.left, right=old.right, bottom=old.bottom,
                    top=Side(style="medium", color="888888")
                )

    # ── Column widths & freeze ───────────────────────────────────
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

    fresh   = []
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
                p("  [!] Save On Foods — no products (likely Cloudflare blocked).")

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

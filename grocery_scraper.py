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
#  Fresh St. Market — categories
#  Parent categories + all named subcategories discovered from the
#  site nav (236 total).  We scrape every leaf that has products.
# ─────────────────────────────────────────────────────────────────
FRESH_ST_BASE = "https://www.freshstmarket.com/sm/pickup/rsid/055"

FRESH_ST_CATEGORIES = [
    # ── Deli & Prepared Foods ──────────────────────────────────
    ("deli-prepared-foods-id-3568",             "Deli & Prepared Foods"),
    ("deli-prepared-foods/deli-meats-id-3569",  "Deli & Prepared Foods / Deli Meats"),
    ("deli-prepared-foods/deli-cheese-id-3570", "Deli & Prepared Foods / Deli Cheese"),
    # ── Baby ──────────────────────────────────────────────────
    ("baby-id-3571",                            "Baby"),
    ("baby/baby-food-id-3572",                  "Baby / Baby Food"),
    ("baby/baby-care-id-3573",                  "Baby / Baby Care"),
    ("baby/diapers-training-pants-id-3574",     "Baby / Diapers"),
    ("baby/formula-id-3575",                    "Baby / Formula"),
    # ── Personal Care ─────────────────────────────────────────
    ("personal-care-items-id-3605",             "Personal Care"),
    ("personal-care-items/hair-care-id-3606",   "Personal Care / Hair Care"),
    ("personal-care-items/skin-care-id-3607",   "Personal Care / Skin Care"),
    ("personal-care-items/oral-care-id-3608",   "Personal Care / Oral Care"),
    ("personal-care-items/feminine-care-id-3609","Personal Care / Feminine Care"),
    ("personal-care-items/mens-grooming-id-3610","Personal Care / Men's Grooming"),
    # ── Dairy, Eggs & Cheese ──────────────────────────────────
    ("dairy-eggs-cheese-id-3622",               "Dairy, Eggs & Cheese"),
    ("dairy-eggs-cheese/cheese-id-3623",        "Dairy / Cheese"),
    ("dairy-eggs-cheese/butter-margarine-id-3624","Dairy / Butter & Margarine"),
    ("dairy-eggs-cheese/eggs-id-3625",          "Dairy / Eggs"),
    ("dairy-eggs-cheese/yogurt-id-3626",        "Dairy / Yogurt"),
    ("dairy-eggs-cheese/milk-id-3627",          "Dairy / Milk"),
    ("dairy-eggs-cheese/cream-creamers-id-3628","Dairy / Cream & Creamers"),
    ("dairy-eggs-cheese/sour-cream-dips-id-3629","Dairy / Sour Cream & Dips"),
    # ── Condiments & Sauces ───────────────────────────────────
    ("condiments-sauces-id-3657",               "Condiments & Sauces"),
    ("condiments-sauces/ketchup-mustard-relish-id-3630","Condiments / Ketchup & Mustard"),
    ("condiments-sauces/salad-dressing-id-3631","Condiments / Salad Dressing"),
    ("condiments-sauces/cooking-sauces-id-3632","Condiments / Cooking Sauces"),
    ("condiments-sauces/pickles-olives-id-3633","Condiments / Pickles & Olives"),
    ("condiments-sauces/oils-vinegars-id-3634", "Condiments / Oils & Vinegars"),
    ("condiments-sauces/mayo-id-3635",          "Condiments / Mayo"),
    ("condiments-sauces/soy-sauce-asian-id-3636","Condiments / Soy & Asian Sauces"),
    ("condiments-sauces/hot-sauce-id-3637",     "Condiments / Hot Sauce"),
    ("condiments-sauces/bbq-sauce-id-3638",     "Condiments / BBQ Sauce"),
    # ── Crackers & Cookies ────────────────────────────────────
    ("crackers-cookies-id-3659",                "Crackers & Cookies"),
    ("crackers-cookies/crackers-id-3660",       "Crackers & Cookies / Crackers"),
    ("crackers-cookies/cookies-id-3661",        "Crackers & Cookies / Cookies"),
    # ── Grains, Pasta & Sides ─────────────────────────────────
    ("grains-pasta-sides-id-3664",              "Grains, Pasta & Sides"),
    ("grains-pasta-sides/pasta-id-3665",        "Grains / Pasta"),
    ("grains-pasta-sides/rice-id-3666",         "Grains / Rice"),
    ("grains-pasta-sides/bread-crumbs-id-3667", "Grains / Bread Crumbs"),
    ("grains-pasta-sides/stuffing-id-3668",     "Grains / Stuffing"),
    ("grains-pasta-sides/soup-id-3669",         "Grains / Soup"),
    # ── International ─────────────────────────────────────────
    ("international-id-3671",                   "International"),
    ("international/asian-id-3672",             "International / Asian"),
    # ── Bakery ────────────────────────────────────────────────
    ("bakery-id-3673",                          "Bakery"),
    ("bakery/bread-id-3674",                    "Bakery / Bread"),
    # ── Bulk ──────────────────────────────────────────────────
    ("bulk-id-3675",                            "Bulk"),
    ("bulk/bulk-nuts-seeds-id-3676",            "Bulk / Nuts & Seeds"),
    ("bulk/bulk-grains-id-3677",                "Bulk / Grains"),
    ("bulk/bulk-dried-fruit-id-3678",           "Bulk / Dried Fruit"),
    ("bulk/bulk-candy-id-3679",                 "Bulk / Candy"),
    ("bulk/bulk-baking-id-3680",                "Bulk / Baking"),
    # ── Beverages ─────────────────────────────────────────────
    ("beverages-id-3691",                       "Beverages"),
    ("beverages/coffee-id-3692",                "Beverages / Coffee"),
    ("beverages/tea-id-3693",                   "Beverages / Tea"),
    ("beverages/juices-id-3639",                "Beverages / Juices"),
    ("beverages/kombucha-id-3640",              "Beverages / Kombucha"),
    ("beverages/meal-replacement-id-3641",      "Beverages / Meal Replacement"),
    ("beverages/milk-substitutes-id-3642",      "Beverages / Milk Substitutes"),
    ("beverages/drink-mixes-crystals-id-3695",  "Beverages / Drink Mixes"),
    ("beverages/soda-pop-sport-energy-drinks-id-3698","Beverages / Soda & Energy"),
    ("beverages/nonalcoholic-id-3750",          "Beverages / Non-Alcoholic"),
    ("beverages/water-id-3751",                 "Beverages / Water"),
    # ── Frozen Food ───────────────────────────────────────────
    ("frozen-food-id-3708",                     "Frozen Food"),
    ("frozen-food/frozen-meals-id-3709",        "Frozen / Meals"),
    ("frozen-food/frozen-pizza-id-3710",        "Frozen / Pizza"),
    ("frozen-food/frozen-vegetables-id-3711",   "Frozen / Vegetables"),
    ("frozen-food/frozen-fruit-id-3712",        "Frozen / Fruit"),
    ("frozen-food/frozen-desserts-id-3713",     "Frozen / Desserts"),
    ("frozen-food/ice-cream-id-3714",           "Frozen / Ice Cream"),
    ("frozen-food/frozen-breakfast-id-3715",    "Frozen / Breakfast"),
    ("frozen-food/frozen-meat-seafood-id-3716", "Frozen / Meat & Seafood"),
    # ── Household Goods ───────────────────────────────────────
    ("household-goods-id-3717",                 "Household Goods"),
    ("household-goods/cleaning-supplies-id-3718","Household / Cleaning"),
    ("household-goods/paper-products-id-3719",  "Household / Paper Products"),
    ("household-goods/laundry-id-3720",         "Household / Laundry"),
    ("household-goods/dish-care-id-3721",       "Household / Dish Care"),
    ("household-goods/bags-wrap-id-3722",       "Household / Bags & Wrap"),
    ("household-goods/kitchen-items-id-3723",   "Household / Kitchen"),
    # ── Breakfast & Cereals ───────────────────────────────────
    ("breakfast-cereals-id-3752",               "Breakfast & Cereals"),
    ("breakfast-cereals/cereal-id-3753",        "Breakfast / Cereal"),
    ("breakfast-cereals/oatmeal-hot-cereal-id-3754","Breakfast / Oatmeal"),
    ("breakfast-cereals/granola-bars-id-3755",  "Breakfast / Granola & Bars"),
    # ── Canned Goods ──────────────────────────────────────────
    ("canned-goods-id-3756",                    "Canned Goods"),
    ("canned-goods/canned-vegetables-id-3757",  "Canned / Vegetables"),
    ("canned-goods/canned-fruit-id-3758",       "Canned / Fruit"),
    ("canned-goods/canned-fish-id-3759",        "Canned / Fish"),
    ("canned-goods/canned-beans-id-3760",       "Canned / Beans"),
    ("canned-goods/canned-tomatoes-id-3761",    "Canned / Tomatoes"),
    ("canned-goods/canned-meat-id-3762",        "Canned / Meat"),
    # ── Snacks & Candy ────────────────────────────────────────
    ("snacks-candy-id-3777",                    "Snacks & Candy"),
    ("snacks-candy/chips-id-3778",              "Snacks / Chips"),
    ("snacks-candy/candy-id-3779",              "Snacks / Candy"),
    ("snacks-candy/chocolate-id-3780",          "Snacks / Chocolate"),
    ("snacks-candy/popcorn-id-3781",            "Snacks / Popcorn"),
    ("snacks-candy/nuts-seeds-id-3782",         "Snacks / Nuts & Seeds"),
    ("snacks-candy/dried-fruit-id-3783",        "Snacks / Dried Fruit"),
    ("snacks-candy/snack-bars-id-3784",         "Snacks / Snack Bars"),
    ("snacks-candy/rice-cakes-id-3785",         "Snacks / Rice Cakes"),
    # ── Pet Care ──────────────────────────────────────────────
    ("pet-care-id-3788",                        "Pet Care"),
    ("pet-care/dog-food-id-3789",               "Pet Care / Dog Food"),
    ("pet-care/cat-food-id-3790",               "Pet Care / Cat Food"),
    ("pet-care/pet-treats-id-3791",             "Pet Care / Pet Treats"),
    ("pet-care/pet-supplies-id-3792",           "Pet Care / Pet Supplies"),
    # ── Wellness ──────────────────────────────────────────────
    ("wellness-products-id-3795",               "Wellness"),
    ("wellness-products/vitamins-id-3793",      "Wellness / Vitamins"),
    ("wellness-products/supplements-id-3794",   "Wellness / Supplements"),
    # ── Meat ──────────────────────────────────────────────────
    ("meat-id-3796",                            "Meat"),
    ("meat/beef-id-3797",                       "Meat / Beef"),
    ("meat/pork-id-3798",                       "Meat / Pork"),
    ("meat/chicken-id-3799",                    "Meat / Chicken"),
    ("meat/turkey-id-3800",                     "Meat / Turkey"),
    ("meat/lamb-id-3801",                       "Meat / Lamb"),
    ("meat/sausages-id-3802",                   "Meat / Sausages"),
    # ── Produce ───────────────────────────────────────────────
    ("produce-id-3803",                         "Produce"),
    ("produce/vegetables-id-3807",              "Produce / Vegetables"),
    ("produce/fruit-id-3808",                   "Produce / Fruit"),
    ("produce/herbs-id-3809",                   "Produce / Herbs"),
    ("produce/salads-id-3810",                  "Produce / Salads"),
    ("produce/organic-produce-id-3811",         "Produce / Organic"),
    # ── Seafood ───────────────────────────────────────────────
    ("seafood-id-3804",                         "Seafood"),
    ("seafood/fresh-seafood-id-3812",           "Seafood / Fresh"),
    ("seafood/frozen-seafood-id-3813",          "Seafood / Frozen"),
    # ── Spices & Baking ───────────────────────────────────────
    ("spices-baking-id-3806",                   "Spices & Baking"),
    ("spices-baking/spices-id-3814",            "Spices / Spices"),
    ("spices-baking/baking-id-3815",            "Spices / Baking"),
    ("spices-baking/flour-id-3816",             "Spices / Flour"),
    ("spices-baking/sugar-id-3817",             "Spices / Sugar"),
]

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
    return bool(keys_lc & {"sku", "upc", "name", "price", "productid",
                            "displayname", "itemid", "wasprize", "waspricestring"})


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
#  Fresh St. Market scraper
# ─────────────────────────────────────────────────────────────────
def scrape_fresh_st(playwright):
    p("=" * 60)
    p("  Scraping FRESH ST. MARKET")
    p("=" * 60)

    browser, ctx = make_browser(playwright)
    page = ctx.new_page()
    all_products = {}   # pnum -> dict

    total_cats = len(FRESH_ST_CATEGORIES)
    for cat_i, (cat_slug, cat_name) in enumerate(FRESH_ST_CATEGORIES, 1):
        p(f"\n  [{cat_i}/{total_cats}] {cat_name}")

        # ── Intercept ONE API response ────────────────────────
        intercepted = []

        def _on_resp(resp, _buf=intercepted):
            if _buf:
                return
            try:
                if resp.status != 200:
                    return
                if "json" not in resp.headers.get("content-type", ""):
                    return
                data = resp.json()
                if not isinstance(data, dict):
                    return
                for key in ("products", "items", "data", "results",
                            "records", "catalogItems", "productList"):
                    val = data.get(key)
                    if _looks_like_products(val):
                        _buf.append((resp.url, data, key))
                        return
            except Exception:
                pass

        page.on("response", _on_resp)
        url = f"{FRESH_ST_BASE}/categories/{cat_slug}"
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=35000)
            time.sleep(3)
        except Exception as e:
            p(f"    [!] Load error: {e}")
            page.remove_listener("response", _on_resp)
            continue

        # Scroll once to trigger lazy API calls
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)
        page.evaluate("window.scrollTo(0, 0)")
        time.sleep(0.5)
        page.remove_listener("response", _on_resp)

        # ── Extract ───────────────────────────────────────────
        if intercepted:
            first_url, first_data, products_key = intercepted[0]
            p(f"    API intercepted — paginating...")
            cat_prods = fetch_all_via_api(
                page, first_url, first_data, products_key, cat_name)
        else:
            p("    No API found — using DOM scroll fallback")
            cat_prods = extract_dom_incremental(page, cat_name)

        # ── Merge (deduplicate by product number) ─────────────
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
        p("  [!] Save On Foods is blocking automated access (Cloudflare).")
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

            try:
                page.wait_for_selector("[data-product-id], .product-tile", timeout=12000)
            except PWTimeout:
                break

            cards = (page.query_selector_all("[data-product-id]") or
                     page.query_selector_all(".product-tile") or [])
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

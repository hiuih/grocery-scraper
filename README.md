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

### Optional: proxy for Save-On-Foods

Cloudflare is more likely to block GitHub Actions' shared datacenter IPs than a normal home connection. If the scheduled run starts failing Cloudflare's check repeatedly, route it through a pool of proxies by setting a `PROXY_LIST` GitHub repo secret (Settings → Secrets and variables → Actions).

`PROXY_LIST` is the raw contents of a proxy list, one proxy per line, in `host:port:username:password` format (the format Webshare exports; use `host:port` with no `:username:password` for an unauthenticated proxy):

```
31.59.20.176:6754:myuser:mypass
45.38.107.97:6014:myuser:mypass
```

Every new browser session (and every Cloudflare-challenge retry) rotates to the next proxy in the list, so a single blocked IP doesn't stall the whole run. Leave `PROXY_LIST` unset to run without a proxy (the default). This only applies to `save_on_foods_scraper.py` — Fresh St. Market has no bot protection and doesn't need one. To test locally, `export PROXY_LIST="$(cat your-proxy-list.txt)"` before running the script.

Note: free public proxy lists are not a good fit here — they're unreliable (most entries are already dead) and often have worse Cloudflare reputation than a plain datacenter IP since they're heavily abused by other bots. Use a paid provider's proxies instead, and prefer **residential** over datacenter — Cloudflare's bot-management scoring weighs IP reputation heavily, and residential IPs start with materially more trust than datacenter ranges. Country matters too: this store (`rsid=1982`) is BC-specific, so a non-Canadian exit IP can render a page that looks fine but silently shows an empty/wrong-region storefront.

### If Cloudflare gets aggressive mid-run

Cloudflare's bot management scores every request on IP reputation, TLS fingerprint, JS-challenge results, and behavior (mouse movement, timing) — not just "is this a real browser." Two things follow from that:

- **IP reputation is dynamic and takes 24-72 hours to decay** after a burst of heavy automated traffic. If a run that used to clear challenges in seconds starts needing the visible-browser fallback on almost every category, that's your IP's threat score climbing, not a code bug — the fix is time, not retries. Hammering it harder (more parallel sessions, immediate re-runs) just digs the hole deeper.
- **Concurrency accelerates this.** Running several sharded processes at once from one IP for hours is exactly the sustained-automated-traffic pattern Cloudflare's reputation system flags hardest. Sharding across 2-3 processes is a reasonable one-off speedup for an urgent run; running it routinely (e.g. every week) will likely degrade that IP's standing over time. For the regular weekly cron, a single unsharded run is the safer default.

`patchright` (the stealth Playwright fork this uses) patches the browser/CDP-level tells — the `HeadlessChrome` UA string, `navigator.webdriver`, the `Runtime.enable` leak — which is what gets past Cloudflare's first-pass JS challenge. It does nothing for TLS fingerprinting or behavioral scoring, which is why `human_pause()` in the code adds small randomized mouse moves/scrolls and jittered delays between requests — cheap mitigation for the layers patchright can't touch, not a guarantee.

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

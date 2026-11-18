#!/usr/bin/env python3
"""
price_fetch.py — Find competitor prices for Lumens products via Google Shopping.

Input:
  File 1.csv         -> product name, Lumens URL, image, Lumens price, rank, MFR ID
  IMAP DETAILS.xlsx  -> Sheet1 col C = seller sites
  The Lumens website is NOT opened; all Lumens data comes from the file.

Flow:
  1. Take the top N products by rank from the file.
  2. Google search "<product name> <MFR ID>" (US results; run with a US VPN),
     then click the Shopping tab. Both pages give seller links.
  3. Keep only links whose domain is in the Seller Sites list (IMAP DETAILS.xlsx, Sheet1 col C).
  4. If the seller link has another finish of the product (303005-1009), swap in our
     MFR ID (303005-1007). Open each seller page, read title / price / image from its product data
     (JSON-LD, meta tags, itemprop), and verify it is the same product using
     MFR ID (in title = MATCH), name similarity and image similarity (img_compare.py).
  5. Write one row per matched seller to output/results_<time>.xlsx (Sheet2 layout).

Google captchas: the browser uses a saved profile (.browser_profile/). If Google
shows a captcha, solve it in the Chrome window — the script waits and continues.

Usage:
  .venv/bin/python price_fetch.py                 # File 1.csv, top 10
  .venv/bin/python price_fetch.py --input "File 1.csv" --top 10
  .venv/bin/python price_fetch.py --limit 5       # first 5 products (test run)
  .venv/bin/python price_fetch.py --start 20 --limit 10
  .venv/bin/python price_fetch.py --fresh         # ignore checkpoint, start over
  .venv/bin/python price_fetch.py --debug         # save Google HTML/screenshots to output/debug/
"""

import argparse
import html
import asyncio
import csv
import io
import json
import random
import re
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from PIL import Image
from playwright.async_api import async_playwright
from rapidfuzz import fuzz

import img_compare as ic

ROOT = Path(__file__).resolve().parent
SELLERS_XLSX = ROOT / "IMAP DETAILS.xlsx"
OUT_DIR = ROOT / "output"
PROFILE_DIR = ROOT / ".browser_profile"
CHECKPOINT = OUT_DIR / "checkpoint.jsonl"     # set per input file in main()

GOOGLE_DELAY = (8, 15)        # seconds between Google searches (random)
SELLER_DELAY = (2, 4)         # seconds between seller page visits
MAX_SELLER_PAGES = 15         # per product, cap on seller pages opened
CAPTCHA_WAIT_SEC = 600        # how long to wait for you to solve a captcha
BLOCK_WAIT_SEC = 21           # how long to wait for a seller's bot check to clear
BLOCK_WORDS = ("just a moment", "attention required", "access denied",
               "performing security verification", "verifying you are human",
               "press & hold", "are you a robot", "verify you are human")
_blocked_hosts: dict[str, int] = {}   # seller -> times blocked this run

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

OUT_COLUMNS = [
    "Product Name", "Url", "MFR ID", "Lumens Price", "Seller Site", "Seller Link",
    "Seller Product Title", "Seller Price", "Lumens rank", "Is price Same",
    "Price Diff", "Match", "Match Reason", "Name Score", "Image Verdict", "Source",
]


# ---------------------------------------------------------------- helpers

def log(msg: str):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def norm_host(url: str) -> str:
    host = urllib.parse.urlparse(url).netloc.lower().split(":")[0]
    return host[4:] if host.startswith("www.") else host


def alnum(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def parse_price(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val) if val > 0 else None
    m = re.search(r"\d[\d,]*(?:\.\d{1,2})?", str(val))
    if not m:
        return None
    try:
        p = float(m.group(0).replace(",", ""))
        return p if p > 0 else None
    except ValueError:
        return None


def fix_url(u: str) -> str:
    u = (u or "").strip()
    u = re.sub(r"^https?:/(?!/)", "https://", u)
    return re.sub(r"(?<!:)//+", "/", u)


def load_products(path: Path) -> list[dict]:
    """Products from the input CSV (e.g. "File 1.csv")."""
    rows, seen = [], set()
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            url = fix_url(r.get("Lumnes Product Url"))
            if not url or url in seen:
                continue
            seen.add(url)
            rows.append({
                "name": (r.get("Lumnes Product Name") or "").strip(),
                "url": url,
                "image": fix_url(r.get("Lumnes Product Image Url")),
                "price": parse_price(r.get("Lumens Product Price")),
                "rank": (r.get("Lumens BSR rank") or "").strip(),
                "mfr": (r.get("MFR ID") or "").strip(),
            })
    return rows


def load_sellers() -> dict[str, str]:
    """host -> original seller URL"""
    wb = openpyxl.load_workbook(SELLERS_XLSX, read_only=True)
    ws = wb["Sheet1"]
    sellers = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) >= 3 and row[2]:
            sellers[norm_host(str(row[2]).strip())] = str(row[2]).strip()
    return sellers


def seller_for(url: str, sellers: dict) -> str | None:
    host = norm_host(url)
    for s in sellers:
        if host == s or host.endswith("." + s):
            return s
    return None


def unwrap_google(href: str) -> str | None:
    """Turn Google redirect links into the real destination URL."""
    if not href:
        return None
    href = html.unescape(href)
    if href.startswith("/"):
        href = "https://www.google.com" + href
    p = urllib.parse.urlparse(href)
    if "google." in p.netloc:
        qs = urllib.parse.parse_qs(p.query)
        for key in ("url", "q", "adurl"):
            if key in qs and qs[key][0].startswith("http"):
                return qs[key][0]
        return None   # internal google link (or /aclk ad redirect, resolved separately)
    return href if href.startswith("http") else None


# ---------------------------------------------------------------- image match

_img_cache: dict[str, Image.Image | None] = {}


def get_image(url: str) -> Image.Image | None:
    if not url:
        return None
    if url in _img_cache:
        return _img_cache[url]
    img = None
    try:
        r = requests.get(url, headers={"User-Agent": UA}, timeout=20)
        r.raise_for_status()
        img = Image.open(io.BytesIO(r.content)).convert("RGB")
    except Exception:
        img = None
    _img_cache[url] = img
    return img


def image_verdict(url1: str, url2: str) -> str:
    a, b = get_image(url1), get_image(url2)
    if a is None or b is None:
        return "N/A"
    try:
        res = {
            "pixel_identical": ic.pixel_identical(a, b),
            "phash_distance": ic.phash_distance(a, b),
            "dhash_distance": ic.dhash_distance(a, b),
            "ahash_distance": ic.ahash_distance(a, b),
            "ssim": ic.ssim_score(a, b),
            "orb_score": ic.orb_match_score(a, b),
        }
        return ic.verdict(res)
    except Exception as e:
        return f"ERR {type(e).__name__}"


def image_is_match(verdict: str) -> bool:
    return verdict.startswith(("IDENTICAL", "VISUALLY", "NEARLY", "VERY SIMILAR", "SIMILAR"))


# ---------------------------------------------------------------- page data

EXTRACT_JS = r"""
() => {
  const out = {title: document.title || "", ld: [], meta: {}, itemprop: {}, h1: ""};
  document.querySelectorAll('script[type="application/ld+json"]').forEach(s => out.ld.push(s.textContent));
  document.querySelectorAll('meta').forEach(m => {
    const k = m.getAttribute('property') || m.getAttribute('name') || m.getAttribute('itemprop');
    if (k) out.meta[k.toLowerCase()] = m.getAttribute('content');
  });
  ['price','name','sku','mpn','image'].forEach(k => {
    const el = document.querySelector('[itemprop="'+k+'"]');
    if (el) out.itemprop[k] = el.getAttribute('content') || el.getAttribute('src') || el.textContent.trim();
  });
  const h1 = document.querySelector('h1'); if (h1) out.h1 = h1.textContent.trim();
  out.text = (document.body ? document.body.innerText : "").slice(0, 200000);
  return out;
}
"""


def _walk_ld(obj, found: list):
    if isinstance(obj, list):
        for o in obj:
            _walk_ld(o, found)
    elif isinstance(obj, dict):
        t = obj.get("@type")
        types = t if isinstance(t, list) else [t]
        if "Product" in types or "ProductGroup" in types:
            found.append(obj)
        for k in ("@graph", "mainEntity", "hasVariant", "itemListElement"):
            if k in obj:
                _walk_ld(obj[k], found)


def _offer_price(offers):
    if isinstance(offers, list):
        prices = [p for p in (_offer_price(o) for o in offers) if p]
        return min(prices) if prices else None
    if isinstance(offers, dict):
        for k in ("price", "lowPrice"):
            p = parse_price(offers.get(k))
            if p:
                return p
        if "priceSpecification" in offers:
            return _offer_price(offers["priceSpecification"])
        if "offers" in offers:
            return _offer_price(offers["offers"])
    return None


def parse_product_page(raw: dict, mfr: str) -> dict:
    """Pull title, price, image, identifiers out of the scraped page data."""
    products = []
    for txt in raw.get("ld", []):
        try:
            _walk_ld(json.loads(txt), products)
        except Exception:
            continue

    meta, iprop = raw.get("meta", {}), raw.get("itemprop", {})
    title = raw.get("h1") or meta.get("og:title") or raw.get("title") or ""
    price = image = None
    ids = []

    # Prefer the JSON-LD product/variant that carries the MFR ID, else the first one
    mfr_n = alnum(mfr)
    chosen = None
    for p in products:
        blob = alnum(" ".join(str(p.get(k, "")) for k in ("sku", "mpn", "gtin", "productID", "name")))
        if mfr_n and mfr_n in blob:
            chosen = p
            break
    if chosen is None and products:
        chosen = products[0]

    if chosen:
        title = chosen.get("name") or title
        price = _offer_price(chosen.get("offers"))
        img = chosen.get("image")
        if isinstance(img, list):
            img = img[0] if img else None
        if isinstance(img, dict):
            img = img.get("url") or img.get("contentUrl")
        image = img
        ids = [str(chosen.get(k)) for k in ("sku", "mpn", "gtin", "productID") if chosen.get(k)]
        if price is None:   # variants
            for v in products:
                price = _offer_price(v.get("offers"))
                if price:
                    break

    price = price or parse_price(meta.get("product:price:amount")) \
        or parse_price(meta.get("og:price:amount")) or parse_price(iprop.get("price"))
    image = image or meta.get("og:image") or iprop.get("image")
    ids += [str(iprop[k]) for k in ("sku", "mpn") if iprop.get(k)]

    return {"title": (title or "").strip(), "price": price, "image": image,
            "ids": ids, "text": raw.get("text", "")}


def fix_variant_url(url: str, mfr: str) -> str:
    """URL has another finish of the same product (e.g. utm_content=303005-1009)?
    Swap in our MFR ID (303005-1007): base number + '-' + digits -> mfr."""
    if not mfr or "-" not in mfr:
        return url
    base = mfr.split("-")[0]
    if len(base) < 5:
        return url
    return re.sub(rf"(?<![0-9A-Za-z]){re.escape(base)}-\d+(?!\d)", mfr, url)


def mfr_match(mfr: str, page: dict) -> str:
    """'full' if exact MFR ID found, 'base' if only the base number found, else ''."""
    if not mfr:
        return ""
    hay = alnum(" ".join(page["ids"]) + " " + page["title"] + " " + page["text"])
    if alnum(mfr) in hay:
        return "full"
    base = mfr.split("-")[0]
    if len(base) >= 5 and re.search(rf"(?<!\d){re.escape(base)}(?!\d)", page["text"] + " " + " ".join(page["ids"])):
        return "base"
    return ""


def decide_match(prod: dict, page: dict, img_verdict: str) -> tuple[str, str, int]:
    name_score = int(fuzz.token_set_ratio(prod["name"].lower(), page["title"].lower()))
    mm = mfr_match(prod["mfr"], page)
    img_ok = image_is_match(img_verdict)
    reasons = []
    if mm:
        reasons.append(f"MFR {mm}")
    reasons.append(f"name {name_score}")
    reasons.append(f"img {img_verdict.split(' ')[0]}")

    mfr_in_name = bool(prod["mfr"]) and alnum(prod["mfr"]) in alnum(page["title"])
    if mfr_in_name:
        reasons.insert(0, "MFR in name")

    if mfr_in_name:
        status = "MATCH"            # seller's product name carries the exact MFR ID
    elif mm == "full" and (name_score >= 60 or img_ok):
        status = "MATCH"
    elif mm == "full":
        status = "LIKELY"
    elif mm == "base" and name_score >= 70:
        status = "LIKELY (variant/finish may differ)"
    elif name_score >= 85 and img_ok:
        status = "LIKELY"
    else:
        status = "NO MATCH"
    return status, ", ".join(reasons), name_score


# ---------------------------------------------------------------- browser

async def wait_captcha(page, debug_dir: Path | None):
    if "/sorry/" not in page.url and "captcha" not in (await page.content())[:5000].lower():
        return
    log("⚠️  Google captcha — please solve it in the Chrome window. Waiting...")
    print("\a", end="", flush=True)
    t0 = time.time()
    while time.time() - t0 < CAPTCHA_WAIT_SEC:
        await page.wait_for_timeout(2000)
        if "/sorry/" not in page.url:
            log("Captcha cleared, continuing.")
            await page.wait_for_timeout(2000)
            return
    raise RuntimeError("Captcha not solved in time")


async def accept_consent(page):
    for label in ("Accept all", "I agree", "Reject all"):
        try:
            btn = page.get_by_role("button", name=label)
            if await btn.count():
                await btn.first.click(timeout=2000)
                await page.wait_for_timeout(1000)
                return
        except Exception:
            pass


COLLECT_LINKS_JS = r"""
() => {
  const out = [];
  document.querySelectorAll('a[href]').forEach(a => {
    // Walk up a few levels to get the tile/card text (title + price usually live there)
    let card = a, txt = "";
    for (let i = 0; i < 6 && card; i++) {
      txt = (card.innerText || "").trim();
      if (txt.length > 40) break;
      card = card.parentElement;
    }
    out.push({href: a.getAttribute('href'), text: (a.innerText || "").trim().slice(0, 300),
              card: txt.slice(0, 600)});
  });
  return out;
}
"""

# Google marks each shopping card / organic result with the merchant domain in
# data-dtld. The real link is hidden behind /goto?url=... or /aclk?... redirects.
COLLECT_CARDS_JS = r"""
() => {
  const out = [];
  document.querySelectorAll('[data-dtld]').forEach(el => {
    const domain = el.getAttribute('data-dtld');
    // climb until the container holds a link and some text
    let box = el, links = [];
    for (let i = 0; i < 10 && box; i++) {
      links = Array.from(box.querySelectorAll('a[href]')).map(a => a.getAttribute('href'));
      if (box.hasAttribute('data-dtld') && box !== el) break;
      if (links.length && (box.innerText || '').trim().length > 40) break;
      box = box.parentElement;
    }
    const pcu = box ? Array.from(box.querySelectorAll('[data-pcu]')).map(e => e.getAttribute('data-pcu')) : [];
    out.push({domain, text: box ? (box.innerText || '').trim().slice(0, 600) : '',
              links, pcu});
  });
  return out;
}
"""


async def open_shopping_tab(page) -> bool:
    """Click the 'Shopping' tab on a Google results page, like a person would."""
    await page.wait_for_timeout(random.randint(1000, 2000))
    for sel in ['a[href*="udm=28"]', 'a[href*="tbm=shop"]',
                'a:has-text("Shopping")', 'div[role="listitem"] a:has-text("Shopping")']:
        loc = page.locator(sel).first
        try:
            if await loc.count() and await loc.is_visible():
                await loc.click()
                await page.wait_for_load_state("domcontentloaded", timeout=30000)
                await page.wait_for_timeout(1500)
                if "udm=28" in page.url or "tbm=shop" in page.url:
                    return True
        except Exception:
            continue
    # tab may be under "More"
    try:
        more = page.locator('div[role="navigation"] :text-is("More"), [aria-label="More filters"]').first
        if await more.count():
            await more.click()
            await page.wait_for_timeout(1000)
            loc = page.locator('a:has-text("Shopping")').first
            if await loc.count():
                await loc.click()
                await page.wait_for_load_state("domcontentloaded", timeout=30000)
                return "udm=28" in page.url or "tbm=shop" in page.url
    except Exception:
        pass
    return False


async def google_search(page, query: str, shopping: bool, sellers: dict,
                        debug_dir: Path | None, tag: str) -> list[dict]:
    params = {"q": query, "gl": "us", "hl": "en", "pws": "0"}
    web_url = "https://www.google.com/search?" + urllib.parse.urlencode(params)
    if not shopping or "google.com/search" not in page.url:
        await page.goto(web_url, wait_until="domcontentloaded", timeout=45000)
        await wait_captcha(page, debug_dir)
        await accept_consent(page)
    if shopping:
        if not await open_shopping_tab(page):
            log("   ! Shopping tab not found, opening Shopping URL directly")
            await page.goto(web_url + "&udm=28", wait_until="domcontentloaded", timeout=45000)
        await wait_captcha(page, debug_dir)
        log(f"   Shopping page: {'yes' if 'udm=28' in page.url or 'tbm=shop' in page.url else 'NO'}")
    await page.wait_for_timeout(random.randint(2500, 4000))
    # scroll to trigger lazy content
    for _ in range(3):
        await page.mouse.wheel(0, 1500)
        await page.wait_for_timeout(700)

    links = await page.evaluate(COLLECT_LINKS_JS)
    cards = await page.evaluate(COLLECT_CARDS_JS)

    if debug_dir:
        (debug_dir / f"{tag}_{'shop' if shopping else 'web'}.html").write_text(await page.content())
        await page.screenshot(path=str(debug_dir / f"{tag}_{'shop' if shopping else 'web'}.png"))

    found, seen = [], set()
    ad_links = []
    for l in links:
        href = l.get("href") or ""
        real = unwrap_google(href)
        if real is None and "/aclk" in href:
            ad_links.append(l)
            continue
        if not real:
            continue
        s = seller_for(real, sellers)
        if not s:
            continue
        key = real.split("#")[0]
        if key in seen:
            continue
        seen.add(key)
        m = re.search(r"\$\s?\d[\d,]*(?:\.\d{2})?", l.get("card") or "")
        found.append({"url": key, "seller": s, "google_price": parse_price(m.group(0)) if m else None,
                      "source": "google-shopping" if shopping else "google-web"})

    # Cards tagged with a seller domain: keep the ones that look like this product,
    # then follow Google's redirect to get the real product URL.
    name_part = re.sub(r"\s+\S*\d\S*$", "", query).lower()
    mfr = query.split()[-1] if re.search(r"\d", query.split()[-1]) else ""
    mfr_base = mfr.split("-")[0] if mfr else ""
    picked = []
    for c in cards:
        s = seller_for("https://" + (c.get("domain") or ""), sellers)
        if not s:
            continue
        text = c.get("text") or ""
        if not text:
            continue
        score = fuzz.partial_ratio(name_part, text.lower())
        has_mfr = bool(mfr_base and mfr_base in text)
        if score < 75 and not has_mfr:
            continue
        direct = [u for u in (c.get("pcu") or []) + (c.get("links") or [])
                  if u and u.startswith("http") and "google." not in u and seller_for(u, sellers)]
        redirect = [u for u in c.get("links") or [] if u and (u.startswith("/goto") or "/aclk" in u)]
        redirect.sort(key=lambda u: 0 if u.startswith("/goto") else 1)
        m = re.search(r"\$\s?\d[\d,]*(?:\.\d{2})?", text)
        picked.append({"seller": s, "direct": direct, "redirect": redirect,
                       "google_price": parse_price(m.group(0)) if m else None,
                       "card_text": text, "rank": (has_mfr, score)})

    # best cards first; at most 3 per seller
    picked.sort(key=lambda p: p["rank"], reverse=True)
    per_seller: dict[str, int] = {}
    for p in picked:
        if per_seller.get(p["seller"], 0) >= 3:
            continue
        dest = None
        for u in p["direct"]:
            # data-pcu is often just the site home page — only use real product paths
            if urllib.parse.urlparse(u).path.strip("/"):
                dest = u
                break
        if not dest and p["redirect"]:
            r0 = html.unescape(p["redirect"][0])
            dest = await resolve_redirect(page.context, "https://www.google.com" + r0 if r0.startswith("/") else r0)
        if not dest or not seller_for(dest, sellers):
            continue
        key = html.unescape(dest).split("#")[0]
        if key in seen:
            continue
        seen.add(key)
        per_seller[p["seller"]] = per_seller.get(p["seller"], 0) + 1
        found.append({"url": key, "seller": p["seller"], "google_price": p["google_price"],
                      "card_text": p["card_text"],
                      "source": "google-shopping" if shopping else "google-web"})
    return found


async def resolve_redirect(ctx, url: str) -> str | None:
    p = await ctx.new_page()
    try:
        await p.goto(url, wait_until="domcontentloaded", timeout=25000)
        await p.wait_for_timeout(1500)
        return p.url.split("#")[0]
    except Exception:
        return None
    finally:
        await p.close()


async def scrape_page(ctx, url: str, mfr: str) -> dict | None:
    p = await ctx.new_page()
    try:
        await p.goto(url, wait_until="domcontentloaded", timeout=45000)
        await p.wait_for_timeout(random.randint(2500, 4000))
        # Cloudflare / PerimeterX bot wall: give it a little time to clear by itself
        blocked = False
        for _ in range(BLOCK_WAIT_SEC // 3):
            head = ((await p.title()) + " " + await p.evaluate(
                "document.body ? document.body.innerText.slice(0, 400) : ''")).lower()
            blocked = any(w in head for w in BLOCK_WORDS)
            if not blocked:
                break
            await p.wait_for_timeout(3000)
        if blocked:
            return {"blocked": True, "final_url": p.url}
        raw = await p.evaluate(EXTRACT_JS)
        data = parse_product_page(raw, mfr)
        data["final_url"] = p.url
        return data
    except Exception as e:
        log(f"   ! failed {url[:90]}: {type(e).__name__}")
        return None
    finally:
        await p.close()


# ---------------------------------------------------------------- main loop

def _int(v):
    try:
        return int(v)
    except (TypeError, ValueError):
        return v


def load_checkpoint() -> dict[str, list[dict]]:
    done = {}
    if CHECKPOINT.exists():
        for line in CHECKPOINT.read_text().splitlines():
            try:
                d = json.loads(line)
                done[d["key"]] = d["rows"]   # key = Lumens URL
            except Exception:
                pass
    return done


REPORT_COLUMNS = ["Product Name", "MFR ID", "Url", "Lumens Price", "Seller Link", "Seller Price",
                  "Lumens rank", "Is price Same", "Coupons(optional)"]   # IMAP DETAILS.xlsx Sheet2 + MFR ID


def is_matched(r: dict) -> bool:
    return (r.get("Match") or "").startswith(("MATCH", "LIKELY"))


def write_excel(all_rows: list[dict], path: Path):
    wb = openpyxl.Workbook()

    # Sheet 1: same layout as IMAP DETAILS.xlsx -> Sheet2. Product details on the
    # first seller row, one row per matched seller underneath.
    rep = wb.active
    rep.title = "IMAP Report"
    rep.append(REPORT_COLUMNS)
    order, groups = [], {}
    for r in all_rows:
        k = r.get("Url")
        if k not in groups:
            order.append(k)
            groups[k] = []
        if is_matched(r):
            groups[k].append(r)
    for k in order:
        first = next(r for r in all_rows if r.get("Url") == k)
        sellers = sorted(groups[k], key=lambda r: (not r["Match"].startswith("MATCH"), r.get("Seller Site") or ""))
        if not sellers:
            rep.append([first["Product Name"], first.get("MFR ID"), k, first.get("Lumens Price"),
                        "No seller found", None, _int(first.get("Lumens rank")), None, None])
            continue
        for i, r in enumerate(sellers):
            head = ([first["Product Name"], first.get("MFR ID"), k, first.get("Lumens Price")]
                    if i == 0 else [None, None, None, None])
            rep.append(head + [r.get("Seller Link"), r.get("Seller Price"),
                               _int(first.get("Lumens rank")) if i == 0 else None,
                               r.get("Is price Same"), None])
            cell = rep.cell(row=rep.max_row, column=REPORT_COLUMNS.index("Seller Link") + 1)
            if cell.value and str(cell.value).startswith("http"):
                cell.hyperlink = cell.value
                cell.style = "Hyperlink"
    rep.freeze_panes = "A2"
    for col, w in zip("ABCDEFGHI", (34, 14, 45, 13, 60, 13, 12, 13, 18)):
        rep.column_dimensions[col].width = w
    for c in rep[1]:
        c.font = openpyxl.styles.Font(bold=True)

    # Sheet 2: full detail of every seller page checked (incl. rejected ones)
    ws = wb.create_sheet("Details")
    ws.append(OUT_COLUMNS)
    for r in all_rows:
        ws.append([r.get(c) for c in OUT_COLUMNS])
    ws.freeze_panes = "A2"
    widths = {"Product Name": 34, "Url": 40, "Seller Link": 50, "Seller Product Title": 40,
              "Match Reason": 36, "Image Verdict": 30}
    for i, c in enumerate(OUT_COLUMNS, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = widths.get(c, 14)
    wb.save(path)


async def process_product(ctx, gpage, prod: dict, sellers: dict, debug_dir) -> list[dict]:
    rows = []

    # 1) Lumens data comes from the input file only (Lumens site is not opened)
    lumens_price = prod["price"]
    base = {"Product Name": prod["name"], "Url": prod["url"], "MFR ID": prod["mfr"],
            "Lumens rank": prod["rank"], "Lumens Price": lumens_price}
    log(f"   Lumens price: {lumens_price}   MFR ID: {prod['mfr'] or '-'}")

    # 2) Google web search, then the Shopping tab
    query = f"{prod['name']} {prod['mfr']}".strip()
    tag = alnum(prod["mfr"] or prod["name"])[:30]
    candidates = []
    for shopping in (False, True):     # web search, then click the Shopping tab
        try:
            candidates += await google_search(gpage, query, shopping, sellers, debug_dir, tag)
        except RuntimeError:
            raise
        except Exception as e:
            log(f"   ! google {'shopping' if shopping else 'web'} failed: {type(e).__name__}: {e}")
        if not shopping:
            await gpage.wait_for_timeout(random.randint(2000, 4000))
    await gpage.wait_for_timeout(random.randint(*GOOGLE_DELAY) * 1000)

    # dedupe by URL, prefer shopping source
    uniq = {}
    for c in sorted(candidates, key=lambda c: "shop" not in c["source"].lower()):
        uniq.setdefault(c["url"], c)
    candidates = list(uniq.values())[:MAX_SELLER_PAGES]
    log(f"   {len(candidates)} seller link(s): {sorted({c['seller'] for c in candidates})}")

    # 3) Visit each seller page and verify
    best_per_seller: dict[str, dict] = {}
    for c in candidates:
        fixed = fix_variant_url(c["url"], prod["mfr"])
        if fixed != c["url"]:
            log(f"   variant link fixed: {c['url']} -> {fixed}")
            c["url"] = fixed
        page = None
        if _blocked_hosts.get(c["seller"], 0) < 2:      # stop trying sites that keep blocking us
            page = await scrape_page(ctx, c["url"], prod["mfr"])
            await asyncio.sleep(random.uniform(*SELLER_DELAY))
        via_google = False
        if page is None or page.get("blocked"):
            if page and page.get("blocked"):
                _blocked_hosts[c["seller"]] = _blocked_hosts.get(c["seller"], 0) + 1
            card = c.get("card_text") or ""
            if not card:
                continue
            lines = [l for l in card.split("\n") if l.strip() and not l.strip().startswith("$")]
            page = {"title": lines[0] if lines else "", "price": None, "image": None, "ids": [],
                    "text": card + " " + c["url"], "final_url": c["url"]}
            via_google = True
        verdict = image_verdict(prod["image"], page["image"]) if page.get("image") else "N/A"
        status, reason, name_score = decide_match(prod, page, verdict)
        price = page["price"] or c.get("google_price")
        row = dict(base)
        row.update({
            "Seller Site": c["seller"], "Seller Link": page.get("final_url") or c["url"],
            "Seller Product Title": page["title"][:200], "Seller Price": price,
            "Match": status, "Match Reason": reason, "Name Score": name_score,
            "Image Verdict": verdict,
            "Source": c["source"] + (" (site blocked, google data)" if via_google
                                     else "" if page["price"] else " (google price)"),
        })
        if lumens_price and price:
            row["Price Diff"] = round(price - lumens_price, 2)
            row["Is price Same"] = "Yes" if abs(price - lumens_price) < 0.01 else "No"
        log(f"   {status:<10} {c['seller']:<24} ${price}  ({reason})")
        rank = {"MATCH": 3, "LIKELY": 2}.get(status.split(" ")[0], 1 if "LIKELY" in status else 0)
        prev = best_per_seller.get(c["seller"])
        if prev is None or rank > prev["_rank"]:
            row["_rank"] = rank
            best_per_seller[c["seller"]] = row

    for r in best_per_seller.values():
        r.pop("_rank", None)
        rows.append(r)
    if not rows:
        r = dict(base)
        r["Match"] = "NO SELLER FOUND"
        rows.append(r)
    return rows


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--start", type=int, default=0, help="row index to start from (0-based)")
    ap.add_argument("--limit", type=int, default=0, help="max products to process (0 = all)")
    ap.add_argument("--fresh", action="store_true", help="ignore checkpoint")
    ap.add_argument("--debug", action="store_true", help="save Google pages to output/debug/")
    ap.add_argument("--skip-ip-check", action="store_true")
    ap.add_argument("--top", type=int, default=10, help="only the top N products by rank (0 = all)")
    ap.add_argument("--input", default="File 1.csv",
                    help="product CSV (name, URL, image, Lumens price, rank, MFR ID)")
    args = ap.parse_args()

    OUT_DIR.mkdir(exist_ok=True)
    debug_dir = OUT_DIR / "debug" if args.debug else None
    if debug_dir:
        debug_dir.mkdir(exist_ok=True)

    if not args.skip_ip_check:
        try:
            country = requests.get("https://ipinfo.io/json", timeout=10).json().get("country")
            log(f"Public IP country: {country}")
            if country != "US":
                log("⚠️  Not on a US IP — connect your US VPN for US Shopping results "
                    "(or pass --skip-ip-check to run anyway).")
                sys.exit(1)
        except requests.RequestException:
            log("Could not check IP country; continuing.")

    in_path = Path(args.input)
    if not in_path.is_absolute():
        in_path = ROOT / in_path
    if not in_path.exists():
        log(f"Input file not found: {in_path}")
        sys.exit(1)
    products = load_products(in_path)
    log(f"Input: {in_path.name}")
    if args.top:
        products = [p for p in products if str(p["rank"]).isdigit() and int(p["rank"]) <= args.top]
    products.sort(key=lambda p: int(p["rank"]) if str(p["rank"]).isdigit() else 10**9)

    # separate progress file per input, so runs on different lists don't mix
    global CHECKPOINT
    CHECKPOINT = OUT_DIR / f"checkpoint_{alnum(in_path.stem)}.jsonl"
    sellers = load_sellers()
    log(f"{len(products)} products, {len(sellers)} seller sites")
    products = products[args.start:]
    if args.limit:
        products = products[:args.limit]

    if args.fresh and CHECKPOINT.exists():
        CHECKPOINT.unlink()
    done = load_checkpoint()

    stamp = datetime.now().strftime("%Y%m%d_%H%M")
    out_path = OUT_DIR / f"results_{stamp}.xlsx"

    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(PROFILE_DIR), channel="chrome", headless=False, locale="en-US",
            viewport={"width": 1366, "height": 900},
            args=["--disable-blink-features=AutomationControlled"],
        )
        gpage = ctx.pages[0] if ctx.pages else await ctx.new_page()
        try:
            for i, prod in enumerate(products, start=args.start):
                key = prod["url"]
                if key in done:
                    continue
                log(f"[{i + 1}] {prod['name']} ({prod['mfr']})")
                rows = await process_product(ctx, gpage, prod, sellers, debug_dir)
                done[key] = rows
                with open(CHECKPOINT, "a") as f:
                    f.write(json.dumps({"key": key, "rows": rows}) + "\n")
                write_excel([r for v in done.values() for r in v], out_path)
        except (KeyboardInterrupt, RuntimeError) as e:
            log(f"Stopped: {e or 'interrupted'} — progress saved, re-run to resume.")
        finally:
            await ctx.close()

    write_excel([r for v in done.values() for r in v], out_path)
    log(f"Saved {out_path}")


if __name__ == "__main__":
    asyncio.run(main())

#!/usr/bin/env python3
"""
collect_products.py — Build the Lumens best-seller product list for every brand
in IMAP DETAILS.xlsx (Sheet1, column B "Brands best seller products link").

Output: brands_products.csv, columns: name, url, image, price, rank, MFR ID, domain. Rank = position on the
brand page sorted by best sellers (revenue). MFR ID is left empty here —
price_fetch.py reads it from each Lumens product page.

Re-collecting a brand replaces that brand's rows; other brands are kept.
price_fetch.py runs this automatically for a brand that hasn't been collected yet.

Usage:
  .venv/bin/python collect_products.py
  .venv/bin/python collect_products.py --brand fermob      # one brand only
"""

import argparse
import asyncio
import csv
import re
import urllib.parse

import openpyxl
from playwright.async_api import async_playwright

from price_fetch import BRANDS_CSV, PROFILE_DIR, SELLERS_XLSX, log

OUT_CSV = BRANDS_CSV
CSV_COLUMNS = ["Lumnes Product Name", "Lumnes Product Url", "Lumnes Product Image Url",
               "Lumens Product Price", "Lumens BSR rank", "MFR ID", "Domain"]

TILES_JS = r"""
() => {
  const out = [];
  document.querySelectorAll('div.ProductTile[data-cnstrc-item-id]').forEach(t => {
    const a = t.querySelector('a.ProductTile__link[href]');
    const img = t.querySelector('img.product-image');
    const price = t.querySelector('.ProductPrice');
    out.push({
      id: t.getAttribute('data-cnstrc-item-id'),
      name: t.getAttribute('data-cnstrc-item-name') || (t.querySelector('.ProductName') || {}).innerText || '',
      href: a ? a.href : null,
      img: img ? (img.getAttribute('src') || img.currentSrc) : null,
      price: price ? price.getAttribute('data-cnstrc-item-price') : null,
    });
  });
  const cnt = document.querySelector('[data-testid="plp-show-more-count"]');
  return {tiles: out, count: cnt ? cnt.innerText : ''};
}
"""


def load_brands() -> list[tuple[str, str]]:
    ws = openpyxl.load_workbook(SELLERS_XLSX, read_only=True)["Sheet1"]
    brands = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] and row[1]:
            brands.append((str(row[0]).strip(), str(row[1]).strip()))
    return brands


def brand_slug(url: str) -> str:
    return urllib.parse.urlparse(url).path.strip("/").split("/")[0]


def parent_url(href: str, item_id: str) -> str:
    """Tile links point at a variant; use the product (parent) URL instead."""
    href = href.split("?")[0]
    return re.sub(r"-[A-Za-z0-9]+\.html$", f"-{item_id}.html", href)


async def collect_brand(page, name: str, url: str) -> list[dict]:
    slug = brand_slug(url)
    rows, seen, total = [], set(), None
    for pno in range(1, 100):
        sep = "&" if "?" in url else "?"
        await page.goto(f"{url}{sep}page={pno}", wait_until="domcontentloaded", timeout=60000)
        try:
            await page.wait_for_selector("div.ProductTile[data-cnstrc-item-id]", timeout=20000)
        except Exception:
            log(f"   page {pno}: no products (blocked or end)")
            break
        await page.wait_for_timeout(1500)
        data = await page.evaluate(TILES_JS)
        m = re.search(r"of\s+([\d,]+)", data["count"] or "")
        if m:
            total = int(m.group(1).replace(",", ""))
        new = 0
        for t in data["tiles"]:
            if not t["id"] or not t["href"] or t["id"] in seen:
                continue
            seen.add(t["id"])
            new += 1
            rows.append({
                "Lumnes Product Name": t["name"].strip(),
                "Lumnes Product Url": parent_url(t["href"], t["id"]),
                "Lumnes Product Image Url": t["img"],
                "Lumens Product Price": "",
                "Lumens BSR rank": len(rows) + 1,
                "MFR ID": "",
                "Domain": slug,
            })
        log(f"   page {pno}: +{new} (have {len(rows)}{f' of {total}' if total else ''})")
        if new == 0 or (total and len(rows) >= total):
            break
        await page.wait_for_timeout(2000)
    return rows


async def collect_and_save(brands: list[tuple[str, str]]):
    """Collect the given (name, url) brands and merge them into brands_products.csv."""
    new_rows = []
    async with async_playwright() as pw:
        ctx = await pw.chromium.launch_persistent_context(
            str(PROFILE_DIR), channel="chrome", headless=False, locale="en-US",
            viewport={"width": 1366, "height": 900},
            args=["--disable-blink-features=AutomationControlled"])
        page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        try:
            for name, url in brands:
                log(f"{name}: {url}")
                rows = await collect_brand(page, name, url)
                log(f"{name}: {len(rows)} products")
                new_rows += rows
        finally:
            await ctx.close()

    done_slugs = {r["Domain"] for r in new_rows}
    keep = []
    if OUT_CSV.exists():
        with open(OUT_CSV, newline="", encoding="utf-8") as f:
            keep = [r for r in csv.DictReader(f) if r.get("Domain") not in done_slugs]
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        w.writeheader()
        w.writerows(keep + new_rows)
    log(f"Saved {len(new_rows)} products ({len(keep) + len(new_rows)} total) to {OUT_CSV.name}")


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--brand", help="only this brand (name or url slug); default: every brand in Sheet1")
    args = ap.parse_args()
    brands = load_brands()
    if args.brand:
        b = args.brand.lower()
        brands = [x for x in brands if b in (x[0].lower(), brand_slug(x[1]))]
    await collect_and_save(brands)


if __name__ == "__main__":
    asyncio.run(main())

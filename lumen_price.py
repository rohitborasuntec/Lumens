import pandas as pd
from curl_cffi import requests
import random, os, re
import time, scrapy, json, gc
from urllib.parse import quote, unquote
from src.logging import logger
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from src.automation.driver import Browser
import json_to_csv 

browsers = [
    "chrome120", "chrome123", "chrome124", "chrome131", "chrome133a",
    "chrome99_android", "chrome131_android", "edge99", "edge101",
    "safari153", "safari155", "safari170", "safari172_ios",
    "safari180", "safari180_ios", "firefox133",
]

JSON_DATA_FOLDER = "json_data"
HTML_DATA_FOLDER = "html_data"

PRODUCTS_JSON_FILE = "temp_data.json"
PROGRESS_FILE     = "progress.json"          # <-- new tracker

FLUSH_EVERY = 25
RESTART_EVERY_LINKS = 5
MAX_PRODUCTS_PER_LINK = 96


class Lumens:

    def __init__(self, file_path):
        self.df = pd.read_excel(file_path, sheet_name="Sheet1")
        self.comp_domain = self.df["Seller Sites"].str.extract(
            r'(?:https?://)?(?:www\.)?([^./]+)'
        )[0].tolist()
        self.bsr_links = self.df["Brands best seller products link."].tolist()

        self.records = self._load_records()
        self._dirty = 0
        self._mfr_index = {r.get("MFR ID"): i
                           for i, r in enumerate(self.records)
                           if r.get("MFR ID")}

        self.progress = self._load_progress()

        self.browser = Browser()
        self.driver, _ = self.browser.launch("Chrome")

    # ------------------------------------------------------------------ #
    #  Single-JSON product store
    # ------------------------------------------------------------------ #
    def _load_records(self):
        if not os.path.isfile(PRODUCTS_JSON_FILE):
            return []
        try:
            with open(PRODUCTS_JSON_FILE, encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, list) else []
        except Exception as e:
            logger.error(f"[store] failed to load {PRODUCTS_JSON_FILE}: {e}")
            return []

    def _flush_records(self, force=False):
        if not force and self._dirty < FLUSH_EVERY:
            return
        tmp = PRODUCTS_JSON_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.records, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PRODUCTS_JSON_FILE)
        self._dirty = 0
        logger.info(f"[store] flushed {len(self.records)} records")

    def save_product_record(self, record):
        mfr = record.get("MFR ID")
        if mfr and mfr in self._mfr_index:
            idx = self._mfr_index[mfr]
            changed = False
            for k, v in record.items():
                if v not in (None, "") and self.records[idx].get(k) != v:
                    self.records[idx][k] = v
                    changed = True
            if changed:
                self._dirty += 1
                self._flush_records(force=True)
        else:
            if mfr:
                self._mfr_index[mfr] = len(self.records)
            self.records.append(record)
            self._dirty += 1
            self._flush_records(force=True)
    # ------------------------------------------------------------------ #
    #  Progress tracker
    # ------------------------------------------------------------------ #
    def _load_progress(self):
        if not os.path.isfile(PROGRESS_FILE):
            return {}
        try:
            with open(PROGRESS_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.info(f"[progress] failed to load: {e}")
            return {}

    def _flush_progress(self):
        tmp = PROGRESS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self.progress, f, ensure_ascii=False, indent=2)
        os.replace(tmp, PROGRESS_FILE)

    def _get_domain_progress(self, domain):
        return self.progress.get(domain)

    def _mark_product_limit_reached(self, state, domain):
        state["status"] = "completed"
        state["cumulative"] = MAX_PRODUCTS_PER_LINK
        state["next_rank"] = MAX_PRODUCTS_PER_LINK + 1
        state["pending_products"] = []
        self._flush_progress()
        logger.info(f"[{domain}] product limit reached ({MAX_PRODUCTS_PER_LINK}) — completed")

    def _init_domain_progress(self, domain):
        self.progress[domain] = {
            "status":            "in_progress",
            "page":              1,       # page we are currently on
            "cumulative":        0,       # products from fully-finished pages
            "next_rank":         1,       # rank for the next product on current page
            "pending_products":  [],      # products of current page not yet written
        }
        self._flush_progress()

    # ------------------------------------------------------------------ #
    def wait_for_element(self, locator, timeout=30, condition="visible",
                         poll=0.5, raise_on_timeout=False):
        conditions = {
            "presence":  EC.presence_of_element_located,
            "visible":   EC.visibility_of_element_located,
            "clickable": EC.element_to_be_clickable,
            "invisible": EC.invisibility_of_element_located,
        }
        if condition not in conditions:
            raise ValueError(f"Unknown condition: {condition}")
        wait = WebDriverWait(self.driver, timeout, poll_frequency=poll)
        try:
            return wait.until(conditions[condition](locator))
        except TimeoutException:
            if raise_on_timeout:
                raise
            return None

    # ------------------------------------------------------------------ #
    def _driver_alive(self):
        try:
            _ = self.driver.current_url
            return True
        except Exception:
            return False

    def restart_driver(self):
        logger.info("[driver] restarting…")
        self.browser.quit("Chrome")
        gc.collect()
        self.driver, _ = self.browser.launch("Chrome")

    # ------------------------------------------------------------------ #
    def safe_get(self, url, attempts=2):
        for i in range(1, attempts + 1):
            if not self._driver_alive():
                self.restart_driver()
            try:
                self.driver.get(url)
                return True
            except (TimeoutException, WebDriverException) as e:
                logger.info(f"[get] attempt {i}/{attempts} failed: {e}")
                try:
                    self.driver.execute_script("window.stop();")
                except Exception:
                    pass
                if i == attempts:
                    self.restart_driver()
                time.sleep(random.uniform(1.0, 3.0))
        return False

    def call_with_retry(self, url, attempts=3):
        RETRYABLE = {429, 500, 502, 503, 504}
        for attempt in range(1, attempts + 1):
            logger.info(f"Attempt {attempt}/{attempts}")
            browser = random.choice(browsers)
            try:
                response = requests.get(url, impersonate=browser, timeout=15)
            except requests.RequestException as e:
                logger.info(f"Request error: {e}")
                response = None

            if response is not None and response.status_code == 200:
                return response
            if response is not None and response.status_code == 404:
                logger.info("Error 404 — not retrying")
                return None
            if attempt == attempts:
                break

            backoff = 2 ** (attempt - 1)
            sleep_time = random.uniform(0, backoff)
            if response is not None and response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                if retry_after:
                    try:
                        sleep_time = max(float(retry_after), sleep_time)
                    except ValueError:
                        pass
            sleep_time = min(sleep_time, 30.0)
            logger.info(f"Retrying in {sleep_time:.2f}s (status={getattr(response, 'status_code', 'N/A')})")
            time.sleep(sleep_time)

        logger.info("All attempts wasted")
        return None

    # ------------------------------------------------------------------ #
    def _clean_file_name(self, file_name: str) -> str:
        decoded = unquote(str(file_name))
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', decoded)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        max_len = 200
        if len(cleaned) > max_len:
            cleaned = cleaned[:max_len]
        return cleaned

    def save_json(self, json_data, file_name, folder):
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".json"):
            cleaned += ".json"
        target_dir = os.path.join(JSON_DATA_FOLDER, folder.lstrip("/"))
        os.makedirs(target_dir, exist_ok=True)
        full_path = os.path.join(target_dir, cleaned)
        with open(full_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        return full_path

    def check_for_json(self, file_name, folder):
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".json"):
            cleaned += ".json"
        full_path = os.path.join(JSON_DATA_FOLDER, folder.lstrip("/"), cleaned)
        if os.path.isfile(full_path):
            return full_path
        return None

    def check_for_html(self, file_name, folder):
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".html"):
            cleaned += ".html"
        full_path = os.path.join(HTML_DATA_FOLDER, folder.lstrip("/"), cleaned)
        if os.path.isfile(full_path):
            return full_path
        return None

    def save_html(self, html_content, file_name, folder):
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".html"):
            cleaned += ".html"
        target_dir = os.path.join(HTML_DATA_FOLDER, folder.lstrip("/"))
        os.makedirs(target_dir, exist_ok=True)
        full_path = os.path.join(target_dir, cleaned)
        with open(full_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return full_path

    # ------------------------------------------------------------------ #
    #  HTML parsing
    # ------------------------------------------------------------------ #
    def extract_details(self, page_source):
        response = scrapy.Selector(text=page_source)

        mfr_id = response.xpath(
            '//div[@data-testid="pdp-manufacturer-id"]//div[@class="id-number"]/text()'
        ).extract_first()
        if mfr_id:
            mfr_id = mfr_id.strip()

        price_re = re.compile(r'"price"\s*:\s*"?([0-9][0-9,]*\.?[0-9]*)"?')
        price = None
        for m in price_re.finditer(page_source):
            val = m.group(1).replace(",", "").strip()
            if val and val not in ("0", "0.0", "0.00"):
                price = m.group(1).strip()
                break

        if not price:
            # fallback to rendered text
            price_candidates = response.xpath(
                '//*[contains(concat(" ", normalize-space(@class), " "), " price ")]//text()'
            ).extract()
            for c in price_candidates:
                c = c.strip()
                if c and re.search(r'\d', c) and ('$' in c or re.search(r'\d+\.\d{2}', c)):
                    price = c
                    break
        
        logger.info(f"MFR ID : {mfr_id} and Price : {price}")

        return mfr_id, price

    # ------------------------------------------------------------------ #
    #  Process one product (used by both fresh & resumed flows)
    # ------------------------------------------------------------------ #
    def _process_one_product(self, product, link, rank):
        """Returns True if the record was saved (or already existed)."""
        folder = link.split(".com")[-1].split("/?")[0]
        lumens_url = "https://www.lumens.com/" + product["url"]
        html_name = lumens_url.split(".com")[-1]

        existing_html = self.check_for_html(html_name, folder)
        if existing_html is None:
            if not self.safe_get(lumens_url):
                logger.info(f"Skipping (could not load): {lumens_url}")
                return False
            if self.wait_for_element((By.CLASS_NAME, "productName")) is None:
                logger.info(f"Skipping (no productName): {lumens_url}")
                return False
            page_source = self.driver.page_source
            # self.save_html(page_source, html_name, folder)
        else:
            with open(existing_html, encoding="utf-8") as f:
                page_source = f.read()

        mfr_id, price = self.extract_details(page_source)
        del page_source

        if not price:
            logger.info(f"[price] missing for {lumens_url}, retrying via Selenium…")
            if self.safe_get(lumens_url) and \
               self.wait_for_element((By.CLASS_NAME, "productName")) is not None:
                mfr_id, price = self.extract_details(self.driver.page_source)
                if price:
                    logger.info(f"[price] recovered: {price}")

        record = {
            "Lumnes Product Name":      product.get("name"),
            "Lumnes Product Url":       lumens_url,
            "Lumnes Product Image Url": product.get("image"),
            "Lumens Product Price":     price.replace(",",""),
            "Lumens BSR rank":          rank,
            "MFR ID":                   mfr_id,
            "Domain":                   link.split("com/")[-1].split("/")[0],
        }
        self.save_product_record(record)
        return True

    # ------------------------------------------------------------------ #
    #  Fetch (or load cached) product list for a given page
    # ------------------------------------------------------------------ #
    def _fetch_page_products(self, link, folder, page_number):
        go_to_link = (
            link
            + f"&c_composable=true&siteId=Lumens"
              f"&c_page={page_number}"
              f"&c_sortOption=revenue%2Bdescending"
              f"&c_queryType=category"
              f"&c_constructorClientId=5d888221-bf57-45e3-add7-3dc217039ee8"
              f"&c_constructorSessionId=8"
              f"&c_pageUrl={quote(link, safe='')}"
        )
        param = folder + str(page_number)

        existing = self.check_for_json(param, folder)
        if existing is None:
            if not self.safe_get(go_to_link):
                return None
            response = scrapy.Selector(text=self.driver.page_source)
            pre_text = response.xpath("//pre/text()").extract_first()
            del response
            if pre_text is None:
                return None
            data = json.loads(pre_text)
            del pre_text
            # path = self.save_json(data, param, folder)
            # logger.info(f"Saved to {path}")
        else:
            logger.info(f"Using cached: {existing}")
            with open(existing, encoding="utf-8") as f:
                data = json.load(f)

        obj = data.get("object") or {}
        products = obj.get("products") or []
        del data
        return products

    # ------------------------------------------------------------------ #
    #  Main per-link logic with tracking + resume
    # ------------------------------------------------------------------ #
    def parse_page(self, link):
        folder = link.split(".com")[-1].split("/?")[0]
        domain = link.split("com/")[-1].split("/")[0]

        state = self._get_domain_progress(domain)
        if state is None:
            self._init_domain_progress(domain)
            state = self.progress[domain]

        if state["status"] == "completed":
            logger.info(f"[{domain}] already completed — skipping")
            return

        page_number = state["page"]
        cumulative  = state["cumulative"]
        next_rank   = state["next_rank"]

        if next_rank > MAX_PRODUCTS_PER_LINK:
            self._mark_product_limit_reached(state, domain)
            return

        # ---- First: drain any pending products left from a crash ----- #
        if state.get("pending_products"):
            logger.info(f"[{domain}] resuming page {page_number} "
                  f"with {len(state['pending_products'])} pending product(s)")
            slots_remaining = MAX_PRODUCTS_PER_LINK - next_rank + 1
            pending = list(state["pending_products"])[:slots_remaining]
            state["pending_products"] = pending
            self._flush_progress()
            for product in list(pending):
                ok = self._process_one_product(product, link, next_rank)
                if ok:
                    # remove from pending, bump rank
                    state["pending_products"].remove(product)
                    next_rank += 1
                    state["next_rank"] = next_rank
                    self._flush_progress()
                else:
                    # keep it pending; will retry on next run
                    logger.info(f"[{domain}] product kept pending: {product.get('url')}")
            if next_rank > MAX_PRODUCTS_PER_LINK:
                self._mark_product_limit_reached(state, domain)
                return
            if state["pending_products"]:
                logger.info(f"[{domain}] page {page_number} still has "
                            f"{len(state['pending_products'])} pending product(s) — "
                            "will retry next run")
                return
            # after successful drain, page is complete
            cumulative += 0  # rank already accounted for
            page_number += 1
            state["page"]      = page_number
            state["cumulative"] = cumulative
            state["next_rank"]  = next_rank
            self._flush_progress()
            logger.info(f"[{domain}] page {page_number - 1} drained → moving to page {page_number}")

        # ---- Continue until the API returns an empty page or the 96-item
        #      per-link cap is reached. -------------------------------- #
        while True:
            products = self._fetch_page_products(link, folder, page_number)

            if products is None:
                # fetch failed — leave state, exit (resume next run)
                logger.info(f"[{domain}] page {page_number} fetch failed — will retry next run")
                return

            if not products:
                # last page reached
                logger.info(f"[{domain}] reached last page (page {page_number} empty) — completed")
                state["status"] = "completed"
                state["pending_products"] = []
                self._flush_progress()
                return

            slots_remaining = MAX_PRODUCTS_PER_LINK - next_rank + 1
            products = products[:slots_remaining]

            # register products as pending BEFORE processing
            state["pending_products"] = list(products)
            state["next_rank"] = next_rank
            self._flush_progress()

            for product in products:
                ok = self._process_one_product(product, link, next_rank)
                if ok:
                    state["pending_products"].remove(product)
                    next_rank += 1
                    state["next_rank"] = next_rank
                    self._flush_progress()
                else:
                    logger.info(f"[{domain}] product kept pending: {product.get('url')}")

            if next_rank > MAX_PRODUCTS_PER_LINK:
                self._mark_product_limit_reached(state, domain)
                return

            # only advance page if ALL products of this page were processed
            if state["pending_products"]:
                logger.info(f"[{domain}] page {page_number} incomplete — "
                      f"{len(state['pending_products'])} pending — will resume next run")
                return

            cumulative = next_rank - 1
            page_number += 1
            state["page"]       = page_number
            state["cumulative"] = cumulative
            state["next_rank"]  = next_rank
            self._flush_progress()
            gc.collect()

    # ------------------------------------------------------------------ #
    def parse(self):
        for i, link in enumerate(self.bsr_links):
            self.parse_page(link)
            if i and i % RESTART_EVERY_LINKS == 0:
                self.restart_driver()
            gc.collect()
        self._flush_records(force=True)


if __name__ == "__main__":
    file_path = "IMAP DETAILS.xlsx"
    lumen = Lumens(file_path)
    try:
        lumen.parse()
    except Exception as e:
        logger.info(f"[fatal] {e}")
        import traceback; traceback.logger.info_exc()
    finally:
        try:
            lumen._flush_records(force=True)
            lumen._flush_progress()
            json_to_csv.build_csv("temp_data.json","File.csv")
            logger.info("File Saved")
        except Exception as e:
            logger.error("Error ", e)
        lumen.browser.quit("Chrome")
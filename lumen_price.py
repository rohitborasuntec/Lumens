import pandas as pd
from curl_cffi import requests
import random , os ,re
import time,scrapy,json
import undetected_chromedriver as uc
from urllib.parse import quote

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By


browsers = [
"chrome120",
"chrome123",
"chrome124",
"chrome131",
"chrome133a",
"chrome99_android",
"chrome131_android",
"edge99",
"edge101",
"safari153",
"safari155",
"safari170",
"safari172_ios",
"safari180",
"safari180_ios",
"firefox133",
]

JSON_DATA_FOLDER = "json_data" 
HTML_DATA_FOLDER = "html_data"

class Lumens:

    def __init__(self , file_path):
        self.df = pd.read_excel(file_path,sheet_name="Sheet1")
        self.comp_domain = self.df["Seller Sites"].str.extract(r'(?:https?://)?(?:www\.)?([^./]+)')[0].tolist()
        self.bsr_links = self.df["Brands best seller products link."].tolist()
        self.driver = self.open_chrome()
        self.products = []   
        
    def wait_for_element(self, locator, timeout=10, condition="visible", poll=0.5, raise_on_timeout=False):
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
    
    def open_chrome(self):
        options = uc.ChromeOptions()
        options.page_load_strategy = "eager"   # don't wait for all subresources
        options.add_argument("--disable-gpu")
        options.add_argument("--no-first-run")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--blink-settings=imagesEnabled=false")  # skip images (faster)
        driver = uc.Chrome(version_main=152, options=options)   # ← real version
        driver.set_page_load_timeout(30)   # raise TimeoutException after 60s
        return driver

    def close_chrome(self):
        self.driver.quit()

    def safe_get(self, url, attempts=2):
        for i in range(1, attempts + 1):
            if not self._driver_alive():
                print("[driver] dead, restarting")
                self.restart_driver()
            try:
                self.driver.get(url)
                return True
            except (TimeoutException, WebDriverException) as e:
                print(f"[get] attempt {i}/{attempts} failed: {e}")
                try:
                    self.driver.execute_script("window.stop();")
                except Exception:
                    pass
                time.sleep(random.uniform(1.0,5.0))
        return False
    
    def call_with_retry(self, url, attempts=3):
        # Retry on these status codes (429 = rate limited, 5xx = server errors)
        RETRYABLE = {429, 500, 502, 503, 504}

        for attempt in range(1, attempts + 1):
            print(f"Attempt {attempt}/{attempts}")
            browser = random.choice(browsers)

            try:
                response = requests.get(url, impersonate=browser, timeout=15)
            except requests.RequestException as e:
                print(f"Request error: {e}")
                response = None

            # Success
            if response is not None and response.status_code == 200:
                return response

            # Non-retryable client error -> bail out immediately
            if response is not None and response.status_code == 404:
                print("Error 404 — not retrying")
                return None

            # If last attempt, don't sleep
            if attempt == attempts:
                break

            # Base exponential backoff: 1s, 2s, 4s, ...
            backoff = 2 ** (attempt - 1)

            sleep_time = random.uniform(0, backoff)

            if response is not None and response.status_code == 429:
                retry_after = response.headers.get("Retry-After")
                if retry_after:
                    try:
                        sleep_time = max(float(retry_after), sleep_time)
                    except ValueError:
                        # Retry-After can be an HTTP-date; ignore for simplicity
                        pass
            # ---------------------------------------------------------------

            # Cap the wait so we never sleep absurdly long
            sleep_time = min(sleep_time, 30.0)

            print(f"Retrying in {sleep_time:.2f}s (status={getattr(response, 'status_code', 'N/A')})")
            time.sleep(sleep_time)

        print("All attempts are wasted")
        return None

    def save_to_csv(self,df, file_name="File.csv"):
        first_time = not os.path.exists(file_name)
        df.to_csv(file_name, mode='a', header=first_time, index=False)
        df = df.drop_duplicates('MFR ID' , keep="last") 

    def search_google(self):
        pass

    def _clean_file_name(self,file_name: str) -> str:
        """
        Clean a raw file name / URL param string into a safe file name.

        Handles cases like:
        'skip=0&count=100&orderBy=name%20desc&category=strong&text=Reykjavik%20Spirits&sizelower=500&sizehigher=500'
        Produces something like:
        'skip=0&count=100&orderBy=name desc&category=strong&text=Reykjavik Spirits&sizelower=500&sizehigher=500.json'

        - URL-decodes %20 -> space
        - Strips illegal filesystem chars (< > : " / \\ | ? *)
        - Collapses whitespace
        - Truncates long names (leaves room for '.json')
        """
        from urllib.parse import unquote

        # Decode URL-encoding (e.g. %20 -> space, %2F -> /)
        decoded = unquote(str(file_name))

        # Remove or replace characters illegal on Windows/Linux filenames
        cleaned = re.sub(r'[<>:"/\\|?*]', '_', decoded)

        # Collapse whitespace
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()

        # Truncate to a safe length (leave 5 chars for ".json")
        max_len = 200
        if len(cleaned) > max_len:
            cleaned = cleaned[:max_len]

        return cleaned

    def extract_details(self,page_source):

        response = scrapy.Selector(text=page_source)
        
        mfr_id = response.xpath('//div[@data-testid="pdp-manufacturer-id"]//div[@class="id-number"]/text()').extract_first()
        price = response.xpath('//span[@class="price"]/text()').extract_first()

        return mfr_id,price

    # def check_for_json(self,file_name,folder):
    #     """
    #     Clean file_name the same way as save_json and look it up in JSON_DATA_FOLDER.
    #     Returns the full path if found, otherwise None.
    #     """
    #     cleaned = self._clean_file_name(file_name)
    #     if not cleaned.endswith(".json"):
    #         cleaned += ".json"

    #     full_path = os.path.join(JSON_DATA_FOLDER, folder,cleaned)

    #     if os.path.isfile(full_path):
    #         return full_path
    #     return None

    def save_json(self, json_data, file_name, folder):
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".json"):
            cleaned += ".json"

        target_dir = os.path.join(JSON_DATA_FOLDER, folder.lstrip("/"))  # ← sanitise subfolder
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
    
    # def save_json(self,json_data, file_name):
    #     """
    #     Save json_data to JSON_DATA_FOLDER/<cleaned file_name>.json.
    #     Creates the folder if it doesn't exist.
    #     Returns the full path to the saved file.
    #     """
    #     os.makedirs(JSON_DATA_FOLDER, exist_ok=True)

    #     cleaned = self._clean_file_name(file_name)
    #     if not cleaned.endswith(".json"):
    #         cleaned += ".json"

    #     full_path = os.path.join(JSON_DATA_FOLDER, cleaned)

    #     with open(full_path, "w", encoding="utf-8") as f:
    #         json.dump(json_data, f, ensure_ascii=False, indent=2)

    #     return full_path


    def check_for_html(self, file_name, folder):
        """
        Clean file_name the same way as save_html and look it up in HTML_DATA_FOLDER/<folder>/.
        Returns the full path if found, otherwise None.
        """
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".html"):
            cleaned += ".html"

        full_path = os.path.join(HTML_DATA_FOLDER, folder.lstrip("/"), cleaned)
        if os.path.isfile(full_path):
            return full_path
        return None

    def save_html(self, html_content, file_name, folder):
        """
        Save html_content to HTML_DATA_FOLDER/<folder>/<cleaned file_name>.html.
        Creates the folder if it doesn't exist.
        Returns the full path to the saved file.
        """
        cleaned = self._clean_file_name(file_name)
        if not cleaned.endswith(".html"):
            cleaned += ".html"

        target_dir = os.path.join(HTML_DATA_FOLDER, folder.lstrip("/"))
        os.makedirs(target_dir, exist_ok=True)

        full_path = os.path.join(target_dir, cleaned)
        with open(full_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        return full_path

    def _driver_alive(self):
        try:
            _ = self.driver.current_url
            return True
        except Exception:
            return False

    def restart_driver(self):
        try:
            self.driver.quit()
        except Exception:
            pass
        self.driver = self.open_chrome()
    
    def parse_page(self, link):

        page_number = 1
        file_name = "File.csv"
        
        if os.path.exists(file_name):
            df = pd.read_csv(file_name)
            domain = link.split("com/")[-1].split("/")[0]
        else:
            df = pd.DataFrame(columns=[ "Lumnes Product Name", "Lumnes Product Url", "Lumnes Product Image Url", "Lumens Product Price", "Lumens BSR rank", "MFR ID", "Domain", ])
            domain = "@"
        if domain not in list(df["Domain"].unique()):
            while True:
                go_to_link = link + f"&c_composable=true&siteId=Lumens&c_page={page_number}&c_sortOption=revenue%2Bdescending&c_queryType=category&c_constructorClientId=5d888221-bf57-45e3-add7-3dc217039ee8&c_constructorSessionId=8&c_pageUrl={quote(link, safe='')}"

                folder = link.split(".com")[-1].split("/?")[0]
                param = folder + str(page_number)

                existing = self.check_for_json(param, folder)
                
                if existing is None:
                    self.driver.get(go_to_link)
                    response = scrapy.Selector(text=self.driver.page_source)
                    pre_text = response.xpath("//pre/text()").extract_first()

                    if pre_text is None:
                        break

                    data = json.loads(pre_text)
                    path = self.save_json(data, param, folder)
                    
                    print(f"Saved to {path}")
                else:
                    print(f"Using cached: {existing}")
                    with open(existing, encoding="utf-8") as f:
                        data = json.load(f)

                obj = data.get("object") or {}
                products = obj.get("products") or []

                if not products:
                    break

                self.products.extend(products)
                page_number += 1

            rank = 1

            for product in self.products:
                lumens_url = "https://www.lumens.com/" + product["url"]
                # HTML cache, same layout as JSON
                html_name = lumens_url.split(".com")[-1]
                # parts = html_name.strip("/").split("/")
                folder = link.split(".com")[-1].split("/?")[0]
                # html_folder = parts[0] if parts and parts[0] else "products"
                existing_html = self.check_for_html(html_name, folder)
                if existing_html is None:
                    if not self.safe_get(lumens_url):
                        print(f"Skipping (could not load): {lumens_url}")
                        rank += 1
                        continue
                    if self.wait_for_element((By.CLASS_NAME, "productName")) is None:
                        print(f"Skipping (no productName): {lumens_url}")
                        rank += 1
                        continue
                    page_source = self.driver.page_source
                    path = self.save_html(page_source, html_name, folder)
                    print(f"Saved to {path}")
                else:
                    print(f"Using cached HTML: {existing_html}")
                    with open(existing_html, encoding="utf-8") as f:
                        page_source = f.read()          # ← the fix
                
                product_name = product.get("name")

                mfr_id, price = self.extract_details(page_source)

                if not price:
                    if not self.safe_get(lumens_url):
                        print(f"Skipping (could not load): {lumens_url}")
                        rank += 1
                        continue
                    if self.wait_for_element((By.CLASS_NAME, "productName")) is None:
                        print(f"Skipping (no productName): {lumens_url}")
                        rank += 1
                        continue
                    page_source = self.driver.page_source
                    path = self.save_html(page_source, html_name, folder)
                    print(f"Saved to {path}")

                row = {
                    "Lumnes Product Name": product_name,
                    "Lumnes Product Url": lumens_url,
                    "Lumnes Product Image Url": product["image"],
                    "Lumens Product Price": price,
                    "Lumens BSR rank": rank,
                    "MFR ID": mfr_id,
                }

                df = pd.DataFrame([row])
                self.save_to_csv(df)
                rank += 1

        else:
            print("Parsing from the csv")
            df['Lumens Product Price'] = df['Lumens Product Price'].astype("string").fillna("")

            for index,row in df.iterrows():
                lumens_url = row["Lumnes Product Url"]
                # HTML cache, same layout as JSON
                html_name = lumens_url.split(".com")[-1]
                # parts = html_name.strip("/").split("/")
                folder = link.split(".com")[-1].split("/?")[0]
                # html_folder = parts[0] if parts and parts[0] else "products"
                existing_html = self.check_for_html(html_name, folder)
                price = row["Lumens Product Price"]
                
                if existing_html is None:
                    # if not self.safe_get(lumens_url):
                    #     print(f"Skipping (could not load): {lumens_url}")
                    #     rank += 1
                    #     continue
                    # if self.wait_for_element((By.CLASS_NAME, "productName")) is None:
                    #     print(f"Skipping (no productName): {lumens_url}")
                    #     rank += 1
                    #     continue
                    page_source = self.driver.page_source
                    path = self.save_html(page_source, html_name, folder)
                    print(f"Saved to {path}")
                else:
                    print(f"Using cached HTML: {existing_html}")
                    with open(existing_html, encoding="utf-8") as f:
                        page_source = f.read()  

                if not price :
                    print("Price not found")
                    mfr_id, price = self.extract_details(page_source)
                    print(f"Price Found :{price}")
                    df.at[index,"Lumens Product Price"] = price


            df.to_csv(file_name,index=False)
                # row = {
                #     "Lumnes Product Name": product_name,
                #     "Lumnes Product Url": lumens_url,
                #     "Lumnes Product Image Url": product["image"],
                #     "Lumens Product Price": price,
                #     "Lumens BSR rank": rank,
                #     "MFR ID": mfr_id,
                # }

                # df = pd.DataFrame([row])
                # self.save_to_csv(df)
                # rank += 1

    def parse(self):

        for link in self.bsr_links[:1]:
            self.products = [] 
            self.parse_page(link)
    
            

if __name__ == "__main__":
    
    file_path = "IMAP DETAILS.xlsx"
    
    # try:
    lumen = Lumens(file_path)
    lumen.parse()
    # except:
    #     print("Errror")

    # finally:
    lumen.close_chrome()





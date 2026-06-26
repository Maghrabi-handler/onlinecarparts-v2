import os
import re
import time
import shutil
import pandas as pd

from datetime import datetime
from playwright.sync_api import sync_playwright


# ==========================================================
# FILES
# ==========================================================

INPUT_FILE = r"C:\Users\Handler-Sm\OneDrive\Desktop\parts.xlsx"

OUTPUT_FILE = r"C:\Users\Handler-Sm\OneDrive\Desktop\parts_result_final.xlsx"

BACKUP_FOLDER = r"C:\Users\Handler-Sm\OneDrive\Desktop\Backup"

MAX_RETRIES = 3

AUTO_SAVE = True


# ==========================================================
# GLOBAL VARIABLES
# ==========================================================

results = []

backup_created = False

start_time = time.time()
# ==========================================================
# BRAND NORMALIZER
# ==========================================================

def normalize_brand(brand):

    brand = str(brand).upper().strip()

    replacements = {

        "WIX": "WIX FILTERS",

        "MANN": "MANN-FILTER",

        "MANN FILTER": "MANN-FILTER"

    }

    return replacements.get(brand, brand)


# ==========================================================
# BACKUP
# ==========================================================

def ensure_backup_folder():

    if not os.path.exists(BACKUP_FOLDER):

        os.makedirs(BACKUP_FOLDER)


def create_backup():

    global backup_created

    if backup_created:

        return

    if not os.path.exists(OUTPUT_FILE):

        return

    ensure_backup_folder()

    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")

    backup_file = os.path.join(

        BACKUP_FOLDER,

        f"parts_result_final_{timestamp}.xlsx"

    )

    shutil.copy2(

        OUTPUT_FILE,

        backup_file

    )

    backup_created = True

    print()

    print("=" * 60)

    print("BACKUP CREATED")

    print(backup_file)

    print("=" * 60)


# ==========================================================
# SAVE
# ==========================================================

def save_results():

    create_backup()

    df = pd.DataFrame(results)

    df.to_excel(

        OUTPUT_FILE,

        index=False

    )

    print()

    print("Results Saved")


# ==========================================================
# COOKIES
# ==========================================================

def close_cookies(page):

    try:

        reject_btn = page.locator(

            "text=Reject all cookies"

        )

        if reject_btn.count() > 0:

            reject_btn.first.click()

            page.wait_for_timeout(1500)

            print("Cookies Closed")

    except:

        pass
# ==========================================================
# LOAD EXISTING RESULTS
# ==========================================================

def load_existing_results():

    global results

    if os.path.exists(OUTPUT_FILE):

        try:

            df = pd.read_excel(OUTPUT_FILE)

            results = df.to_dict("records")

            print()

            print(f"Loaded {len(results)} existing rows.")

        except:

            results = []

    else:

        results = []


# ==========================================================
# READ INPUT
# ==========================================================

def load_input():

    print()

    print("Reading Input Excel...")

    return pd.read_excel(INPUT_FILE)


# ==========================================================
# MENU
# ==========================================================

def show_menu():

    print()

    print("=" * 60)

    print(" OnlineCarParts Scraper V2")

    print("=" * 60)

    print()

    print("1 - Start From Beginning")

    print("2 - Resume Last Job")

    print("3 - Start From Specific PART")

    print("4 - Start From Row Number")

    print()

    while True:

        choice = input("Choose : ").strip()

        if choice in ["1", "2", "3", "4"]:

            return choice

        print("Invalid Choice")


# ==========================================================
# START POSITION
# ==========================================================

def get_start_index(df):

    choice = show_menu()

    if choice == "1":

        print()

        print("Starting From Beginning")

        return 0


    if choice == "2":

        load_existing_results()

        if len(results) == 0:

            print("No previous results found.")

            return 0

        last_part = results[-1]["INPUT_PART"]

        print()

        print("Last Part :", last_part)

        matches = df.index[df["PART"].astype(str).str.strip() == str(last_part)]

        if len(matches) == 0:

            print("Last Part not found in Input.")

            return 0

        return matches[0] + 1


    if choice == "3":

        part = input("Enter Part : ").strip()

        matches = df.index[df["PART"].astype(str).str.strip() == part]

        if len(matches) == 0:

            print("Part not found.")

            return 0

        return matches[0]


    if choice == "4":

        row = int(input("Row Number : "))

        if row < 1:

            row = 1

        return row - 1
# ==========================================================
# REMOVE OLD RESULTS
# ==========================================================

def remove_old_results(part):

    global results

    before = len(results)

    results = [

        r

        for r in results

        if str(r["INPUT_PART"]).strip() != str(part).strip()

    ]

    removed = before - len(results)

    if removed > 0:

        print()

        print(f"Removed {removed} old rows.")


# ==========================================================
# ADD RESULT
# ==========================================================

def add_result(

        input_part,

        brand,

        article,

        result,

        link

):

    global results

    record = {

        "INPUT_PART": input_part,

        "SEARCH_BRAND": brand,

        "ARTICLE_NUMBER": article,

        "RESULT": result,

        "PRODUCT_LINK": link

    }

    duplicate = False

    for r in results:

        if (

            str(r["INPUT_PART"]) == str(record["INPUT_PART"])

            and

            str(r["ARTICLE_NUMBER"]) == str(record["ARTICLE_NUMBER"])

            and

            str(r["RESULT"]) == str(record["RESULT"])

        ):

            duplicate = True

            break

    if not duplicate:

        results.append(record)


# ==========================================================
# LOG
# ==========================================================

def print_status(

        row,

        total,

        part,

        brand

):

    elapsed = time.time() - start_time

    avg = elapsed / max(row, 1)

    remaining = avg * (total - row)

    print()

    print("=" * 60)

    print(f"Row            : {row}/{total}")

    print(f"Part           : {part}")

    print(f"Brand          : {brand}")

    print(f"Elapsed        : {elapsed/60:.1f} min")

    print(f"Remaining      : {remaining/60:.1f} min")

    print("=" * 60)
# ==========================================================
# MAIN
# ==========================================================

print()

print("=" * 60)

print("OnlineCarParts Scraper V2")

print("=" * 60)

load_existing_results()

df = load_input()

start_index = get_start_index(df)

total_rows = len(df)

with sync_playwright() as p:

    browser = p.chromium.launch(

        channel="chrome",

        headless=False,

        args=[

            "--disable-blink-features=AutomationControlled"

        ]

    )

    context = browser.new_context(

        viewport={

            "width": 1400,

            "height": 900

        }

    )

    page = context.new_page()

    print()

    print("Opening Website...")

    page.goto(

        "https://www.onlinecarparts.co.uk",

        wait_until="domcontentloaded"

    )

    page.wait_for_timeout(5000)

    close_cookies(page)

    print("Website Ready")

    for index in range(start_index, total_rows):

        row = df.iloc[index]

        part = str(

            row["PART"]

        ).strip()

        brand = normalize_brand(

            row["BRAND"]

        )

        part = re.sub(

            r"\s+",

            "",

            part

        )

        print_status(

            index + 1,

            total_rows,

            part,

            brand

        )

        remove_old_results(

            part

        )

        for retry in range(MAX_RETRIES):

            try:            
                print(f"Attempt {retry + 1}/{MAX_RETRIES}")

                # =====================================
                # OPEN SEARCH PAGE
                # =====================================

                search_url = (
                    "https://www.onlinecarparts.co.uk/"
                    f"spares-search.html?keyword={part}"
                )

                print()

                print("Opening Search")

                print(search_url)

                page.goto(

                    search_url,

                    wait_until="domcontentloaded",

                    timeout=60000

                )

                page.wait_for_timeout(5000)

                close_cookies(page)

                # =====================================
                # GET BRAND ID
                # =====================================

                brand_id = None

                brands = page.locator(
                    "input[name='brand[]']"
                )

                brands_count = brands.count()

                print()

                print(f"Brands Found : {brands_count}")

                for i in range(brands_count):

                    checkbox = brands.nth(i)

                    value = checkbox.get_attribute("value")

                    checkbox_id = checkbox.get_attribute("id")

                    if not checkbox_id:
                        continue

                    label = page.locator(
                        f"label[for='{checkbox_id}'] img"
                    )

                    brand_alt = ""

                    try:

                        brand_alt = (
                            label.get_attribute("alt")
                            .upper()
                            .strip()
                        )

                    except:
                        pass

                    print(f"Brand : {brand_alt}")

                    if brand.upper() in brand_alt:

                        brand_id = value

                        print()

                        print("Brand Matched")

                        print(f"Brand ID : {brand_id}")

                        break
                # =====================================
                # BRAND NOT FOUND
                # =====================================

                if not brand_id:

                    print("Brand Not Found")

                    add_result(

                        part,

                        brand,

                        "",

                        "BRAND NOT FOUND",

                        ""

                    )

                    save_results()

                    break

                # =====================================
                # FILTERED SEARCH
                # =====================================

                filtered_url = (

                    "https://www.onlinecarparts.co.uk/"

                    f"spares-search.html?keyword={part}"

                    f"&brand%5B%5D={brand_id}"

                )

                print()

                print("Opening Filtered Search")

                print(filtered_url)

                page.goto(

                    filtered_url,

                    wait_until="domcontentloaded",

                    timeout=60000

                )

                page.wait_for_timeout(5000)

                close_cookies(page)

                # =====================================
                # WAIT PRODUCTS
                # =====================================

                page.wait_for_selector(

                    "div.product-card",

                    timeout=60000

                )

                products = page.locator(

                    "div.product-card"

                )

                product_count = products.count()

                print()

                print(f"Products Found : {product_count}")

                product_found = False

                # =====================================
                # LOOP PRODUCTS
                # =====================================

                for x in range(product_count):

                    product = products.nth(x)

                    text = product.inner_text().upper()

                    html = product.inner_html().upper()

                    combined = text + " " + html

                    clean_combined = re.sub(

                        r"[^A-Z0-9]",

                        "",

                        combined

                    )

                    clean_part = re.sub(

                        r"[^A-Z0-9]",

                        "",

                        part.upper()

                    )

                    if clean_part not in clean_combined:

                        continue
                    # =====================================
                    # OPEN PRODUCT
                    # =====================================

                    link = product.locator(
                        "a.product-card__title-link"
                    ).get_attribute("href")

                    if not link:

                        continue

                    if not link.startswith("http"):

                        link = (
                            "https://www.onlinecarparts.co.uk"
                            + link
                        )

                    print()

                    print("Opening Product")

                    print(link)

                    page.goto(

                        link,

                        wait_until="domcontentloaded",

                        timeout=60000

                    )

                    page.wait_for_timeout(5000)

                    close_cookies(page)

                    product_found = True

                    # =====================================
                    # ARTICLE NUMBER
                    # =====================================

                    article_number = ""

                    try:

                        body = page.inner_text(
                            "body"
                        )

                        m = re.search(

                            r"Article number:\s*([^\n]+)",

                            body

                        )

                        if m:

                            article_number = (

                                m.group(1)

                                .split("Manufacturer:")[0]

                                .strip()

                            )

                            print()

                            print("ARTICLE")

                            print(article_number)

                    except:

                        pass

                    # =====================================
                    # OEM EXTRACTION
                    # =====================================

                    extracted = []

                    try:

                        oems = page.locator(
                            "a.product-oem__link"
                        )

                        oem_count = oems.count()

                        print()
                        print(f"OEM Found : {oem_count}")

                        for z in range(oem_count):

                            try:
                                oem = (
                                    oems.nth(z)
                                    .text_content()
                                    .strip()
                                )

                                if oem and oem not in extracted:
                                    extracted.append(oem)

                            except:
                                pass

                    except:
                        pass
                    # =====================================
                    # SAVE RESULTS
                    # =====================================

                    if extracted:

                        for item in extracted:

                            print()

                            print("FOUND :", item)

                            add_result(

                                part,

                                brand,

                                article_number,

                                item,

                                link

                            )

                    else:

                        add_result(

                            part,

                            brand,

                            article_number,

                            "OEM NOT FOUND",

                            link

                        )

                    break

                # =====================================
                # PRODUCT NOT FOUND
                # =====================================

                if not product_found:

                    add_result(

                        part,

                        brand,

                        "",

                        "PRODUCT NOT FOUND",

                        ""

                    )

                # =====================================
                # SAVE PROGRESS
                # =====================================

                save_results()

                print()

                print("Progress Saved")

                break

            except Exception as e:

                print()

                print("ERROR :", e)

                if retry == MAX_RETRIES - 1:

                    add_result(

                        part,

                        brand,

                        "",

                        f"ERROR : {e}",

                        ""

                    )

                    save_results()

                else:

                    print("Retrying...")

                    page.wait_for_timeout(5000)

    browser.close()

print()

print("=" * 60)

print("DONE")

print("=" * 60)

save_results()

print()

print("Output File")

print(OUTPUT_FILE)
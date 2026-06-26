import pandas as pd
import re
from playwright.sync_api import sync_playwright

# =========================================
# FILES
# =========================================

INPUT_FILE = r"C:\Users\Handler-Sm\OneDrive\Desktop\parts.xlsx"

OUTPUT_FILE = r"C:\Users\Handler-Sm\OneDrive\Desktop\parts_result_final.xlsx"

MAX_RETRIES = 3

# =========================================
# BRAND NORMALIZER
# =========================================

def normalize_brand(brand):

    brand = str(brand).upper().strip()

    replacements = {

        "WIX": "WIX FILTERS",
        "MANN": "MANN-FILTER",
        "MANN FILTER": "MANN-FILTER"
    }

    return replacements.get(brand, brand)

# =========================================
# CLOSE COOKIES
# =========================================

def close_cookies(page):

    try:

        reject_btn = page.locator(
            "text=Reject all cookies"
        )

        if reject_btn.count() > 0:

            reject_btn.first.click()

            page.wait_for_timeout(1500)

            print("Cookies closed")

    except:
        pass

# =========================================
# READ EXCEL
# =========================================

print("Reading Excel...")

df = pd.read_excel(INPUT_FILE)

results = []

# =========================================
# PLAYWRIGHT
# =========================================

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

    # =====================================
    # OPEN WEBSITE
    # =====================================

    print("Opening website...")

    page.goto(

        "https://www.onlinecarparts.co.uk",

        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(5000)

    close_cookies(page)

    print("Website opened")

    # =====================================
    # LOOP EXCEL
    # =====================================

    for index, row in df.iterrows():

        part = str(row["PART"]).strip()

        brand = normalize_brand(
            row["BRAND"]
        )

        # remove spaces
        part = re.sub(
            r"\s+",
            "",
            part
        )

        print("\n" + "=" * 60)

        print("PART:", part)

        print("BRAND:", brand)

        for retry in range(MAX_RETRIES):

            try:

                print(
                    f"Attempt {retry + 1}"
                )

                # =====================================
                # OPEN SEARCH PAGE
                # =====================================

                search_url = (
                    "https://www.onlinecarparts.co.uk/"
                    f"spares-search.html?keyword={part}"
                )

                print(
                    "Opening search:",
                    search_url
                )

                page.goto(

                    search_url,

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

                print(
                    "Products found:",
                    product_count
                )

                product_found = False

                # =====================================
                # LOOP PRODUCTS
                # =====================================

                for x in range(product_count):

                    product = products.nth(x)

                    text = (
                        product.inner_text().upper()
                    )

                    html = (
                        product.inner_html().upper()
                    )

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

                    # exact part match
                    if clean_part not in clean_combined:
                        continue

                    # brand match
                    if brand.upper() not in combined:
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

                    print(
                        "Opening product:",
                        link
                    )

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

                            print(
                                "ARTICLE:",
                                article_number
                            )

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

                        print(
                            "OEM FOUND:",
                            oem_count
                        )

                        for z in range(oem_count):

                            try:

                                oem = (
                                    oems.nth(z)
                                    .text_content()
                                    .strip()
                                )

                                if oem and oem not in extracted:

                                    extracted.append(
                                        oem
                                    )

                            except:
                                pass

                    except:
                        pass

                    # =====================================
                    # SAVE
                    # =====================================

                    if extracted:

                        for item in extracted:

                            print(
                                "FOUND:",
                                item
                            )

                            results.append({

                                "INPUT_PART":
                                    part,

                                "SEARCH_BRAND":
                                    brand,

                                "ARTICLE_NUMBER":
                                    article_number,

                                "RESULT":
                                    item,

                                "PRODUCT_LINK":
                                    link
                            })

                    else:

                        results.append({

                            "INPUT_PART":
                                part,

                            "SEARCH_BRAND":
                                brand,

                            "ARTICLE_NUMBER":
                                article_number,

                            "RESULT":
                                "OEM NOT FOUND",

                            "PRODUCT_LINK":
                                link
                        })

                    break

                # =====================================
                # PRODUCT NOT FOUND
                # =====================================

                if not product_found:

                    results.append({

                        "INPUT_PART":
                            part,

                        "SEARCH_BRAND":
                            brand,

                        "ARTICLE_NUMBER":
                            "",

                        "RESULT":
                            "PRODUCT NOT FOUND",

                        "PRODUCT_LINK":
                            ""
                    })

                # =====================================
                # SAVE PROGRESS
                # =====================================

                temp_df = pd.DataFrame(results)

                temp_df.to_excel(

                    OUTPUT_FILE,

                    index=False
                )

                print("Saved progress")

                break

            except Exception as e:

                print("ERROR:", e)

                if retry == MAX_RETRIES - 1:

                    results.append({

                        "INPUT_PART":
                            part,

                        "SEARCH_BRAND":
                            brand,

                        "ARTICLE_NUMBER":
                            "",

                        "RESULT":
                            f"ERROR: {e}",

                        "PRODUCT_LINK":
                            ""
                    })

                else:

                    print("Retrying...")

                    try:
                        page.wait_for_timeout(5000)
                    except:
                        pass

    # =====================================
    # FINAL SAVE
    # =====================================

    final_df = pd.DataFrame(results)

    final_df.to_excel(

        OUTPUT_FILE,

        index=False
    )

    print("\nDONE!")

    print("Saved:", OUTPUT_FILE)

    browser.close()
from playwright.sync_api import sync_playwright
import time
from urllib.parse import unquote
from Industry_link import options_list
def get_company_urls_from_industry_page(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        print(f"\n Opening Industry Page: {url}")

        page.goto(url, timeout=60000)
        page.wait_for_selector("table thead tr td.tableheading", timeout=15000)
        time.sleep(2)

        company_urls = page.eval_on_selector_all(
            "td:nth-child(2) a",
            "elements => elements.map(el => el.href)"
        )

        browser.close()
        return company_urls

if __name__ == "__main__":
    # 📝 List of multiple industry URLs
    industry_page_urls = options_list

    for industry_page_url in industry_page_urls:
        try:
            parts = industry_page_url.split('scripname=')

            if len(parts) > 1:
                industry_raw = parts[1]

                if '&' in industry_raw:
                    industry_raw = industry_raw.replace('&', '%26', 1)

                industry_encoded = industry_raw.split('&')[0]
                industry_name = unquote(industry_encoded).strip().replace("+"," ")

                print(f"Full Industry Name: {industry_name}")
            else:
                print("scripname not found.")
        except Exception as e:
            print(f"Error: {e}")

        company_urls = get_company_urls_from_industry_page(industry_page_url)
        print(company_urls)

        print("\n Extracted Company URLs:")
        # for link in company_urls:
        #     print(len(link))

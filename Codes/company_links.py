from playwright.sync_api import sync_playwright
import time
from urllib.parse import unquote

def get_company_urls_from_industry_page(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)  # Set to True if you don't want to open the browser window
        page = browser.new_page()
        print(f"\n Opening Industry Page: {url}")

        # Navigate to the page
        page.goto(url, timeout=60000)

        # Wait for the table heading to ensure table is loaded
        page.wait_for_selector("table thead tr td.tableheading", timeout=15000)

        # Optional: Give time for Angular to render rows
        time.sleep(2)

        # Extract all company URLs from the 2nd column (Security Name)
        company_urls = page.eval_on_selector_all(
            "td:nth-child(2) a",
            "elements => elements.map(el => el.href)"
        )

        browser.close()
        return company_urls


# Example Industry URL
industry_page_url = "https://www.bseindia.com/markets/Equity/EQReports/IndustryView.html?expandable=2&page=IN010302001&scripname=Aluminium"

try:
    # Split on 'scripname='
    parts = industry_page_url.split('scripname=')

    if len(parts) > 1:
        industry_raw = parts[1]

        # Problem: raw & inside value breaks it
        # So temporarily replace first & with %26
        if '&' in industry_raw:
            industry_raw = industry_raw.replace('&', '%26', 1)

        # Now split again safely
        industry_encoded = industry_raw.split('&')[0]

        # Decode %20, %26 etc.
        industry_name = unquote(industry_encoded).strip()

        print(f"Full Industry Name: {industry_name}")
    else:
        print("scripname not found.")
except Exception as e:
    print(f"Error: {e}")
# Get company URLs from the industry page
company_urls = get_company_urls_from_industry_page(industry_page_url)
print(company_urls)

# Print extracted URLs
print("\n Extracted Company URLs:")
if __name__ == "__main__":

    for link in company_urls:
        print(link)


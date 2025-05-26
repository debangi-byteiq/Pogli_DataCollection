import os
import time
import warnings
from playwright.sync_api import sync_playwright
from company_links import company_urls
# Suppress all warnings
warnings.filterwarnings("ignore")

def sanitize_filename(name):
    """Sanitize a string to make it a valid file name."""
    return "".join(c if c.isalnum() or c in " _-" else "_" for c in name)

def extract_company_name_from_url(url):
    """Extract and format company name from BSE URL."""
    try:
        parts = url.strip('/').split('/')
        if len(parts) >= 5:
            return sanitize_filename(parts[4].replace('-', ' ').title())
    except Exception:
        pass
    return "Unknown_Company"

def download_latest_annual_report(page, company_name):
    try:
        page.click('a#afi')
        time.sleep(1)
        page.click('a#l62')
        page.wait_for_selector('table[ng-if="loader.ARState==\'loaded\'"]', timeout=10000)

        rows = page.query_selector_all('table[ng-if="loader.ARState==\'loaded\'"] tbody tr')
        if rows:
            link = rows[0].query_selector('td:nth-child(6) a')
            if link:
                with page.expect_download() as download_info:
                    link.click()
                download = download_info.value

                os.makedirs("annual_reports", exist_ok=True)
                file_path = os.path.join("annual_reports", f"{company_name}.pdf")
                download.save_as(file_path)
                print(f"Annual Report downloaded: {file_path}")
            else:
                print(f" No PDF link for Annual Report: {company_name}")
        else:
            print(f" No rows found in Annual Report table: {company_name}")
    except Exception as e:
        print(f" Error downloading Annual Report for {company_name}: {e}")

def download_brsr_report(page, company_name):
    try:
        page.click('div#l103')
        page.wait_for_selector('table[ng-if="loader.SHState == \'loaded\'"]', timeout=10000)

        rows = page.query_selector_all('table[ng-if="loader.SHState == \'loaded\'"] tbody tr')
        for row in rows:
            pdf_link_element = row.query_selector('td:last-child a')
            if pdf_link_element:
                pdf_url = pdf_link_element.get_attribute("href")
                if not pdf_url.startswith("http"):
                    pdf_url = "https://www.bseindia.com" + pdf_url

                with page.expect_download() as download_info:
                    pdf_link_element.click()
                download = download_info.value

                os.makedirs("brsr_reports", exist_ok=True)
                file_path = os.path.join("brsr_reports", f"{company_name}.pdf")
                download.save_as(file_path)
                print(f" BRSR Report downloaded: {file_path}")
                return
        print(f"No downloadable BRSR PDF found for {company_name}")
    except Exception as e:
        print(f" Error downloading BRSR for {company_name}: {e}")

def download_reports_for_company(url):
    company_name = extract_company_name_from_url(url)
    print(f"\n Processing: {company_name}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        try:
            page.goto(url, timeout=60000)
            download_latest_annual_report(page, company_name)
            download_brsr_report(page, company_name)
        except Exception as e:
            print(f" Error navigating {company_name}: {e}")
        finally:
            browser.close()

 # List of Company URLs
company_url = company_urls
print(len(company_url))

#  Run for each company
for url in company_url:
    download_reports_for_company(url)

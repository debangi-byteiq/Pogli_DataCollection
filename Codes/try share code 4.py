import re
import pandas as pd
import os
from playwright.sync_api import sync_playwright
import time
from urllib.parse import unquote
from Industry_link import options_list

def get_company_urls_from_industry_page(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        print(f"\nOpening Industry Page: {url}")

        page.goto(url, timeout=60000)
        page.wait_for_selector("table thead tr td.tableheading", timeout=15000)
        time.sleep(2)

        company_urls = page.eval_on_selector_all(
            "td:nth-child(2) a",
            "elements => elements.map(el => el.href)"
        )

        browser.close()
        return company_urls

def extract_industry_name(industry_url):
    try:
        parts = industry_url.split('scripname=')
        if len(parts) > 1:
            industry_raw = parts[1]
            if '&' in industry_raw:
                industry_raw = industry_raw.replace('&', '%26', 1)
            industry_encoded = industry_raw.split('&')[0]
            industry_name = unquote(industry_encoded).strip().replace("+"," ")

            return industry_name
        else:
            return "Unknown"
    except Exception as e:
        print(f"Error extracting industry name from {industry_url}: {e}")
        return "Unknown"

def extract_company_info(url):
    try:
        parts = url.strip('/').split('/')
        company_raw = parts[4]
        script_code = parts[6]

        company_name = re.sub(r'-+', ' ', company_raw).title()
        return company_name, script_code
    except Exception as e:
        print(f"Error processing URL {url}: {e}")
        return None, None

industry_page_urls =options_list
output_filename = "company_share1.xlsx"

all_company_data = []

for industry_page_url in industry_page_urls:
    industry_name = extract_industry_name(industry_page_url)
    print(f"\nProcessing Industry: {industry_name}")

    company_urls = get_company_urls_from_industry_page(industry_page_url)

    for link in company_urls:
        company_name, script_code = extract_company_info(link)
        if company_name and script_code:
            all_company_data.append({
                'Industry': industry_name,
                'Company Name': company_name,
                'Script Code': script_code,
                'URL': link
            })
            print(f"{company_name} | {script_code}|{link}")

# Create DataFrame for new data
new_df = pd.DataFrame(all_company_data)

# Check if Excel file exists
if os.path.exists(output_filename):
    print(f"\nExcel file '{output_filename}' exists. Appending data...")
    existing_df = pd.read_excel(output_filename)
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
else:
    print(f"\nExcel file '{output_filename}' not found. Creating new file...")
    combined_df = new_df

# Optional: Remove duplicates (based on Script Code)
combined_df.drop_duplicates(subset=['Script Code'], inplace=True)

# Save combined data back to Excel
combined_df.to_excel(output_filename, index=False)
print(f"\n Data saved to Excel file: {output_filename}")


from playwright.sync_api import sync_playwright
import pandas as pd
import re
import time


def scrape_data():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_viewport_size({"width": 1920, "height": 1080})

        # Navigate to the industry page
        page.goto(
            "https://www.bseindia.com/markets/Equity/EQReports/IndustryView.html?expandable=2&page=IN070203001&scripname=Heavy%20Electrical%20Equipment")
        time.sleep(5)  # Allow time for the page to load

        # Extract company URLs
        elements = page.locator('td.tdcolumn.text-center a.tablebluelink.ng-binding').all()
        urls = list(set([element.get_attribute('href') for element in elements]))

        print("Extracted URLs:", urls)

        data = []

        for url in urls:
            page.goto(url)
            time.sleep(5)

            try:
                company_name = \
                page.locator('//div[@class="col-lg-12 col-md-12 col-sm-12 col-xs-12 companyname"]').inner_text().split(
                    '\n')[0].strip()
                market_cap = page.locator(
                    '//td[contains(text(), "Mcap Full (Cr.)")]/following-sibling::td').inner_text()

                print("Company Name -", company_name)
                print("Market Cap -", market_cap)

                # Click on Financials > Results
                page.click("#afi")
                time.sleep(2)
                page.click("#l61")
                time.sleep(2)

                # Extract financial data
                rows = page.locator("tbody tr").all()
                financials = {}

                for row in rows:
                    columns = row.locator("td").all()
                    if len(columns) > 1:
                        key = columns[0].inner_text().strip()
                        value = columns[1].inner_text().strip()
                        financials[key] = value

                # Click on Corporate Information
                corp_info_link = page.locator("//div[@id='l13']//a[contains(@href, 'corp-information')]")
                page.evaluate("(el) => el.scrollIntoView()", corp_info_link)
                time.sleep(2)
                corp_info_link.click()
                time.sleep(2)

                website_url = page.locator(
                    "#deribody div:nth-child(6) table tbody tr:nth-child(5) td:nth-child(2) a").get_attribute("href")
                print("Website -", website_url)

                data.append({
                    'Company Name': company_name,
                    'Market Cap (Cr)': market_cap,
                    'Revenue': financials.get('Revenue', ''),
                    'Other Income': financials.get('Other Income', ''),
                    'Total Income': financials.get('Total Income', ''),
                    'Expenditure': financials.get('Expenditure', ''),
                    'Operating Profit': financials.get('PBDT', ''),
                    'Interest': financials.get('Interest', ''),
                    'Depreciation': financials.get('Depreciation', ''),
                    'PBT': financials.get('PBT', ''),
                    'Tax': financials.get('Tax', ''),
                    'Net Profit': financials.get('Net Profit', ''),
                    'Equity': financials.get('Equity', ''),
                    'EPS': financials.get('EPS', ''),
                    'CEPS': financials.get('CEPS', ''),
                    'OPM %': financials.get('OPM %', ''),
                    'NPM %': financials.get('NPM %', ''),
                    "Website": website_url
                })
            except Exception as e:
                print(f"Error occurred for URL: {url}, Error: {str(e)}")
                continue

        # Convert data to DataFrame
        df = pd.DataFrame(data)
        df.insert(0, 'Industry', 'Heavy Electrical Equipment  ')

        # Function to clean market cap
        def clean_market_cap(market_cap_text):
            if re.match(r'^[\d,.]+$', market_cap_text.strip()):
                market_cap_text = market_cap_text.replace(',', '').strip()
                return float(market_cap_text) if market_cap_text != '-' else 0.0
            return None

        df['Market Cap (Cr)'] = df['Market Cap (Cr)'].apply(clean_market_cap)
        df = df[(df['Market Cap (Cr)'] >= 5000) & (df['Market Cap (Cr)'] <= 20000)]

        df.to_excel("Havey.xlsx", index=False)
        print("Data extraction and filtering completed. Results saved to 'iron.xlsx'.")

        browser.close()


scrape_data()

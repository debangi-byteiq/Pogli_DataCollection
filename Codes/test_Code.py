from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import warnings
import pandas as pd
import os


def get_dividends_data(page):
    print("Scraping Dividends data...")
    try:
        # Click on 'Peer Group' tab
        page.locator('div#l14').click()
        time.sleep(1)

        # Click on 'Bonus & Dividends' tab
        page.get_by_role("link", name="Bonus & Dividends").click()
        time.sleep(2)  # Wait for the content to load

        # Extract table data
        table_html = page.query_selector('div.tab-pane.active.largetable').inner_html()
        soup = BeautifulSoup(table_html, 'html.parser')

        # Find the table containing "Dividend"
        dividend_table = None
        for table in soup.find_all('table'):
            if "Dividend" in table.text:
                dividend_table = table
                break

        if not dividend_table:
            print("Dividend table not found")
            return pd.DataFrame()

        # Extract headers (only first 2 columns)
        headers = [header.text.strip() for header in dividend_table.find_all('td', class_='tableheading')][:2]

        # Extract rows (only first 2 columns)
        rows = []
        for row in dividend_table.find_all('tr'):
            cells = [cell.text.strip() for cell in row.find_all('td', class_='tdcolumn')][:2]
            if cells:
                rows.append(cells)

        # Ensure rows match the headers in length
        consistent_rows = [row for row in rows if len(row) == len(headers)]

        # Get only the last 5 rows
        df_dividends = pd.DataFrame(consistent_rows, columns=headers).tail(5)

        # Rename columns
        df_dividends.rename(columns={df_dividends.columns[0]: "Date", df_dividends.columns[1]: "Dividend Value"},
                            inplace=True)

        df_dividends.reset_index(drop=True, inplace=True)

        return df_dividends

    except Exception as e:
        print(f"Error in get_dividends_data: {e}")
        return pd.DataFrame()


def main(url):
    l = url.split('/')
    company = l[4].strip().capitalize()
    warnings.filterwarnings("ignore")

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True)
        page = browser.new_page()
        page.goto(url, timeout=60000)
        time.sleep(2)
        print(f"Navigating to the page {url}")
        print("This might take a while...")

        # Define the Excel file path
        excel_path = "../ExcelFiles/Dividends_Data3.xlsx"

        try:
            # Scrape Dividend data only (first 2 columns, last 5 rows)
            df_dividends = get_dividends_data(page)
            df_dividends['Company Name'] = company
            df_dividends['Industry Name'] = "Iron & Steel Products"

            # Check if the file exists
            if os.path.exists(excel_path):
                with pd.ExcelFile(excel_path) as xls:
                    df_dividends_existing = pd.read_excel(xls, sheet_name='Dividends')
                df_dividends = pd.concat([df_dividends_existing, df_dividends], ignore_index=True)
            else:
                print("Creating a new Excel file...")

            # Save data to Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                df_dividends.to_excel(writer, sheet_name='Dividends', index=False)
                print(f"Dividends data scraped successfully for {company}")
        except Exception as e:
            print(e)
            print("Failed to scrape data")


if __name__ == "__main__":
    urls = ['https://www.bseindia.com/stock-share-price/jindal-saw-ltd/jindalsaw/500378/']  # Add URLs here
    for url in urls:
        main(url)
        time.sleep(5)

from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import warnings
import pandas as pd
import os
from company_links import company_urls

def get_peer_data(page):
    print("Scraping peer group data")
    try:
        page.locator('div#l14').click()
        time.sleep(1)
        page.get_by_role("link", name="Annual Trends").click()
        time.sleep(2)
        # soup = BeautifulSoup(page.content(), 'html.parser')
        table_html = page.query_selector('div.tab-pane.active.largetable').inner_html()
        soup = BeautifulSoup(table_html, 'html.parser')
        headers = [header.text.strip() for header in soup.find_all('td', class_='tableheading')]

        rows = []
        for row in soup.find_all('tr'):
            cells = [cell.text.strip() for cell in row.find_all('td', class_='tdcolumn')]
            if cells:
                rows.append(cells)

        consistent_rows = [row for row in rows if len(row) == len(headers)]
        df_peer = pd.DataFrame(consistent_rows, columns=headers)

        # Transpose the dataframe to make it more manageable
        df_peer_transposed = df_peer.T
        new_header = df_peer_transposed.iloc[0]
        df_peer_transposed = df_peer_transposed[1:]
        df_peer_transposed.columns = new_header

        # Rename the 'Results (in Cr.)' column
        df_peer_transposed.rename(columns={'Results (in Cr.)  View in (Million)': 'Year'}, inplace=True)

        # Split the '52 W H/L' column into '52 W H' and '52 W L'
        if '52 W H/L' in df_peer_transposed.columns:
            df_peer_transposed[['52 W H', '52 W L']] = df_peer_transposed['52 W H/L'].str.split('/', expand=True)
            df_peer_transposed.drop(columns=['52 W H/L'], inplace=True)

        df_peer_transposed.reset_index(inplace=True)
        df_peer_transposed.rename(columns={'index': 'Peer Company'}, inplace=True)

        # Filter the dataframe to include only the required columns
        required_columns = [
            'Peer Company', 'LTP', 'Change %', 'Year', 'Sales', 'PAT', 'Equity',
            'Face Value', 'OPM %', 'NPM %', 'EPS', 'CEPS', 'PE',
            '52 W H', '52 W L'
        ]

        df_peer_filtered = df_peer_transposed[required_columns]

        return df_peer_filtered
    except Exception as e:
        print(f"Error in get_peer_data: {e}")
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
        excel_path = "../ExcelFiles/Peer_Company-try.xlsx"

        try:
            # Scrape only peer company data
            df_peers = get_peer_data(page)
            df_peers['Company Name'] = company
            df_peers['Industry Name'] = "Iron & Steel Products"

            # Check if the file exists
            if os.path.exists(excel_path):
                with pd.ExcelFile(excel_path) as xls:
                    df_peer_existing = pd.read_excel(xls, sheet_name='Peer_Annual Data')
                df_peers = pd.concat([df_peer_existing, df_peers], ignore_index=True)
            else:
                print("Could not find the excel file")

            # Save only peer data to an Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                df_peers.to_excel(writer, sheet_name='Peer_Annual Data', index=False)
                print(f"Peer data scraped successfully for {company}")
        except Exception as e:
            print(e)
            print("Failed to scrape data")

if __name__ == "__main__":
    urls = company_urls  # Add URLs here
    for url in urls:
        main(url)
        time.sleep(5)

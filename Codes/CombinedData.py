from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import warnings
import pandas as pd
import os


def get_equity_data(page):
    soup = BeautifulSoup(page.content(), 'html.parser')
    dict1 = dict()
    print("Scraping equity data")
    div_elements = soup.find_all('div', class_='col-lg-13')
    for element in div_elements:
        try:
            tr_tags = element.find('tbody').find_all('tr')
        except:
            continue
        for tr_tag in tr_tags:
            try:
                key = tr_tag.find('td', class_='textsr').text.strip()
                value = tr_tag.find('td', class_='textvalue ng-binding').text.strip()

                # Assign the key-value pair
                dict1[key] = value
            except:
                continue

    # Convert the extracted dictionary to DataFrame
    row_list = []
    data = dict1
    row_list.append(data.values())
    df_equity = pd.DataFrame(row_list, columns=data.keys())

    # Check if 'PE/PB' column exists in the DataFrame
    if 'PE/PB' in df_equity.columns:
        # Split the 'PE/PB' column into two columns 'PE' and 'PB'
        df_equity[['PE', 'PB']] = df_equity['PE/PB'].str.split(' / ', expand=True)

        # Drop the original 'PE/PB' column
        df_equity.drop(columns=['PE/PB'], inplace=True)

    return df_equity

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

def get_corpgov_data(page):
    print("Scraping corporate governance data")
    try:
        page.get_by_text('Corporate Governance').click()
        time.sleep(4)

        soup = BeautifulSoup(page.content(), 'html.parser')
        element = soup.find_all('table', class_='ng-scope')[4]
        row_list = []
        tr_body = element.find('tbody').find_all('tr')

        for tr_tag in tr_body:
            l = []
            try:
                td_tags = tr_tag.find_all('td', class_='ng-binding')
                for td_tag in td_tags:
                    l.append(td_tag.text.strip())
            except:
                continue
            row_list.append(l)

        # Original column data
        column_data = ['Sr', 'Title (Mr/Ms)', 'Name of the Director', 'DIN', 'Category',
                       'Whether the director is disqualified?', 'Start Date of disqualification',
                       'End Date of disqualification', 'Details of disqualification',
                       'Current status', 'Whether special resolution passed?',
                       'Date of passing special resolution', 'Initial Date of Appointment',
                       'Date of Re-appointment', 'Date of cessation', 'Tenure of Director (in months)',
                       'No of Directorship in listed entities',
                       'No of Independent Directorship in listed entities',
                       'Number of memberships in Audit/ Stakeholder Committee(s)',
                       'No of post of Chairperson in Audit/ Stakeholder Committee',
                       'Reason for Cessation', 'Notes for not providing PAN', 'Notes for not providing DIN']

        dataframe = pd.DataFrame(row_list, columns=column_data)

        # Combine 'Title (Mr/Ms)' and 'Name of the Director' into one column
        dataframe['Director Name'] = dataframe['Title (Mr/Ms)'] + " " + dataframe['Name of the Director']

        # Dropping original columns after combining
        dataframe.drop(columns=['Title (Mr/Ms)', 'Name of the Director'], inplace=True)

        # Create a filtered DataFrame with only the needed columns
        needed_row_list = []
        for lis in row_list:
            combined_name = lis[1] + " " + lis[2]  # Combining title and name for filtered DataFrame
            row_list1 = [lis[0], combined_name, lis[4], lis[9], lis[12], lis[13]]
            needed_row_list.append(row_list1)

        # Define the new column names for the filtered DataFrame
        column_data1 = ['Sr', 'Director Name', 'Category', 'Current status',
                        'Initial Date of Appointment', 'Date of Re-appointment']

        dataframe1 = pd.DataFrame(needed_row_list, columns=column_data1)

    except Exception as e:
        print(f"Error in get_corpgov_data: {e}")
        dataframe = pd.DataFrame()
        dataframe1 = pd.DataFrame()
    return dataframe, dataframe1


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
        excel_path = "../ExcelFiles/Combined_data_BSE.xlsx"

        try:
            #Scrape all the data
            df_equity = get_equity_data(page)
            df_corpgov, df_corpgov_filtered = get_corpgov_data(page)
            df_peers = get_peer_data(page)
            for df in [df_equity, df_corpgov, df_corpgov_filtered, df_peers]:
                df['Company Name'] = company 
                df['Industry Name'] = "Auto Components & Equipments"

            # Check if the file exists
            if os.path.exists(excel_path):
                # Load existing data
                with pd.ExcelFile(excel_path) as xls:
                    df_equity_existing = pd.read_excel(xls, sheet_name='Equity Data')
                    df_corpgov_existing = pd.read_excel(xls, sheet_name='Corporate Governance')
                    df_corpgov_filtered_existing = pd.read_excel(xls, sheet_name='Board_Members')
                    df_peer_existing = pd.read_excel(xls, sheet_name='Peer_Annual Data')

                # Append new data to existing data
                df_equity = pd.concat([df_equity_existing, df_equity], ignore_index=True)
                df_corpgov = pd.concat([df_corpgov_existing, df_corpgov], ignore_index=True)
                df_corpgov_filtered = pd.concat([df_corpgov_filtered_existing, df_corpgov_filtered], ignore_index=True)
                df_peers = pd.concat([df_peer_existing, df_peers], ignore_index=True)
            else:
                print("Could not find the excel file")

            # Save all DataFrames to different sheets in an Excel file
            with pd.ExcelWriter(excel_path, engine='openpyxl', mode='w') as writer:
                df_equity.to_excel(writer, sheet_name='Equity Data', index=False)
                df_corpgov.to_excel(writer, sheet_name='Corporate Governance', index=False)
                df_corpgov_filtered.to_excel(writer, sheet_name='Board_Members', index=False)
                df_peers.to_excel(writer, sheet_name='Peer_Annual Data', index=False)
                print(f"Data scraped successfully for {company}")
        except Exception as e:
            print(e)
            print("Failed to scrape data")


if __name__ == "__main__":
    urls = [
        'https://www.bseindia.com/stock-share-price/shivam-autotech-ltd/shivamauto/532776/',
        'https://www.bseindia.com/stock-share-price/fiem-industries-ltd/fiemind/532768/',
        'https://www.bseindia.com/stock-share-price/india-nippon-electricals-ltd/indnippon/532240/',
        'https://www.bseindia.com/stock-share-price/lumax-auto-technologies-ltd/lumaxtech/532796/',
        'https://www.bseindia.com/stock-share-price/sandhar-technologies-ltd/sandhar/541163/',
        'https://www.bseindia.com/stock-share-price/nrb-bearings-ltd/nrbbearing/530367/',
        'https://www.bseindia.com/stock-share-price/mmforgings-ltd/mmfl/522241/',
        'https://www.bseindia.com/stock-share-price/steel-strips-wheels-ltd/sswl/513262/',
        'https://www.bseindia.com/stock-share-price/igarashi-motors-india-ltd/igarashi/517380/',
        'https://www.bseindia.com/stock-share-price/jay-bharat-maruti-ltd/jaybarmaru/520066/',
        'https://www.bseindia.com/stock-share-price/g-n-a-axles-ltd/gna/540124/',
        'https://www.bseindia.com/stock-share-price/talbros-automotive-components-ltd/talbroauto/505160/',
        'https://www.bseindia.com/stock-share-price/lumax-industries-ltd/lumaxind/517206/',
        'https://www.bseindia.com/stock-share-price/wheels-india-ltd/wheels/590073/',
        'https://www.bseindia.com/stock-share-price/bharat-seats-ltd/bharatse/523229/',
        'https://www.bseindia.com/stock-share-price/alicon-castalloy-limited/alicon/531147/',
    ]
    for url in urls:
        main(url)
        time.sleep(5)

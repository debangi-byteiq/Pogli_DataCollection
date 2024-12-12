from playwright.sync_api import sync_playwright
import time
from bs4 import BeautifulSoup
import warnings
import pandas as pd
from datetime import datetime
import os


# To get the present date 
def get_today_date():
    today = datetime.now()
    # Return the date formatted as DD/MM/YYYY
    return f"{today.day}/{today.month:02}/{today.year}"


def get_financial_data(page, company, industry, excel_path):
    page.query_selector('a#afi').click()
    time.sleep(2)

    # Scrape Quarterly Financial Data
    page.locator('a#l61').click()
    time.sleep(2)
    print("Getting quarterly financial data...")
    soup = BeautifulSoup(page.content(), 'html.parser')
    try:
        tables = soup.find_all('table', class_="ng-binding")
    except:
        print("No tables found with class 'ng-binding'")
        return pd.DataFrame(), pd.DataFrame()

    # Process Quarterly trends
    financial_quarterly_df = process_table(tables[0], 'Quarterly')
    financial_quarterly_df['Company Name'] = company
    financial_quarterly_df['Industry Name'] = industry

    # Scrape Annual Financial Data
    print("Getting annual financial data...")
    page.get_by_role("link", name="Annual Trends").click()
    soup1 = BeautifulSoup(page.content(), 'html.parser')
    tables1 = soup1.find_all('table', class_="ng-binding")

    if len(tables1) < 4:
        print(f"Expected at least 4 tables, but found {len(tables1)}")
        return pd.DataFrame(), pd.DataFrame()

    # Process Annual Trends
    financial_annual_df = process_table(tables1[3], 'Annual')
    financial_annual_df['Company Name'] = company
    financial_annual_df['Industry Name'] = industry
    page.go_back()
    time.sleep(2)

    return financial_quarterly_df, financial_annual_df


def process_table(table, trend_type):
    row_list = []
    columns = []
    print("Processing table")
    try:
        tr_heads = table.find('thead').find_all('tr')
        tr_body = table.find('tbody').find_all('tr')
    except AttributeError:
        print(f"Table structure for {trend_type} Trends is not as expected")
        return pd.DataFrame(), pd.DataFrame()
    # Extract data from body rows
    for tr_tag in tr_body:
        text_list = []
        try:
            td_tags = tr_tag.find_all('td', class_='tdcolumn')
            for td_tag in td_tags:
                text_list.append(td_tag.text.strip())
        except:
            continue
        row_list.append(text_list)

    # Extract column headers
    for th_tag in tr_heads:
        td_tags = th_tag.find_all('td', class_='tableheading')
        for td_tag in td_tags:
            columns.append(td_tag.text.strip())

    dataframe = pd.DataFrame(row_list, columns=columns)
    dataframe = dataframe[:-3]  # Remove unwanted rows
    dataframe = dataframe.drop(dataframe.index[0])  # Remove header row
    if trend_type == 'Quarterly':
        if dataframe.columns[-1].startswith('FY'):
            dataframe = dataframe.iloc[:, :-1]  # Drop the last column

    # Transpose the data
    dataframe = dataframe.set_index(dataframe.columns[0]).transpose()
    dataframe.reset_index(inplace=True)
    dataframe.rename(columns={'index': 'Date' if trend_type == 'Quarterly' else 'Year'}, inplace=True)

    return dataframe


def get_meetings_data(page, company, link, industry):

    # Function to extract table data for meetings
    def extract_meeting_data(link_id, state_attr, meeting_type):
        page.locator('a#ame').click()
        time.sleep(2)
        page.locator(link_id).click()
        time.sleep(4)  # Wait for the page to load
        # Parse the page content with BeautifulSoup
        soup = BeautifulSoup(page.content(), 'html.parser')

        # Locate the table for the meeting data
        table = soup.find('table', {'ng-if': f"loader.{state_attr}=='loaded'"})
        data = []
        if table:
            rows = table.find_all('tr')[1:]  # Skip the header row
            for row in rows:
                cols = row.find_all('td')
                if len(cols) >= 2:  # Ensure there are at least 2 columns
                    meeting_date = cols[0].text.strip()
                    purpose = cols[1].text.strip()
                    data.append([meeting_date, purpose, meeting_type, company, industry])
        else:
            print(f"Table for {state_attr} not found.")
        page.go_back()
        time.sleep(3)  # Wait for the page to load again

        return data


    # Step 1: Scrape Board Meetings Data
    print("Scraping Board Meetings data")
    board_meetings_data = extract_meeting_data("a#l71", "BMState", "Board Meeting")
    board_meetings_df = pd.DataFrame(board_meetings_data,
                                     columns=['Meeting Date', 'Purpose', 'Meeting Type', 'Company Name',
                                              'Industry'])
    time.sleep(3)
    # Step 2: Scrape Shareholder Meetings Data
    print("Scraping Shareholder Meetings data")
    shareholder_meetings_data = extract_meeting_data("a#l72", "SHState", "Shareholder Meeting")
    shareholder_meetings_df = pd.DataFrame(shareholder_meetings_data,
                                           columns=['Meeting Date', 'Purpose', 'Meeting Type', 'Company Name',
                                                    'Industry'])
    # Combine both dataframes into one
    combined_meetings_df = pd.concat([board_meetings_df, shareholder_meetings_df], ignore_index=True)

    return combined_meetings_df



# def get_peer_group(page,company):
#     page.get_by_role("heading", name="Peer Group").get_by_role("link").click()
#     time.sleep(3)
#     soup = BeautifulSoup(page.content(), 'html.parser')
#     element1 = soup.find('div',id = "qtly")
#     element2 = element1.find_all('table')[0]
#     element = element2.find('table')
#     row_list= list()
#     tr_heads = element.find('thead').find_all('tr')
#     tr_body = element.find('tbody').find_all('tr')
#     columns = list()
#     for tr_tag in tr_body:
#         l = list()
#         try:
#             td_tags = tr_tag.find_all('td', class_='tdcolumn')
#             for td_tag in td_tags:
#                 l.append(td_tag.text.strip())
#         except:
#             continue
#         row_list.append(l)
#     for th_tag in tr_heads:
#         td_tags = th_tag.find_all('td', class_='tableheading')
#         for td_tag in td_tags:
#             columns.append(td_tag.text.strip())
#     dataframe1 = pd.DataFrame(row_list,columns=columns)
#     dataframe1 = dataframe1[:-3]
#     page.get_by_role('link', name="Annual Trends").click()
#     time.sleep(3)
#     soup1 = BeautifulSoup(page.content(), 'html.parser')
#     element3 = soup1.find('div',id = "ann")
#     element4 = element3.find_all('table')[0]
#     element5 = element4.find('table')
#     row_list1= list()
#     tr_heads1 = element5.find('thead').find_all('tr')
#     tr_body1 = element5.find('tbody').find_all('tr')
#     columns1 = list()
#     for tr_tag1 in tr_body1:
#         l1 = list()
#         try:
#             td_tags1 = tr_tag1.find_all('td', class_='tdcolumn')
#             for td_tag1 in td_tags1:
#                 l1.append(td_tag1.text.strip())
#         except:
#             continue
#         row_list1.append(l1)
#     for th_tag1 in tr_heads1:
#         td_tags1 = th_tag1.find_all('td', class_='tableheading')
#         for td_tag1 in td_tags1:
#             columns1.append(td_tag1.text.strip())
#     dataframe2 = pd.DataFrame(row_list1,columns=columns1)
#     dataframe2 = dataframe2[:-3]
#
#     page.get_by_role('link', name="Bonus & Dividends").click()
#     time.sleep(3)
#     soup1 = BeautifulSoup(page.content(), 'html.parser')
#     element6 = soup1.find('div',id = "bnd")
#     element7 = element6.find_all('table')[0]
#     element8 = element7.find('table')
#     row_list2= list()
#     tr_heads2 = element8.find('thead').find_all('tr')
#     tr_body2 = element8.find('tbody').find_all('tr')
#     columns2 = list()
#     for tr_tag2 in tr_body2:
#         l2 = list()
#         try:
#             td_tags2 = tr_tag2.find_all('td', class_='tdcolumn')
#             for td_tag2 in td_tags2:
#                 l2.append(td_tag2.text.strip())
#         except:
#             continue
#         row_list2.append(l2)
#     for th_tag2 in tr_heads2:
#         td_tags2 = th_tag2.find_all('td', class_='tableheading')
#         for td_tag2 in td_tags2:
#             columns2.append(td_tag2.text.strip())
#     dataframe3 = pd.DataFrame(row_list2,columns=columns2)
#     # # dataframe = pd.concat([dataframe1,dataframe2], axis=1)
#     file_path = f'{company}_peer_group.xlsx'
#
# # Use ExcelWriter to write multiple DataFrames to different sheets
#     with pd.ExcelWriter(file_path) as writer:
#         dataframe1.to_excel(writer, sheet_name='Sheet1', index=False)
#         dataframe2.to_excel(writer, sheet_name='Sheet2', index=False)
#         dataframe3.to_excel(writer, sheet_name='Sheet3', index=False)
#
#     # dataframe1.to_excel(f'{company}_financial_results_quaterly.xlsx', index=False)
#     # dataframe2.to_excel(f'{company}_financial_results_annual.xlsx', index=False)


def main(url):
    l = url.split('/')
    company = l[4].strip().capitalize()
    warnings.filterwarnings("ignore")

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=True)
        page = browser.new_page()
        print(f"Navigating to the page {url}")
        print("This might take a while...")
        page.goto(url, timeout=60000)
        time.sleep(2)

        # Specify the path where you want to save the Excel file
        excel_path = "Financial_Bse_data.xlsx"
        all_financial_quarterly_data = pd.DataFrame()
        all_financial_annual_data = pd.DataFrame()
        all_meetings_data = pd.DataFrame()
        try:
            # Get financial results with industry name
            financial_quarterly_data, financial_annual_data = get_financial_data(page, company,"Auto Components & Equipments", excel_path)
            all_financial_quarterly_data = pd.concat([all_financial_quarterly_data, financial_quarterly_data], ignore_index=True)
            all_financial_annual_data = pd.concat([all_financial_annual_data, financial_annual_data], ignore_index=True)

            # Get meeting data
            meetings_data = get_meetings_data(page, company, url, "Auto Components & Equipments")
            all_meetings_data = pd.concat([all_meetings_data, meetings_data], ignore_index=True)

            # Check if the Excel file exists
            if os.path.exists(excel_path):
                # Load existing data from Excel
                with pd.ExcelFile(excel_path) as xls:
                    # Read existing data from sheets
                    existing_quarterly_df = pd.read_excel(xls, sheet_name='Financial_Quarterly Trend')
                    existing_annual_df = pd.read_excel(xls, sheet_name='Financial_Annual Trend')
                    existing_meetings_df = pd.read_excel(xls, sheet_name='Meeting_Data')

                    # Append new data to existing data
                    all_financial_quarterly_data = pd.concat([existing_quarterly_df, all_financial_quarterly_data],
                                                             ignore_index=True)
                    all_financial_annual_data = pd.concat([existing_annual_df, all_financial_annual_data], ignore_index=True)
                    all_meetings_data = pd.concat([existing_meetings_df, all_meetings_data], ignore_index=True)
            else:
                print("Excel File could not be found")
                # Save the collected data to Excel
            with pd.ExcelWriter(excel_path) as writer:
                all_financial_quarterly_data.to_excel(writer, sheet_name='Financial_Quarterly Trend',
                                                      index=False)  # Changed sheet name
                all_financial_annual_data.to_excel(writer, sheet_name='Financial_Annual Trend',
                                                   index=False)  # Changed sheet name
                all_meetings_data.to_excel(writer, sheet_name='Meeting_Data', index=False)
            print(f"Financial Data scraped successfully for {company}")
            print(".......................................")
        except Exception as e:
            print(e)
            print('Failed to scrape Data')


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

from playwright.sync_api import sync_playwright
import re
import time
import pandas as pd

def find_tickerRating(page, link):
    page.goto(link)
    time.sleep(2)
    try:
        # Getting the html element containing the ratings text
        ticker = page.locator("span#mainContent_ltrlOverAllRating").get_attribute('aria-label')
        rating = re.search(r'(\d*)\s*out\s*of\s*5', ticker).group(1)
    except:
        rating = 'Not found'
    print("Ticker Finology Rating is " + rating)
    return rating


def find_glassdoorRating(page, link):
    page.goto(link)
    time.sleep(2)
    try:
        # Getting the html element containing the ratings text
        text = page.query_selector("p.rating-headline-average_rating__J5rIy").text_content().strip()
        rating = re.search(r'\s*(\d*\.?\d*)', text).group(1)
    except:
        rating = 'Not found'
    print("Glassdoor Rating is" + rating)
    return rating

def find_justDialRating(page, link):
    page.goto(link)
    time.sleep(2)
    try:
        # Getting the html element containing the ratings text
        text = page.query_selector_all("div.jsx-2673505135.RatingBox.mr-6")[1].query_selector("span").text_content()
        rating = re.search(r'\s*(\d*\.?\d*)', text).group(1)
    except:
        rating = 'Not found'
    print("Justdial Rating is " + rating)
    return rating

def find_crisilRating(page, link):
    page.goto(link)
    time.sleep(2)
    try:
        # Getting the html element containing the ratings text
        text = page.query_selector_all('div.comp-fs-instrument-content')[0].text_content().strip()
        rating = re.search(r'Ratings\s*CRISIL\s([A-Z0-3\+\-]+)', text).group(1).strip()
    except:
        rating = 'Not found'
    print("Crisil Rating is " + rating)
    return rating



def find_ambitionBoxRating(page, link):
    page.goto(link)
    time.sleep(2)
    try:
        # Getting the html element containing the ratings text
        text = page.query_selector_all("div.rating_text.rating_text--md ")[0].query_selector('div').text_content()
        rating = re.search(r'\s*(\d*\.?\d*)', text).group(1)
    except:
        rating = 'Not found'
    print("AmbitionBox Rating is " + rating)
    return rating


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        print("Collecting all Ratings data")
        glassdoor_link = 'https://www.glassdoor.co.in/Overview/Working-at-Signpost-India-EI_IE2372115.11,25.htm'
        ambitionBox_link = 'https://www.ambitionbox.com/reviews/signpost-india-reviews'
        justDial_link ='https://www.justdial.com/jdmart/Mumbai/Signpost-India-Pvt-Ltd-Registered-Office-Near-Santacruz-Airport-Terminal-Vile-Parle-East/022PXX22-XX22-181127114814-B8L6_BZDET/catalogue'
        crisil_link = 'https://www.crisilratings.com/en/home/our-business/ratings/company-factsheet.CTODAL.html'
        ticker_link = 'https://ticker.finology.in/company/SIGNPOST'
        rating =dict()
        rating['company_name'] = 'Signpost India Ltd'
        rating['glassdoor'] = find_glassdoorRating(page, glassdoor_link)
        rating['ambitionbox'] = find_ambitionBoxRating(page, ambitionBox_link)
        rating['justdial'] = find_justDialRating(page, justDial_link)
        rating['crisil'] = find_crisilRating(page, crisil_link)
        rating['ticker'] = find_tickerRating(page, ticker_link)
        browser.close()
        print(rating)
        keys = list(rating.keys())
        values = list(rating.values())

        # Creating the DataFrame with keys in 'Entity' and values in 'Value'
        df = pd.DataFrame({
            'Entity': keys,
            'Value': values
        })
        df = df.transpose()
        # Save the DataFrame to an Excel file
        df.to_excel('company_ratings.xlsx', index=False)

if __name__ == "__main__":
    main()


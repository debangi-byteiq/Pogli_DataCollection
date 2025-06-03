import time
from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup

# List of URLs to scrape
url = [
    "https://www.investing.com/equities/tata-communications-historical-data?cid=39698",

]

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_p
        age()
        page.goto(url[0])
        time.sleep(6)
        page.close()

if __name__ == "__main__":
    main()
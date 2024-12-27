import re
import time
import requests
import warnings
from datetime import date
import pandas as pd
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto('https://www.screener.in/company/LIBERTSHOE/')
        page.get_by_text('Read More').click()
        time.sleep(2)
        email_input = page.locator("input#id_email")
        email_input.fill("debangipanda@gmail.com")
        email2_input = page.locator('input#id_email2')
        email2_input.fill("debangipanda@gmail.com")
        password_input = page.locator('input#id_password')
        password_input.fill("Debangi@1922")
        time.sleep(1)
        search_button = page.locator("button.button.button-primary.u-full-width.text-align-center")
        search_button.click()
        time.sleep(10)
        page.get_by_text('Read More').click()
        time.sleep(2)

        details = page.query_selector_all('div.sub')
        details_list = []
        for detail in details:
            p_tags = detail.query_selector_all('p')
            for p_tag in p_tags:
                details_list.append(p_tag.text_content())
        print(details_list)


if __name__ == '__main__':
    main()
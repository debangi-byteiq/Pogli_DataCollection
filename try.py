from playwright.sync_api import sync_playwright
import re
import pandas as pd
import time


def search_and_get_rating(page, base_url, search_selector, query, result_selector, rating_selector, rating_regex):
    # Navigate to the base URL
    page.goto(base_url)
    time.sleep(2)

    try:
        # Enter the query in the search bar and submit
        search_bar = page.locator(search_selector)
        search_bar.fill(query)
        search_bar.press("Enter")
        time.sleep(3)  # Wait for the results to load

        # Click the first relevant result
        result_link = page.locator(result_selector).first
        result_link.click()
        time.sleep(3)  # Wait for the company page to load

        # Extract the rating
        rating_element = page.locator(rating_selector)
        text = rating_element.text_content().strip()
        rating = re.search(rating_regex, text).group(1)
    except Exception as e:
        print(f"Error: {e}")
        rating = "Not found"

    return rating


def main():
    company_name = "Signpost India"
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        # Define details for each website
        websites = [
            {
                "name": "Glassdoor",
                "base_url": "https://www.glassdoor.co.in/Reviews/index.htm",
                "search_selector": "input#companyAutocomplete-companyDiscover-employerSearch",  # Update with the actual search bar selector
                "result_selector": "ul.suggestions.down",  # Update with the actual result selector
                "rating_selector": "p.rating-headline-average_rating__J5rIy",
                "rating_regex": r"(\d*\.?\d*)"
            },
            # {
            #     "name": "AmbitionBox",
            #     "base_url": "https://www.ambitionbox.com/",
            #     "search_selector": "input[placeholder='Search Companies']",  # Example selector
            #     "result_selector": "a[class='company-result']",  # Example selector
            #     "rating_selector": "div.rating_text.rating_text--md",
            #     "rating_regex": r"(\d*\.?\d*)"
            # },
            # Add more sites as needed
        ]

        ratings = {"Company Name": company_name}
        for site in websites:
            print(f"Fetching rating from {site['name']}...")
            rating = search_and_get_rating(
                page=page,
                base_url=site["base_url"],
                search_selector=site["search_selector"],
                query=company_name,
                result_selector=site["result_selector"],
                rating_selector=site["rating_selector"],
                rating_regex=site["rating_regex"]
            )
            ratings[site["name"]] = rating

        browser.close()

        # Save ratings to an Excel file
        df = pd.DataFrame([ratings])
        df.to_excel("ratings.xlsx", index=False)
        print("Ratings saved to ratings.xlsx")


if __name__ == "__main__":
    main()

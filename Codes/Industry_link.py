import requests
from bs4 import BeautifulSoup
URL = "https://www.bseindia.com/markets/Equity/EQReports/industrywatch.aspx?page=IN070201001&scripname=Aerospace+&+Defense"
HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


def fetch_page_content(url: str) -> BeautifulSoup:
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        return BeautifulSoup(response.content, "html.parser")
    except requests.RequestException as e:
        raise e


soup = fetch_page_content(URL)
options = soup.select("#ContentPlaceHolder1_ddlIndustry > option")
options_list = []
for index, option in enumerate(options):
    opt_name = option.get_text()
    opt_value = option.get("value")

    base_url = f"https://www.bseindia.com/markets/Equity/EQReports/IndustryView.html?expandable=2&page={opt_value}%20&scripname={opt_name.replace(' ', '+')}"
    # print(f"{index+1}. {opt_name} {opt_value}",base_url)
    options_list.append(base_url)
print((options_list))

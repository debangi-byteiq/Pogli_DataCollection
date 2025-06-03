import asyncio
from playwright.async_api import async_playwright
from bs4 import BeautifulSoup


async def fetch_html(url):
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page()
        await page.goto(url)
        content = await page.content()
        await browser.close()
        return content


async def main():
    url = "https://www.bseindia.com/stock-share-price/ankit-metal--power-ltd/ankitmetal/532870/"

    html_content = await fetch_html(url)

    soup = BeautifulSoup(html_content, 'html.parser')

    target_div = soup.find('div', class_='container-fluid whitebox marketstartarea ng-scope')

    a= soup.find('div',class_='row newgraphdayarea' ).find('a', id='lnk6M')
    print(a)



asyncio.run(main())

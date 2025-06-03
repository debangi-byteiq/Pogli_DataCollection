import requests
import json
import pandas as pd

cookies = {
    '_ga': 'GA1.1.196071947.1748244001',
    'TBMCookie_7892651951369330535': '1968200017485896608hfMeLXfM/utO0VsT98j1gN0xm0=',
    '___utmvm': '###########',
    '__gads': 'ID=aadb186b006514b5:T=1748496467:RT=1748597683:S=ALNI_MZZlL_OfrQqNOZ2l86rG-tE0ju93Q',
    '__gpi': 'UID=0000110517e68842:T=1748496467:RT=1748597683:S=ALNI_Ma1tzi5ZO1b5JwZhdPt0wamoOt6vA',
    '__eoi': 'ID=4bc3ac02b29e654f:T=1748496467:RT=1748597683:S=AA-AfjZSLecscAlTwxL-kRtch1fN',
    '_ga_TM52BJH9HF': 'GS2.1.s1748597682$o23$g1$t1748597683$j59$l0$h0',
}

headers = {
    'Accept': '*/*',
    'Accept-Language': 'en-US,en;q=0.9',
    'Connection': 'keep-alive',
    # 'Content-Length': '0',
    'Origin': 'https://charting.bseindia.com',
    'Referer': 'https://charting.bseindia.com/index.html?SYMBOL=500078',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36',
    'X-Requested-With': 'XMLHttpRequest',
    'sec-ch-ua': '"Google Chrome";v="137", "Chromium";v="137", "Not/A)Brand";v="24"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    # 'Cookie': '_ga=GA1.1.196071947.1748244001; TBMCookie_7892651951369330535=1968200017485896608hfMeLXfM/utO0VsT98j1gN0xm0=; ___utmvm=###########; __gads=ID=aadb186b006514b5:T=1748496467:RT=1748597683:S=ALNI_MZZlL_OfrQqNOZ2l86rG-tE0ju93Q; __gpi=UID=0000110517e68842:T=1748496467:RT=1748597683:S=ALNI_Ma1tzi5ZO1b5JwZhdPt0wamoOt6vA; __eoi=ID=4bc3ac02b29e654f:T=1748496467:RT=1748597683:S=AA-AfjZSLecscAlTwxL-kRtch1fN; _ga_TM52BJH9HF=GS2.1.s1748597682$o23$g1$t1748597683$j59$l0$h0',
}

response = requests.post(
    'https://charting.bseindia.com/charting/RestDataProvider.svc/getDat?exch=N&scode=500078&type=b&mode=bseL&fromdate=01-01-1991-01:01:00-AM',
    cookies=cookies,
    headers=headers,
)
a = response.json()['getDatResult']


a =json.loads(a)
print(len(a['DataInputValues']))
# df = pd.DataFrame(a['DataInputValues'])
# print(df)

print(response.json())
# open_str = a['DataInputValues'][0]['OpenData'][0]['Open']
# date_str = a['DataInputValues'][0]['DateData'][0]['Date']
#
# open_list = open_str.split(',')
# date_list = date_str.split(',')
#
# # # Convert to DataFrame
# # df = pd.DataFrame({
# #     'Date': date_list,
# #     'Open': open_list
# # })
#
#
# df = pd.DataFrame({
#     'Date': pd.to_datetime(date_list, format='%d/%m/%Y %I:%M:%S %p', errors='coerce'),
#     'Open': pd.to_numeric(open_list, errors='coerce')
# })
#
# # Drop any rows with bad date conversion
# df = df.dropna(subset=['Date'])
#
# # Add year and month columns
# df['Year'] = df['Date'].dt.year
# df['Month'] = df['Date'].dt.month
#
# # Sort by date ascending
# df = df.sort_values('Date')
# first_per_month = df.groupby(['Year', 'Month']).first().reset_index()
# # Drop 'Year' and 'Month' columns
# first_per_month = first_per_month.drop(['Year', 'Month'], axis=1)
# last_12_months = first_per_month.tail(12)
#
# print(last_12_months)
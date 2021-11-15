import requests
import pandas as pd
from bs4 import BeautifulSoup


headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0 Win64 x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"}

url = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY?stockNo={}&response=html'.format(
    '1711')
url = 'https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_print.php?l=zh-tw&stkno={}&s=0,asc,0'.format(
    '3141')

df = pd.read_html(url)
df = df[0]
print(df)
df.to_excel('TEMP_TWSE.xlsx')

df = pd.read_excel('TEMP_TWSE.xlsx')
print(df['Unnamed: 7'])

import xlwings as xw
import pandas as pd
import numpy as np
from datetime import date
import datetime
import requests
from bs4 import BeautifulSoup
import time


ToExcel = True
DoNotRequest = False
FromGoodInfo = False


headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0 Win64 x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"}

today = date.today()


# 計算每次交易記錄當下支出成本
# `TradingHistory.xlsx` 交易記錄
# 每次交易成本 = Position * Price
def calculate_each_cost(each_cost_data):
    return each_cost_data['PositionSize'] * each_cost_data['Price']


# 下載當日收盤價
def download_today_close(each_today_close):
    if FromGoodInfo:
        url = 'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID={}'.format(
            each_today_close['Stock_Id'])
        response = requests.get(url, headers=headers)
        response.encoding = 'utf8'
        soup = BeautifulSoup(response.text, 'lxml')
        table = soup.find('table', class_='b1 p4_2 r10')
        try:
            tr = table.find('tr', attrs={'align': 'center'})
            td = tr.find('td')
            print('Stock_Id:',
                  each_today_close['Stock_Id'], '下載收盤價 : ', td.string)
            time.sleep(15)
            return(td.string)
        except:
            print('瀏覽量異常 from goodinfo')
            return(0)
    else:
        url = 'https://www.twse.com.tw/exchangeReport/STOCK_DAY?stockNo={}&response=html'.format(
            str(each_today_close['Stock_Id'])[:4])
        try:
            temp = pd.read_html(url)
            temp = temp[0]
            temp.to_excel('TEMP_TWSE.xlsx')
            temp = pd.read_excel('TEMP_TWSE.xlsx')
            print('Stock_Id:',
                  str(each_today_close['Stock_Id'])[:4], '下載收盤價 : ', temp['Unnamed: 7'].to_list()[-1])
            time.sleep(15)
            return(temp['Unnamed: 7'].to_list()[-1])
        except:
            url = 'https://www.tpex.org.tw/web/stock/aftertrading/daily_trading_info/st43_print.php?l=zh-tw&stkno={}&s=0,asc,0'.format(
                str(each_today_close['Stock_Id'])[:4])
            temp = pd.read_html(url)
            temp = temp[0]
            temp.to_excel('TEMP_TWSE.xlsx')
            temp = pd.read_excel('TEMP_TWSE.xlsx')
            print('Stock_Id:',
                  str(each_today_close['Stock_Id'])[:4], '下載收盤價 : ', temp['Unnamed: 7'].to_list()[-2])
            time.sleep(15)
            return(temp['Unnamed: 7'].to_list()[-2])


# 計笡庫存個別資產
def calculate_each_capital(each_capital):
    return float(each_capital['PositionSize']) * float(each_capital['TodayClose'])


def calculate_each_profit_percent(each_profit_percent):
    cost = each_profit_percent['EachCost_x'] - \
        each_profit_percent['EachCost_y']
    flat_price = each_profit_percent['EachCost_x']
    profit = (flat_price - cost) / cost
    profit = round(profit, 2)
    return profit


def calculate_each_profit_money(each_profit_money):
    return each_profit_money['EachCost_y'] * (-1)

# ===================================================================================================


df_trading_history = pd.read_excel(
    'TradingModel_TradingHistory.xlsx', sheet_name='交易記錄')


# 計算每次交易記錄當下支出成本
df_trading_history['EachCost'] = df_trading_history.apply(
    calculate_each_cost, axis=1)
if ToExcel:
    df_trading_history.to_excel('TEMP_TradingModel_TradingHistory.xlsx')


# 計算/輸出當日平倉表單
# 損益
today_close_position = df_trading_history[df_trading_history['Date'] == str(
    today)]
today_close_position = today_close_position[today_close_position['Action'] == 'short']
if ToExcel:
    today_close_position.to_excel(
        'TEMP_TradingModel_TodayClosePosition.xlsx', sheet_name='今日平倉', index=False)


# 計算/輸出目前總庫存表單
stock_history_group = df_trading_history.groupby('Stock_Id')
total_open_postion = stock_history_group.sum()
total_open_postion.reset_index(inplace=True)
# TEMP_TradingModel_TodayCloseHistory >>
today_close_history = pd.merge(
    today_close_position, total_open_postion, on='Stock_Id')
try:
    today_close_history['ProfitPercent'] = today_close_history.apply(
        calculate_each_profit_percent, axis=1)
    today_close_history['ProfitMoney'] = today_close_history.apply(
        calculate_each_profit_money, axis=1)
    today_close_history = today_close_history[[
        'Date', 'Stock_Id', 'ProfitPercent', 'ProfitMoney']]
    if ToExcel:
        today_close_history.to_excel(
            'TEMP_TradingModel_TodayCloseHistory.xlsx', sheet_name='今日平倉損益')
except:
    print('TodayCloseHistory : None')
# <<
total_open_postion = total_open_postion[total_open_postion['PositionSize'] != 0]
realtime_watching = total_open_postion
total_open_postion = total_open_postion.drop(
    axis=1, columns=['Price', 'EachCost'])  # drop column, axis=1
if ToExcel:
    total_open_postion.to_excel(
        'TradingModel_TotalOpenPosition.xlsx', sheet_name='庫存')


if not DoNotRequest:
    # 記錄庫存今日收盤價
    open_position_today_close = total_open_postion.drop(
        axis=1, columns=['PositionSize'])
    open_position_today_close["TodayClose"] = open_position_today_close.apply(
        download_today_close, axis=1)
    if ToExcel:
        open_position_today_close.to_excel(
            'TradingModel_OpenPositionTodayClose.xlsx', sheet_name='收盤價')
else:
    open_position_today_close = pd.read_excel(
        'TradingModel_OpenPositionTodayClose.xlsx')


# 輸出每日總資產
# 'TradingModel_TotalCapital.xlsx'
total_open_position_capital = pd.merge(
    total_open_postion, open_position_today_close)
total_open_position_capital['EachCapital'] = total_open_position_capital.apply(
    calculate_each_capital, axis=1)

total_capital = pd.read_excel('TradingModel_TotalCapital.xlsx')
if total_capital['Date'].to_list()[-1].date() != today:
    total_capital = total_capital.append(
        {'Date': today, 'TotalCapital': total_open_position_capital['EachCapital'].sum()}, ignore_index=True)
    if ToExcel:
        total_capital.to_excel(
            'TradingModel_TotalCapital.xlsx', index=False)


# 輸出個股歷史平倉損益
# TradingModel_TotalEachProfitHistory.xlsx
total_each_profit = pd.read_excel('TradingModel_TotalEachProfitHistory.xlsx')
if total_each_profit['Date'].to_list()[-1].date() != today:
    total_each_profit = total_each_profit.append(
        today_close_history, ignore_index=True)
if ToExcel:
    total_each_profit.to_excel(
        'TradingModel_TotalEachProfitHistory.xlsx', index=False)


# 輸出歷史總平倉損益
# TradingModel_TotalProfit.xlsx
total_profit = pd.read_excel('TradingModel_TotalProfit.xlsx')
if total_profit['Date'].to_list()[-1].date() != today:
    total_profit = total_profit.append(
        {'Date': today, 'TotalProfit': total_each_profit['ProfitMoney'].sum()}, ignore_index=True)
    if ToExcel:
        total_profit.to_excel('TradingModel_TotalProfit.xlsx', index=False)

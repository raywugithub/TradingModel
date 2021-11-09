import xlwings as xw
import pandas as pd
import numpy as np
from datetime import date
import datetime
import requests
from bs4 import BeautifulSoup
import time
from openpyxl import load_workbook

ToExcel = True
DoNotRequest = False

headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0 Win64 x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/95.0.4638.69 Safari/537.36"}

today = date.today()


# 計算每次交易記錄當下支出成本
# `TradingHistory.xlsx` 交易記錄
# 每次交易成本 = Position * Price
def calculate_each_cost(each_cost_data):
    return each_cost_data['Position'] * each_cost_data['Price']


# 下載當日收盤價
def download_today_close(each_today_close):
    url = 'https://goodinfo.tw/StockInfo/StockDetail.asp?STOCK_ID={}'.format(
        each_today_close['Stock_Id'])
    response = requests.get(url, headers=headers)
    response.encoding = 'utf8'
    soup = BeautifulSoup(response.text, 'lxml')
    table = soup.find('table', class_='b1 p4_2 r10')
    try:
        tr = table.find('tr', attrs={'align': 'center'})
        td = tr.find('td')
        print('Stock_Id:', each_today_close['Stock_Id'], '下載收盤價')
        time.sleep(15)
        return(td.string)
    except:
        print('瀏覽量異常 from goodinfo')
        return(0)


# 計笡庫存個別資產
def calculate_each_capital(each_capital):
    return float(each_capital['Position']) * float(each_capital['TodayClose'])


# 檢查目前庫存的收盤價是否進入平台
def check_each_open_opsition_lost_previous_n(each_open_position):
    if each_open_position['TodayClose'] < each_open_position['Previous_N_High']:
        return True
    else:
        return False


def check_each_open_opsition_lost_previous_platform(each_open_position):
    if each_open_position['TodayClose'] < each_open_position['Previous_Platform_High']:
        return True
    else:
        return False
# ===================================================================================================


df_trading_history = pd.read_excel(
    'TEST_TradingModel_TradingHistory.xlsx', sheet_name='交易記錄')


# 計算每次交易記錄當下支出成本
df_trading_history['EachCost'] = df_trading_history.apply(
    calculate_each_cost, axis=1)
if ToExcel:
    df_trading_history.to_excel('TEMP_TradingModel_TradingHistory.xlsx')


# 計算/輸出目前總庫存表單
stock_history_group = df_trading_history.groupby('Stock_Id')
total_open_postion = stock_history_group.sum()
total_open_postion = total_open_postion.drop(
    axis=1, columns=['Price', 'EachCost'])  # drop column, axis=1
total_open_postion.reset_index(inplace=True)
if ToExcel:
    total_open_postion.to_excel(
        'TEST_TradingModel_TotalOpenPosition.xlsx', sheet_name='庫存')


if not DoNotRequest:
    # 記錄庫存今日收盤價
    open_position_today_close = total_open_postion.drop(
        axis=1, columns=['Position'])
    open_position_today_close["TodayClose"] = open_position_today_close.apply(
        download_today_close, axis=1)
    if ToExcel:
        open_position_today_close.to_excel(
            'TEST_TradingModel_OpenPositionTodayClose.xlsx', sheet_name='收盤價')
else:
    open_position_today_close = pd.read_excel(
        'TEST_TradingModel_OpenPositionTodayClose.xlsx')


# 輸出每日總資產
# 'TEST_TradingModel_TotalCapital.xlsx'
total_open_position_capital = pd.merge(
    total_open_postion, open_position_today_close)
total_open_position_capital['EachCapital'] = total_open_position_capital.apply(
    calculate_each_capital, axis=1)

total_capital = pd.read_excel('TEST_TradingModel_TotalCapital.xlsx')
if total_capital['Date'].to_list()[-1].date() != today:
    total_capital = total_capital.append(
        {'Date': today, 'TotalCapital': total_open_position_capital['EachCapital'].sum()}, ignore_index=True)
    if ToExcel:
        total_capital.to_excel(
            'TEST_TradingModel_TotalCapital.xlsx', index=False)


# 檢查目前庫存的收盤價是否進入平台
# 'TEST_TradingModel_OpenPositionWatching.xlsx'
open_position_watching = pd.read_excel(
    'TEST_TradingModel_OpenPositionWatching.xlsx')
open_position_check = pd.merge(
    open_position_today_close, open_position_watching)
open_position_check['Lost_Previous_N'] = open_position_check.apply(
    check_each_open_opsition_lost_previous_n, axis=1)
open_position_check['Lost_Previous_Platform'] = open_position_check.apply(
    check_each_open_opsition_lost_previous_platform, axis=1)
for n in range(len(open_position_check['Stock_Id'])):
    if open_position_check['Lost_Previous_N'].to_list()[n] == True:
        print('Warning !!!!',
              open_position_check['Stock_Id'].to_list()[n], '失守前 N')
    if open_position_check['Lost_Previous_Platform'].to_list()[n] == True:
        print('Warning !!!!',
              open_position_check['Stock_Id'].to_list()[n], '進入前整理平台')
open_position_check.to_excel(
    'TEMP_TradingModel_OpenPositionWatching.xlsx', index=False)

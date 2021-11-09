import pandas as pd

sorted = True

if not sorted:
    goodinfo_stocklist = pd.read_csv('goodinfo_StockList_21WW45.csv')

    sorted_id = []
    sorted_name = []
    sorted_low_month = []
    sorted_low_three_month = []
    sorted_low_six_month = []
    sorted_low_year = []
    sorted_diff_five_day = []
    sorted_diff_month = []
    sorted_diff_three_month = []
    sorted_diff_six_month = []
    sorted_diff_year = []
    sorted_big_buy = []
    sorted_big_buy_percent = []

    str_low_month = '一個月最低股價'
    str_low_three_month = '三個月最低股價'
    str_low_six_month = '半年最低股價'
    str_low_year = '一年最低股價'
    str_diff_five_day = '5日累計漲跌(%)'
    str_diff_month = '一個月累計漲跌(%)'
    str_diff_three_month = '三個月累計漲跌(%)'
    str_diff_six_month = '半年累計漲跌(%)'
    str_diff_year = '一年累計漲跌(%)'
    str_big_buy = '21W45外資買賣超張數'
    str_big_buy_percent = '21W45外資買賣超佔成交(%)'

    list_low_month = goodinfo_stocklist[str_low_month].to_list()
    list_low_three_month = goodinfo_stocklist[str_low_three_month].to_list()
    list_low_six_month = goodinfo_stocklist[str_low_six_month].to_list()
    list_low_year = goodinfo_stocklist[str_low_year].to_list()
    list_diff_five_day = goodinfo_stocklist[str_diff_five_day].to_list()
    list_diff_month = goodinfo_stocklist[str_diff_month].to_list()
    list_diff_three_month = goodinfo_stocklist[str_diff_three_month].to_list()
    list_diff_six_month = goodinfo_stocklist[str_diff_six_month].to_list()
    list_diff_year = goodinfo_stocklist[str_diff_year].to_list()
    list_big_buy = goodinfo_stocklist[str_big_buy].to_list()
    list_big_buy_percent = goodinfo_stocklist[str_big_buy_percent].to_list()

    for n in range(len(goodinfo_stocklist['代號'])):
        if (list_low_six_month[n] > list_low_year[n]) and (list_low_three_month[n] > list_low_six_month[n]) and (list_low_month[n] > list_low_three_month[n]):
            sorted_id.append(goodinfo_stocklist['代號'].to_list()[n])
            sorted_name.append(goodinfo_stocklist['名稱'].to_list()[n])
            sorted_low_month.append(goodinfo_stocklist['一個月最低股價'].to_list()[n])
            sorted_low_three_month.append(
                goodinfo_stocklist['三個月最低股價'].to_list()[n])
            sorted_low_six_month.append(
                goodinfo_stocklist['半年最低股價'].to_list()[n])
            sorted_low_year.append(goodinfo_stocklist['一年最低股價'].to_list()[n])
            sorted_diff_five_day.append(
                goodinfo_stocklist['5日累計漲跌(%)'].to_list()[n])
            sorted_diff_month.append(
                goodinfo_stocklist['一個月累計漲跌(%)'].to_list()[n])
            sorted_diff_three_month.append(
                goodinfo_stocklist['三個月累計漲跌(%)'].to_list()[n])
            sorted_diff_six_month.append(
                goodinfo_stocklist['半年累計漲跌(%)'].to_list()[n])
            sorted_diff_year.append(
                goodinfo_stocklist['一年累計漲跌(%)'].to_list()[n])
            sorted_big_buy.append(
                goodinfo_stocklist['21W45外資買賣超張數'].to_list()[n])
            sorted_big_buy_percent.append(
                goodinfo_stocklist['21W45外資買賣超佔成交(%)'].to_list()[n])

    df = pd.DataFrame({
        "Stock_Id": sorted_id,
        "Name": sorted_name,
        'low_month': sorted_low_month,
        'low_three_month': sorted_low_three_month,
        'low_six_month': sorted_low_six_month,
        'low_year': sorted_low_year,
        'diff_five_day': sorted_diff_five_day,
        'diff_month': sorted_diff_month,
        'diff_three_month': sorted_diff_three_month,
        'diff_six_month': sorted_diff_six_month,
        'diff_year': sorted_diff_year,
        'big_buy': sorted_big_buy,
        'big_buy_percent': sorted_big_buy_percent
    })

    # df.to_excel('goodinfo_StockList_21WW45.xlsx')
else:
    print('sorted')
    sorted_df = pd.read_excel('goodinfo_StockList_21WW45_sorted.xlsx')

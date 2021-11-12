import pandas as pd


sorted = True
input_csv = 'goodinfo_StockList_21W45_Thu_2nd.csv'
output_xlsx = 'goodinfo_StockList_21W45_Thu_sorted.xlsx'


def stair_condition(stair_df):
    if (stair_df['半年最低股價'] > stair_df['一年最低股價']):
        if (stair_df['三個月最低股價'] > stair_df['半年最低股價']):
            if(stair_df['一個月最低股價'] > stair_df['三個月最低股價']):
                return 'YES'


if not sorted:
    goodinfo_stocklist = pd.read_csv(input_csv)

    # 代號
    # 名稱
    # 成交
    # 漲跌幅
    # 一個月最低股價
    # 三個月最低股價
    # 半年最低股價
    # 一年最低股價
    # 5日累計漲跌(%)
    # 一個月累計漲跌(%)
    # 三個月累計漲跌(%)
    # 半年累計漲跌(%)
    # 一年累計漲跌(%)
    # 21W45外資買賣超張數
    # 21W45外資買賣超佔成交(%)
    # 一個月最高股價
    # 三個月最高股價
    # 半年最高股價
    # 一年最高股價
    sorted_df = pd.DataFrame()
    for n in range(len(goodinfo_stocklist['代號'])):
        goodinfo_stocklist['Condition'] = goodinfo_stocklist.apply(
            stair_condition, axis=1)
    goodinfo_stocklist = goodinfo_stocklist[goodinfo_stocklist['Condition'] == 'YES']
    goodinfo_stocklist.to_excel(output_xlsx)

else:
    print('sorted')
    #sorted_df = pd.read_excel(output_xlsx).rename(columns={'代號': 'Stock_Id'})
    #class_df = pd.read_excel('goodinfo_StockList_21W45_Thu_sorted.xlsx')

    ##sorted_df = sorted_df[['Stock_Id']]
    ##class_df = class_df[['Stock_Id']]
    #df = sorted_df.merge(class_df, on='Stock_Id')
    # df = df[['Stock_Id', '名稱', '成交', '漲跌幅', '一個月最低股價', '三個月最低股價', '半年最低股價', '一年最低股價', '5日累計漲跌(%)', '一個月累計漲跌(%)', '三個月累計漲跌(%)', '半年累計漲跌(%)', '一年累計漲跌(%)', '21W45外資買賣超張數', '21W45外資買賣超佔成交(%)', '一個月最高股價', '三個月最高股價', '半年最高股價', '一年最高股價', 'Condition', 'Class'
    #         ]]
    #df = df.rename(columns={'Stock_Id': '代號'})
    # df.to_excel('goodinfo_StockList_21W45_Thu_sorted.xlsx')

    sorted_df = pd.read_excel(output_xlsx)
    sorted_df = sorted_df[['代號', '名稱', 'Class']].sort_values(by='Class')
    sorted_df = sorted_df[sorted_df['Class'] >= 2]
    sorted_df.to_excel('temp_'+output_xlsx)

# "=XQCTYAP|Quote!'1201.TW-Name,Price'"

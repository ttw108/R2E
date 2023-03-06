import os
import  pandas as pd
import warnings
from sklearn.decomposition import PCA
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import numpy as np

def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col

#1 處理日期 開啟 book.xlsx '新增 profit 工作表 、加入日期180天
def open_xlsx():
    # 打開一個新的Excel文件
    global wb
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = True
    wb = xw.Book('book.xlsx')
    global sht_profit
    if not 'profit' in wb.sheet_names:
        sht_profit = wb.sheets.add('profit')
    else:
        wb.sheets['profit'].delete()
        sht_profit = wb.sheets.add('profit')

    # 從當前工作表中取得第一個儲存格（A2）
    cell = sht_profit.range('A2')
    # 使用Python中的datetime模組來得到當前月份的資訊
    import datetime
    # 第一步基本資料處理
    # 將開始日期輸入到A1儲存格中
    dayrng = 180
    start_date = datetime.date.today() - datetime.timedelta(days=dayrng)
    cell.value = start_date
    # 將開始日期往下移動一格，並輸出相對應的日期資訊
    sht_profit.cells(1, 1).value = '日期'

    for i in range(1, dayrng):
        cell.offset(row_offset=i, column_offset=0).value = start_date + datetime.timedelta(days=i)
        # 找到記錄XLS"Book1.xls"的最後一行
        # 找到最後一行
    global last_cell
    last_cell = sht_profit.range('A2').current_region.rows.count
    print("Profit_工作表_日期建立完成")
open_xlsx()

#2先找到目錄中    ".xls" 把180天交易損益寫進 PROFIT工作表
def fill_strategy_profit():
    fd = "./trading_xls"
    files = os.listdir(fd)
    # 將檔案按建立時間排序
    files = sorted(files, key=lambda x: os.stat(os.path.join(fd, x)).st_mtime)

    fn = 1
    global c
    c = 1
    for filename in files:
        if filename.split(".")[-1] == "xls" or filename.split(".")[-1] =='xlsx':
            if re.match('.*策略回測績效報告.*', filename):
                order_name = filename
                # print(order_name)
                app1 = xw.App(visible=False, add_book=False)
                app1.display_alerts = False
                app1.screen_updating = False
                wb1 = app1.books.open(fd + "/" + order_name)
                sht1 = wb1.sheets('交易明細')

                # 將 sheet 內容轉換為 dataframe
                df = sht1.range('B3').options(pd.DataFrame, expand='table').value
                # 移除 column_name 為空的 row
                df.dropna(subset=[r"獲利(¤)"], inplace=True)
                df['日期'] = pd.to_datetime(df['日期']).dt.date

                # 180日期中，一個一個日期字串拉出來 篩選 並取得加總：
                sht_profit.cells(1, c + 1).value = c
                d0 = 1
                for dstr in sht_profit.range('A2:A{}'.format(last_cell)).value:
                    profit = df['獲利(¤)'][(df['日期'] == pd.to_datetime(dstr))].sum()
                    sht_profit.cells(d0 + 1, fn + 1).value = profit
                    d0 = d0 + 1
                wb1.close()
                c = c + 1

            sht_profit.cells(200 + fn, 1).value = c - 1
            sht_profit.cells(200 + fn, 2).value = filename
            print(str(fn)+": "+ filename)
            fn = fn + 1
fill_strategy_profit()

#conv_to_df()策略日損益存至book.xlsx（很慢）
def conv_to_df():
    last_column = sht_profit.range('A2').current_region.columns.count
    rng_all = sht_profit.range('A1:{}'.format(num_to_col(last_column) + str(last_cell)))
    global dfa
    dfa = pd.DataFrame(rng_all.value, columns=rng_all.value[0])
    # 將 '日期' 列轉換為索引
    dfa = df = dfa.drop(dfa.index[0])
    dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
    dfa = dfa.set_index('日期')
    dfa = dfa.rename(columns=lambda x: str(x).strip())
    dfa = dfa.rename(columns=lambda x: x[:-2] if x.endswith('.0') else x)
    dfa.to_pickle('dfa.pkl')
    pd.read_pickle('dfa.pkl')
conv_to_df()

def sqn_func(dfa):
    # 读取数据
    dfs = dfa
    # 计算每个策略的平均每笔收益和标准差
    avg_returns = dfs.mean()
    std_returns = dfs.std()

    # 定义SQN函数
    def sqn(weights, avg_returns, std_returns):
        combined_returns = np.dot(weights, avg_returns)
        combined_std = np.sqrt(np.dot(weights.T, np.dot(np.cov(dfs.T), weights)))
        sqn = np.sqrt(len(dfs)) * combined_returns / combined_std
        return -sqn  # 目标函数是最小化负的SQN值

    # 定义约束条件
    n_assets = len(dfs.columns)
    constraints = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
    bounds = [(0, 1) for i in range(n_assets)]
    # 初始权重
    weights = np.ones(n_assets) / n_assets
    # 最小化负的SQN函数，求得最优权重
    result = minimize(sqn, weights, args=(avg_returns, std_returns), method='SLSQP',
                      bounds=bounds, constraints=constraints)
    # 打印结果
    print('Optimal weights:', result.x)
    print('SQN value:', -sqn(result.x, avg_returns, std_returns))
    ws = result.x
    dfx = pd.DataFrame({'Strategy': np.arange(1, ws.shape[0] + 1),
                        'Weight': ws})
    return(dfx)

#將 a1,b1,c1 做平均加總(a1,b1,c1)及加權加總（a1*0.5+b1*0.3+c1*0.2）
def sqn_abc_all_export():
    global df_all
    # 本週90日的SQN
    dfa.tail(1).index
    df_a1 = dfa.iloc[-90:, :]
    # 從最後一天往前推 10 天，再選擇 90 天的範圍
    df_b1 = dfa.iloc[-110:-20, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_c1 = dfa.iloc[-130:-40, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_d1 = dfa.iloc[-150:-60, :]

    global  sqn_a1, sqn_b1, sqn_c1, sqn_d1
    sqn_a1 = sqn_func(df_a1)
    sqn_b1 = sqn_func(df_b1)
    sqn_c1 = sqn_func(df_c1)
    sqn_d1 = sqn_func(df_d1)

    sqn_a1.to_pickle('sqn_a1.pkl')
    sqn_b1.to_pickle('sqn_b1.pkl')
    sqn_c1.to_pickle('sqn_c1.pkl')
    sqn_d1.to_pickle('sqn_d1.pkl')
    # test=pd.read_pickle('sqn_a1.pkl')
sqn_abc_all_export()
################################################################
dfa = pd.read_pickle('dfa.pkl')
#df_all = pd.read_pickle('df_all.pkl')
################################################################
def sqn_func(dfa):
    # 读取数据
    dfs = dfa
    # 计算每个策略的平均每笔收益和标准差
    avg_returns = dfs.mean()
    std_returns = dfs.std()

    # 定义SQN函数
    def sqn(weights, avg_returns, std_returns):
        combined_returns = np.dot(weights, avg_returns)
        combined_std = np.sqrt(np.dot(weights.T, np.dot(np.cov(dfs.T), weights)))
        sqn = np.sqrt(len(dfs)) * combined_returns / combined_std
        return -sqn  # 目标函数是最小化负的SQN值

    # 定义约束条件
    n_assets = len(dfs.columns)
    constraints = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
    bounds = [(0, 1) for i in range(n_assets)]
    # 初始权重
    weights = np.ones(n_assets) / n_assets
    # 最小化负的SQN函数，求得最优权重
    result = minimize(sqn, weights, args=(avg_returns, std_returns), method='SLSQP',
                      bounds=bounds, constraints=constraints)
    # 打印结果
    print('Optimal weights:', result.x)
    print('SQN value:', -sqn(result.x, avg_returns, std_returns))
    ws = result.x
    dfx = pd.DataFrame({'Strategy': np.arange(1, ws.shape[0] + 1),
                        'Weight': ws})
    return(dfx)

#將 a1,b1,c1 做平均加總(a1,b1,c1)及加權加總（a1*0.5+b1*0.3+c1*0.2）
def sqn_abc_all_export():
    global df_all
    # 本週90日的SQN
    dfa.tail(1).index
    df_a1 = dfa.iloc[-90:, :]
    # 從最後一天往前推 10 天，再選擇 90 天的範圍
    df_b1 = dfa.iloc[-110:-20, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_c1 = dfa.iloc[-130:-40, :]
    # 從最後一天往前推 20 天，再選擇 90 天的範圍
    df_d1 = dfa.iloc[-150:-60, :]

    global  sqn_a1, sqn_b1, sqn_c1, sqn_d1
    sqn_a1 = sqn_func(df_a1)
    sqn_b1 = sqn_func(df_b1)
    sqn_c1 = sqn_func(df_c1)
    sqn_d1 = sqn_func(df_d1)

    sqn_a1.to_pickle('sqn_a1.pkl')
    sqn_b1.to_pickle('sqn_b1.pkl')
    sqn_c1.to_pickle('sqn_c1.pkl')
    sqn_d1.to_pickle('sqn_d1.pkl')
    # test=pd.read_pickle('sqn_a1.pkl')
sqn_abc_all_export()


# 策略名稱
# 選擇要導出的範圍
def data_load():
    global df_all,sqn_lis1, sqn_lis2,sqnb1_sum_cumsum0,sqnb2_sum_cumsum0,sqn_d1_b_cumsum0,sqn_cd_b_cumsum0
    global hf
    df_all = pd.DataFrame()
    df_all['Strategy'] = sqn_a1['Strategy'].copy()
    app = xw.apps.active
    global sht_profit
    #創建profit 工作表！
    if app is not None:
        print("!")
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        wb = xw.Book('book.xlsx')
        sht_profit = wb.sheets('profit')
    else:
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        wb = xw.Book('book.xlsx')
        sht_profit = wb.sheets('profit')

    last_row = sht_profit.range('A201').current_region.rows.count + 200
    rangevalue = sht_profit.range('B201:B{}'.format(last_row))
    df_all['name'] = rangevalue.value
    df_all['a1'] = sqn_a1['Weight'].copy()
    df_all['b1'] = sqn_b1['Weight'].copy()
    df_all['c1'] = sqn_c1['Weight'].copy()
    df_all['d1'] = sqn_d1['Weight'].copy()
    df_all['abc_sum'] = df_all.iloc[:, 2:5].sum(axis=1)
    df_all['ab_sum'] = df_all.iloc[:, 2:4].sum(axis=1)
    weights = [0.5, 0.3, 0.2]
    ab_weights = [0.7, 0.3]
    df_all['abc_weighted_sum'] = (df_all.iloc[:, 2:5] * weights).sum(axis=1)
    df_all['ab_weighted_sum'] = (df_all.iloc[:, 2:4] * ab_weights).sum(axis=1)
    df_all.to_pickle('df_all.pkl')
    df_all = pd.read_pickle('df_all.pkl')

    df_sqn_sort1 = df_all.sort_values('abc_sum', ascending=False)
    hf = int(len(df_sqn_sort1) / 2)  # 取策略的一半
    sqn_sun_best = df_sqn_sort1.iloc[:hf, :]
    sqn_sun_best = sqn_sun_best[sqn_sun_best['abc_sum'] > 0.1]

    df_sqn_sort2 = df_all.sort_values('abc_weighted_sum', ascending=False)
    sqn_wsun_best = df_sqn_sort2.iloc[:hf, :]
    sqn_wsun_best = sqn_wsun_best[sqn_wsun_best['abc_weighted_sum'] > 0.1]

    df_sqn_ab_sort = df_all.sort_values('ab_weighted_sum', ascending=False)
    sqn_ab_wsun_best = df_sqn_ab_sort.iloc[:hf, :]
    sqn_ab_wsun_best = sqn_ab_wsun_best[sqn_ab_wsun_best['ab_weighted_sum'] > 0.1]
    ##################################################
    ##################################################
    sqn_lis1 = sqn_sun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_lis1 = list(map(str, sqn_lis1))
    sqnb1 = dfa.loc[:, sqn_lis1]
    sqnb1 = sqnb1.tail(60)
    sqnb1_sum = sqnb1.groupby(sqnb1.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqnb1_sum_cumsum0 = sqnb1.sum(axis=1).cumsum()
    ##################################################
    ##################################################
    sqn_lis2 = sqn_wsun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_lis2 = list(map(str, sqn_lis2))
    sqnb2 = dfa.loc[:, sqn_lis2]
    sqnb2 = sqnb2.tail(60)
    sqnb2_sum = sqnb2.groupby(sqnb2.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqnb2_sum_cumsum0 = sqnb2.sum(axis=1).cumsum()

    ##################################################
    from joblib import dump, load
    # 寫入檔案
    dump(sqn_lis1, 'sqn_list1.joblib')
    dump(sqn_lis2, 'sqn_list2.joblib')
    dump(sqn_lis2, 'sqn_list2.joblib')
    # 讀取檔案
    s1_list = load('sqn_list1.joblib')
    s2_list = load('sqn_list2.joblib')
data_load()

def sqn_d1_cd():
    ################################################################
    # 用D1的最佳策略來評估
    global sqn_d1_lis1,sqn_cd_lis1,sqn_a1_lis1
    df_sqn_d1_sort = df_all.sort_values('d1', ascending=False)
    hf = int(len(df_sqn_d1_sort) / 2)  # 取策略的一半
    sqn_d1_best = df_sqn_d1_sort.iloc[:hf, :]
    sqn_d1_best = sqn_d1_best[sqn_d1_best['d1'] > 0.1]

    df_sqn_a1_sort = df_all.sort_values('a1', ascending=False)
    hf = int(len(df_sqn_a1_sort) / 2)  # 取策略的一半
    sqn_a1_best = df_sqn_a1_sort.iloc[:hf, :]
    sqn_a1_best = sqn_a1_best[sqn_a1_best['a1'] > 0.1]

    sqn_a1_lis1 = sqn_a1_best['Strategy'].to_list()
    sqn_a1_lis1 = list(map(str, sqn_a1_lis1))
    sqn_d1_lis1 = sqn_d1_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_d1_lis1 = list(map(str, sqn_d1_lis1))
    ######################################################################
    ######################################################################
    from joblib import dump, load
    dump(sqn_a1_lis1, 'sqn_a1_lis1.joblib')
    dump(sqn_d1_lis1, 'sqn_d1_lis1.joblib')
    # 讀取檔案
    sqn_a1_lis1 = load('sqn_a1_lis1.joblib')
    sqn_d1_lis1 = load('sqn_d1_lis1.joblib')
    ######################################################################
    ######################################################################

    sqn_d1_b = dfa.loc[:, sqn_d1_lis1]
    sqn_d1_b = sqn_d1_b.tail(60)
    sqn_d1_sum = sqn_d1_b.groupby(sqn_d1_b.columns, axis=1).sum().cumsum()
    # 计算累积总和
    sqn_d1_b_cumsum0 = sqn_d1_b.sum(axis=1).cumsum()
    ########################################################################
    ########################################################################
    # 用C1+D1的最佳策略來評估
    df_all['cd_sum'] = df_all.iloc[:, 4:6].sum(axis=1)
    weights = [0.7, 0.3]
    df_all['weighted_cd_sum'] = (df_all.iloc[:, 4:6] * weights).sum(axis=1)
    df_sqn_cd_sort = df_all.sort_values('weighted_cd_sum', ascending=False)
    hf = int(len(df_sqn_cd_sort) / 2)  # 取策略的一半
    sqn_cd_wsun_best = df_sqn_cd_sort.iloc[:hf, :]
    sqn_cd_wsun_best = sqn_cd_wsun_best[sqn_cd_wsun_best['cd_sum'] > 0.1]

    sqn_cd_lis1 = sqn_cd_wsun_best['Strategy'].to_list()
    # 使用 map() 和 str() 函数将数字列表转换为字符串列表
    sqn_cd_lis1 = list(map(str, sqn_cd_lis1))
    sqn_cd_b = dfa.loc[:, sqn_cd_lis1]
    sqn_cd_b = sqn_cd_b.tail(60)
    sqn_cd_sum = sqn_cd_b.groupby(sqn_cd_b.columns, axis=1).sum().cumsum()

    # 计算累积总和
    sqn_cd_b_cumsum0 = sqn_cd_b.sum(axis=1).cumsum()
sqn_d1_cd()

################################################################
################################################################
#PCA
def corr_f():
    # 以下是針對關連性做排序
    # 針對策略做相關性測試
    global dfb,dfaa,dfab,dfac,dfad, dfaa_cumsum,dfab_cumsum,dfac_cumsum,dfad_cumsum
    global sqn_aa_best,sqn_ab_best,sqn_ac_best,sqn_ad_best,hf,sqn_aa2ad_cumsum,sqn_aa2ad_lis
    global dfaa_cumsum_all,dfaa_cumsum_all2,sqn_a2d,sqn_a2d,sqn_a2d_lis,sqn_a2d_cumsum,sqn_a2d_cumsum_all
    dfaa=dfa.iloc[-90:,:]
    dfab=dfa.iloc[-110:-20,:]
    dfac = dfa.iloc[-130:-40, :]
    dfad = dfa.iloc[-150:-60, :]

    dfaa_cumsum =dfaa.cumsum(axis=0)
    dfab_cumsum = dfab.cumsum(axis=0)
    dfac_cumsum = dfac.cumsum(axis=0)
    dfad_cumsum = dfad.cumsum(axis=0)
    def sqn_func1(dfa):
        # 读取数据
        dfs = dfa.copy()
        # 计算每个策略的平均每笔收益和标准差
        avg_returns = dfs.mean()
        std_returns = dfs.std()

        # 定义SQN函数
        def sqn(weights, avg_returns, std_returns):
            combined_returns = np.dot(weights, avg_returns)
            combined_std = np.sqrt(np.dot(weights.T, np.dot(np.cov(dfs.T), weights)))
            sqn = np.sqrt(len(dfs)) * combined_returns / combined_std
            return -sqn  # 目标函数是最小化负的SQN值

        # 定义约束条件
        n_assets = len(dfs.columns)
        constraints = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
        bounds = [(0, 1) for i in range(n_assets)]
        # 初始权重
        weights = np.ones(n_assets) / n_assets
        # 最小化负的SQN函数，求得最优权重
        result = minimize(sqn, weights, args=(avg_returns, std_returns), method='SLSQP',
                          bounds=bounds, constraints=constraints)
        # 打印结果
        print('Optimal weights:', result.x)
        print('SQN value:', -sqn(result.x, avg_returns, std_returns))
        ws = result.x
        dfx = pd.DataFrame({'Strategy': np.arange(1, ws.shape[0] + 1),
                            'Weight': ws})
        # 選出平加總及加權加總的優良策略組合
        sqn_sort1 =  dfx.sort_values('Weight', ascending=False)
        hf = int(len(sqn_sort1) / 2)  # 取策略的一半
        sqn_best = sqn_sort1.iloc[:hf, :]
        sqn_best = sqn_best[sqn_best['Weight'] > 0.1]
        return (sqn_best)
    sqn_aa_best=sqn_func1(dfaa_cumsum)
    sqn_ab_best = sqn_func1(dfab_cumsum)
    sqn_ac_best = sqn_func1(dfac_cumsum)
    sqn_ad_best = sqn_func1(dfad_cumsum)

    sqn_aa2ad_best = pd.concat([sqn_aa_best, sqn_ab_best,sqn_ac_best,sqn_ad_best], axis=0)
    # 根據策略名稱進行分組，計算每個分組中的最大權重
    max_weight = sqn_aa2ad_best.groupby('Strategy')['Weight'].transform(max)
    # 選擇權重最大的策略
    sqn_aa2ad_best = sqn_aa2ad_best[sqn_aa2ad_best['Weight'] == max_weight]
    sqn_aa2ad_best=sqn_aa2ad_best.sort_values('Weight', ascending=False)
    sqn_aa2ad_best = sqn_aa2ad_best.iloc[:hf, :]
    sqn_aa2ad_best = sqn_aa2ad_best[sqn_aa2ad_best['Weight'] > 0.1]

    from joblib import dump, load
    # 寫入檔案
    dump(sqn_aa2ad_best, 'sqn_aa2ad_best.joblib')
    # 讀取檔案
    sqn_aa2ad_best = load('sqn_aa2ad_best.joblib')
    ########################################################################
    ########################################################################
    sqn_aa2ad_best['Strategy'] = sqn_aa2ad_best['Strategy'].astype(str)
    sqn_aa2ad_lis = sqn_aa2ad_best['Strategy'].to_list()
    sqn_aa2ad = dfaa.loc[:, sqn_aa2ad_lis]
    sqn_aa2ad['cumsum'] = sqn_aa2ad.sum(axis=1).cumsum()
    sqn_aa2ad_cumsum = sqn_aa2ad['cumsum']
    # 计算累积总和
    dfaa_cumsum_all=dfaa_cumsum.copy()
    dfaa_cumsum_all.columns = dfaa_cumsum_all.columns.astype(str)
    dfaa_cumsum_all=dfaa_cumsum_all.loc[:,list(map(str,sqn_aa_best.Strategy.tolist()))]
    dfaa_cumsum_all2=dfaa_cumsum_all.sum(axis=1)


    df_corr_a = dfaa.corr().round(2)
    df_corr_b = dfab.corr().round(2)
    df_corr_c = dfac.corr().round(2)
    df_corr_d = dfad.corr().round(2)
    ########################################################################
    ########################################################################
    sqn_ax=sqn_func1(dfaa)
    sqn_bx = sqn_func1(dfab)
    sqn_cx = sqn_func1(dfac)
    sqn_dx = sqn_func1(dfad)
    sqn_a2d = pd.concat([sqn_ax, sqn_bx, sqn_cx, sqn_dx], axis=0)
    # 根據策略名稱進行分組，計算每個分組中的最大權重
    max_weight_v = sqn_a2d.groupby('Strategy')['Weight'].transform(max)
    # 選擇權重最大的策略
    sqn_a2d = sqn_a2d[sqn_a2d['Weight'] == max_weight_v]
    sqn_a2d = sqn_a2d.sort_values('Weight', ascending=False)
    sqn_a2d = sqn_a2d.iloc[:hf, :]
    sqn_a2d = sqn_a2d[sqn_a2d['Weight'] > 0.1]
    from joblib import dump, load
    # 寫入檔案 # 讀取檔案
    dump(sqn_a2d, 'sqn_a2d.joblib')
    sqn_a2d = load('sqn_a2d.joblib')

    sqn_a2d['Strategy'] = sqn_a2d['Strategy'].astype(str)
    sqn_a2d_lis = sqn_a2d['Strategy'].to_list()
    sqn_a2d = dfaa.loc[:, sqn_a2d_lis]
    sqn_a2d['cumsum'] = sqn_a2d.sum(axis=1).cumsum()
    sqn_a2d_cumsum = sqn_a2d['cumsum']
corr_f()


def pca_f(df_, color_list, label_dict,color):
    from sklearn.preprocessing import MaxAbsScaler
    #returns = df_
    # 創建MaxAbsScaler對象
    scaler = MaxAbsScaler()
    # 對稀疏數據進行標準化
    data_scaled = scaler.fit_transform(df_)
    returns = pd.DataFrame(data_scaled, columns=df_.columns)
    pca = PCA()
    pca.fit(returns)
    #plt.tight_layout()
    # 输出主成分分析结果
    print('Explained variance ratio:', pca.explained_variance_ratio_)
    print('Principal components:', pca.components_)
    cc = pca.components_
    # 就此結果PCA繪圖
    import matplotlib.pyplot as plt
    # 绘制主成分分析结果散点图
    # 设置绘图区域的背景色为淡黄色
    plt.figure(figsize=(6, 4))
    # 绘制 sqn7 的策略点
    #plt.scatter(pca.components_[0], pca.components_[1], c='red', s=30, marker='o', facecolors='none')
    # 添加坐标轴标签
    plt.title('PCA_{}'.format(color_list), fontsize=12)
    plt.xlabel('PC1')
    plt.ylabel('PC2')
    # 添加每个交易策略的标签
    strategies = list(returns.columns)
    #color_list=sqn_a1_lis1
    # strategies = [x[:-2] if x.endswith('.0') else x for x in strategies]
    for i, strategy in enumerate(strategies):
        if strategy in color_list:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='red', s=20, marker='o', facecolors='none')
        else:
            plt.scatter(pca.components_[0][i], pca.components_[1][i], c='blue', s=10, marker='o', facecolors='none')
        #plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]), color='blue', fontsize=12)
        # 检查该策略的标签是否在label_dict中，如果是，则使用label_dict中的新标签
        if strategy in label_dict:
            label = strategy
            plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color=color, fontsize=18)
        else:
            label = strategy
            plt.annotate(label, (pca.components_[0][i], pca.components_[1][i]), color='gray', fontsize=10)
        # 显示图形
        # 設置座標軸的格線顏色
    plt.grid(color='#D3D3D3')
    # 将绘图区域的背景色改为淡黄色
    plt.gca().set_facecolor('#ffffcc')
    pp = plt.plot
    plt.tight_layout()
#dfaa##########################################################
pca_f(dfaa,sqn_a1_lis1,sqn_a1_lis1,'red')
plt.tight_layout()
# plt.show()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
if 'PCA_Tsne_Umap' not in [s.name for s in wb.sheets]:
    sht_pic = wb.sheets.add('PCA_Tsne_Umap')
else:
    wb.sheets['PCA_Tsne_Umap'].delete()
    sht_pic = wb.sheets.add('PCA_Tsne_Umap')
    sht_pic =wb.sheets['PCA_Tsne_Umap']
sht_pic.pictures.add(pic_path)
sht_pic.pictures.add(pic_path, name='dfaa', left=sht_pic.range('A1').left, top=sht_pic.range('A1').top)
################################################################
pca_f(dfab,sqn_a1_lis1,sqn_a1_lis1,'red')
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
sht_pic.pictures.add(pic_path, name='dfab', left=sht_pic.range('I1').left, top=sht_pic.range('I1').top)

pca_f(dfac,sqn_a1_lis1,sqn_a1_lis1,'red')
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
sht_pic.pictures.add(pic_path)
sht_pic.pictures.add(pic_path, name='dfad', left=sht_pic.range('Q1').left, top=sht_pic.range('Q1').top)
################################################################

########################################################################
########################################################################
#t-sne
from sklearn.manifold import TSNE
from sklearn.preprocessing import MinMaxScaler

import matplotlib.pyplot as plt
def tsne_f(df_, target_col, labels,sqnlis):
    df_t = df_.T # 转置 dataframe
    # 對數據進行MinMax標準化
    scaler = MinMaxScaler()
    df_t = pd.DataFrame(scaler.fit_transform(df_t), columns=df_t.columns)

    tsne = TSNE(n_components=2, verbose=1, perplexity=2, n_iter=5000,learning_rate=10,angle=0.1,init='pca',metric='cosine',early_exaggeration=10)
    tsne_results = tsne.fit_transform(df_t)
    df_tsne = pd.DataFrame({'X': tsne_results[:, 0], 'Y': tsne_results[:, 1], target_col: df_t.index})
    plt.figure(figsize=(10, 10))
    #sns.scatterplot(x="X", y="Y", hue=target_col, palette=sns.color_palette("hls", len(df_tsne[target_col].unique())), data=df_tsne, legend=False, alpha=1)
    for i, label in enumerate(labels):
        color = 'r' if label in sqnlis else 'b'
        fs=28 if label in sqnlis else 18
        s = 100 if label in sqnlis else 50
        plt.scatter(df_tsne.loc[i, 'X'], df_tsne.loc[i, 'Y'], c=color, s=s)
        plt.text(df_tsne.iloc[i]['X'], df_tsne.iloc[i]['Y'], label, fontsize=fs, color=color)
    plt.title('t-SNE plot')
    plt.grid(color='#D3D3D3')
    # 将绘图区域的背景色改为淡黄色
    plt.gca().set_facecolor('#f0dab6')
    plt.show()
labels=list(dfa.columns)
tsne_f(dfaa,','.join(dfa.columns),labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfaa_t', left=sht_pic.range('A71').left, top=sht_pic.range('A71').top)
pic.height *= 0.77
pic.width *= 0.71
################################################################
tsne_f(dfab,','.join(dfa.columns),labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfab_t', left=sht_pic.range('I71').left, top=sht_pic.range('I71').top)
pic.height *= 0.77
pic.width *= 0.71

################################################################
tsne_f(dfac,','.join(dfa.columns),labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfac_t', left=sht_pic.range('Q71').left, top=sht_pic.range('Q71').top)
pic.height *= 0.77
pic.width *= 0.71
################################################################
################################################################
################################################################
################################################################
import umap.umap_ as umap
def umap_f(df_, target_col, labels,sqnlis):
    df_t = df_.T # 转置 dataframe
    # 對數據進行MinMax標準化
    scaler = MinMaxScaler()
    df_t = pd.DataFrame(scaler.fit_transform(df_t), columns=df_t.columns)
    umap_results = umap.UMAP(n_neighbors=500, min_dist=1, n_components=10, repulsion_strength=0.1, learning_rate=0.01).fit_transform(df_t)
    df_umap = pd.DataFrame({'X': umap_results[:, 0], 'Y': umap_results[:, 1], target_col: df_t.index})
    plt.figure(figsize=(8, 8))
    #sns.scatterplot(x="X", y="Y", hue=target_col, palette=sns.color_palette("hls", len(df_umap[target_col].unique())), data=df_umap, legend=False, alpha=0.9)
    for i, label in enumerate(labels):
        color = 'red' if label in sqnlis else "blue"
        fs=26 if label in sqnlis else 16
        s = 100 if label in sqnlis else 50
        plt.scatter(df_umap.loc[i, 'X'], df_umap.loc[i, 'Y'], c=color, s=s)
        plt.text(df_umap.iloc[i]['X'], df_umap.iloc[i]['Y'], label, fontsize=fs, color=color)
    plt.title('UMAP plot')
    plt.grid(color='#D3D3D3')
    # 将绘图区域的背景色改为淡黄色
    plt.gca().set_facecolor('#b6d4f0')
    plt.show()
labels=list(dfaa.columns)
umap_f(dfaa, ','.join(dfa.columns), labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfaa_u', left=sht_pic.range('A45').left, top=sht_pic.range('A45').top)
pic.height *= 0.85
pic.width *= 0.78
################################################################
################################################################
umap_f(dfab, ','.join(dfa.columns), labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfab_u', left=sht_pic.range('i45').left, top=sht_pic.range('i45').top)
pic.height *= 0.85
pic.width *= 0.78
################################################################
################################################################
umap_f(dfac, ','.join(dfa.columns), labels,sqn_a1_lis1)
plt.tight_layout()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='dfac_u', left=sht_pic.range('q45').left, top=sht_pic.range('q45').top)
pic.height *= 0.85
pic.width *= 0.78


from pacmap import PaCMAP
import pacmap
import numpy as np
import matplotlib.pyplot as plt
from sklearn.preprocessing import MinMaxScaler

def pacmap_f(df_, target_col, labels, sqnlis):
    # 轉換為 numpy array 並標準化
    X = df_.values.T
    X = MinMaxScaler().fit_transform(X)

    # 初始化 paCMAP 模型
    # 如果 n_neighbors 設定為 None，則會使用預設值（50 或資料點數的 10%）
    #embedding = pacmap.PaCMAP(n_components=2, n_neighbors=None, MN_ratio=0.5, FP_ratio=2.0,verbose=True)
    embedding = pacmap.PaCMAP(n_components=2, n_neighbors=5, MN_ratio=0.8, FP_ratio=2.0, verbose=True)

    # 轉換資料
    X_transformed = embedding.fit_transform(X, init="pca")

    # 將轉換後的資料轉換為 DataFrame，並加入目標欄位
    df_pacmap = pd.DataFrame({'X': X_transformed[:, 0], 'Y': X_transformed[:, 1], target_col: df_.columns})

    # 繪製 paCMAP 圖表
    plt.figure(figsize=(8, 8))
    for i, label in enumerate(labels):
        color = 'red' if label in sqnlis else "blue"
        fs = 20 if label in sqnlis else 10
        s = 50 if label in sqnlis else 20
        plt.scatter(df_pacmap.loc[i, 'X'], df_pacmap.loc[i, 'Y'], c=color, s=s)
        plt.text(df_pacmap.iloc[i]['X'], df_pacmap.iloc[i]['Y'], label, fontsize=fs, color=color)
    plt.title('paCMAP plot')
    plt.grid(color='#D3D3D3')
    # 将绘图区域的背景色改为淡黄色
    plt.gca().set_facecolor('#f0d9b6')
    plt.show()
labels = list(dfaa.columns)
pacmap_f(dfaa, ','.join(dfa.columns), labels, sqn_a1_lis1)
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='pacam_1', left=sht_pic.range('A18').left, top=sht_pic.range('A18').top)
pic.height *= 0.87
pic.width *= 0.80

pacmap_f(dfab, ','.join(dfa.columns), labels, sqn_a1_lis1)
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='pacam_2', left=sht_pic.range('I18').left, top=sht_pic.range('I18').top)
pic.height *= 0.87
pic.width *= 0.80

pacmap_f(dfac, ','.join(dfa.columns), labels, sqn_a1_lis1)
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
pic=sht_pic.pictures.add(pic_path, name='pacam_3', left=sht_pic.range('Q18').left, top=sht_pic.range('Q18').top)
pic.height *= 0.87
pic.width *=0.80





#用cum_sum來做PCA_tsne_umap_pacmap
def dfaa_cumsum_():
    if 'PCA_Tsne_Umap_cumcum' not in [s.name for s in wb.sheets]:
        sht_pic = wb.sheets.add('PCA_Tsne_Umap_cumcum')
    else:
        wb.sheets['PCA_Tsne_Umap_cumcum'].delete()
        sht_pic = wb.sheets.add('PCA_Tsne_Umap_cumcum')
        sht_pic =wb.sheets['PCA_Tsne_Umap_cumcum']
    # dfaa CumSum ##########################################################
    pca_f(dfaa_cumsum, sqn_a1_lis1, sqn_a1_lis1,'red')
    plt.tight_layout()
    # plt.show()
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    sht_pic.pictures.add(pic_path, name='dfaa', left=sht_pic.range('A1').left, top=sht_pic.range('A1').top)
    labels = list(dfaa_cumsum.columns)
    tsne_f(dfaa_cumsum, ','.join(dfa.columns), labels, sqn_a1_lis1)
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    pic = sht_pic.pictures.add(pic_path, name='dfaa_t', left=sht_pic.range('J1').left, top=sht_pic.range('J1').top)
    pic.height *= 0.77
    pic.width *= 0.71
    labels = list(dfaa_cumsum.columns)
    umap_f(dfaa_cumsum, ','.join(dfa.columns), labels, sqn_a1_lis1)
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    pic = sht_pic.pictures.add(pic_path, name='dfaa_u', left=sht_pic.range('A25').left, top=sht_pic.range('A25').top)
    pic.height *= 0.85
    pic.width *= 0.78

    pacmap_f(dfaa_cumsum, ','.join(dfa.columns), labels, sqn_a1_lis1)
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    pic = sht_pic.pictures.add(pic_path, name='pacam_3K', left=sht_pic.range('J25').left, top=sht_pic.range('J25').top)
    pic.height *= 0.87
    pic.width *= 0.80
dfaa_cumsum_()

################################################################
################################################################
#波動SQN、累積SQN及第一段（90日）優良策略至少出現2次的。
# 将三个列表合并为一个集合
set1 = set(sqn_a1_lis1) #A段best
set2 = set(sqn_aa2ad_lis) #累積SQN_Best
set3 = set(sqn_a2d_lis)#波動A:D
union_set = set1.union(set2, set3)
# 找到至少出现两次的数字
set_all = []
for num in union_set:
    if (sqn_a1_lis1.count(num) + sqn_aa2ad_lis.count(num) + sqn_a2d_lis.count(num)) >= 2:
        set_all.append(num)
#將三類hit至少兩次的策略加總損益
set_all_collect = dfaa.loc[:, set_all]
set_all_collect['cumsum'] = set_all_collect.sum(axis=1).cumsum()
set_all_collect_cumsum = set_all_collect['cumsum']



########################################################################
########################################################################

##################################################
# 绘制折线图
#ax = sqnb1_sum.plot.line(linewidth=0.5, alpha=0.15)
plt.figure(figsize=(10, 8))
#sns.lineplot(x=sqnb1_sum_cumsum0.index, y=sqnb1_sum_cumsum0[:], data=sqnb1_sum_cumsum0, linestyle="-", color="red", label="abc_combined-{}")
#sns.lineplot(x=sqnb2_sum_cumsum0.index, y=sqnb2_sum_cumsum0[:], data=sqnb2_sum_cumsum0, linestyle="--", color="gray", label="abc_Wcombined-{}")
sns.lineplot(x=sqn_aa2ad_cumsum.index, y=sqn_aa2ad_cumsum[:], data=sqn_aa2ad_cumsum, linestyle=":", color="green", label="ABCD_Cum-{}".format(sqn_aa2ad_lis))
sns.lineplot(x=dfaa_cumsum_all2.index, y=dfaa_cumsum_all2[:], data=dfaa_cumsum_all2, linestyle="-", color="red", label="Only_A_{}".format(sqn_a1_lis1))
sns.lineplot(x=sqn_a2d_cumsum.index, y=sqn_a2d_cumsum[:], data=sqn_a2d_cumsum, linestyle="--", color="blue", label="ABCD_Vola-{}".format(sqn_a2d_lis))
sns.lineplot(x=set_all_collect_cumsum.index, y=set_all_collect_cumsum[:], data=set_all_collect_cumsum, linestyle="-", color="orange", label="Vola+Cumsum+SecA-{}".format(set_all))


# 繪製垂直線
last_date = sqnb1_sum_cumsum0.index[-1]
line_date = last_date - pd.DateOffset(days=40)
plt.axvline(x=line_date, linestyle="--", color="red", linewidth=0.3)


plt.xlabel("Date", fontdict={'fontsize': 10})
plt.ylabel("Total Return", fontdict={'fontsize': 10})
plt.title("SQN_rolling", fontdict={'fontsize': 18})
# 在 x 轴上绘制垂直线
#ax.axvline(x=70, color='r')
plt.legend(loc='upper left', borderaxespad=0., fontsize='large', fancybox=True, edgecolor='navy', framealpha=0.2,
           handlelength=1.5, handletextpad=0.5, borderpad=0.5, labelspacing=0.5)

plt.rcParams['xtick.labelsize']=8
plt.rcParams['ytick.labelsize']=8
plt.tight_layout()
plt.show()
plt.savefig('temp.png')
plt.close()



##################################################
##################################################
global cum_vola_A
if 'cum_vola_A' in wb.sheet_names:
    wb.sheets['cum_vola_A'].delete()
    sht_cum_vola_A = wb.sheets.add('cum_vola_A')
else:
    sht_cum_vola_A = wb.sheets.add('cum_vola_A')

#圖1：折線圖
pic_path = (os.path.join(os.getcwd(), "temp.png"))
#sht_cum_vola_A.pictures.add(pic_path)
pic =sht_cum_vola_A.pictures.add(pic_path, name='comp_line', left=sht_cum_vola_A.range('A2').left, top=sht_cum_vola_A.range('A2').top)
pic.height *= 0.75
pic.width *= 0.70
################################################################
#圖2，dfaa 累積圖
ws_source = wb.sheets['PCA_Tsne_Umap_cumcum']
ws_target = wb.sheets['cum_vola_A']
# 選擇工作表1和圖片1
sht1 = wb.sheets['PCA_Tsne_Umap_cumcum']
pic1 = sht1.pictures['dfaa']
# 選擇工作表2
sht2 = wb.sheets['cum_vola_A']
# 複製圖片並粘貼到工作表2
pic1.api.Copy()
sht2.api.Paste()
rng=sht2.range("I2")
ws_target.pictures[-1].top = rng.top
ws_target.pictures[-1].left = rng.left
######################################################
######################################################
#圖3：
pca_f(dfaa, sqn_a2d_lis, sqn_a2d_lis,'blue')
plt.tight_layout()
# plt.show()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))

ws_target.pictures.add(pic_path, name='dfaa_V', left=ws_target.range('R2').left, top=ws_target.range('R2').top)

######################################################
######################################################
#圖4：
pca_f(dfaa, sqn_aa2ad_lis, sqn_aa2ad_lis,'green')
plt.tight_layout()
# plt.show()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
ws_target.pictures.add(pic_path, name='a2d_cum', left=ws_target.range('a23').left, top=ws_target.range('a23').top)

######################################################
######################################################
#圖5：
pca_f(dfaa, set_all, set_all,'orange')
plt.tight_layout()
# plt.show()
plt.savefig('temp.png')
plt.close()
pic_path = (os.path.join(os.getcwd(), "temp.png"))
ws_target.pictures.add(pic_path, name='a2d_V_cum_secA', left=ws_target.range('J23').left, top=ws_target.range('J23').top)

ws_target.range('A1').value= "SQN權益比較圖"
ws_target.range('i1').value= "SectionA__累積損益數據_最佳SQN圖"
ws_target.range('R1').value= "Section_A2D_每日波動數據_最佳SQN圖"
ws_target.range('a22').value= "Section_A2D_累積損益數據_最佳SQN圖"
ws_target.range('I22').value= "SecA+A2D波動+A2D累積_數據最佳SQN圖"

ws_target.range('A1').color = (255, 0, 0)
ws_target.range('I22').color = (237, 175, 31)
ws_target.range('I1').color = (255, 0, 0)
ws_target.range('R1').color = (0, 0, 255)
ws_target.range('A22').color = (0, 255, 0)


wb.save('book.xlsx')







'''
#################################################
################################################################

#選出平加總及加權加總的優良策略組合
df_sqn_sort1=df_all.sort_values('abc_sum', ascending=False)
hf=int(len(df_sqn_sort1)/2)#取策略的一半
sqn_sun_best=df_sqn_sort1.iloc[:hf,:]
sqn_sun_best= sqn_sun_best[sqn_sun_best['abc_sum']>0.05]

df_sqn_sort2=df_all.sort_values('weighted_sum', ascending=False)
sqn_wsun_best=df_sqn_sort2.iloc[:hf,:]
sqn_wsun_best= sqn_sun_best[sqn_sun_best['weighted_sum']>0.05]

##################################################
##################################################
sqn_lis1=sqn_sun_best['Strategy'].to_list()
# 使用 map() 和 str() 函数将数字列表转换为字符串列表
sqn_lis1 = list(map(str, sqn_lis1))
sqnb1=dfa.loc[:,sqn_lis1]
#sqnb1=sqnb1.tail(60)
sqnb1_sum = sqnb1.groupby(sqnb1.columns, axis=1).sum().cumsum()
# 计算累积总和
sqnb1_sum_cumsum0 = sqnb1.sum(axis=1).cumsum()
##################################################
##################################################
sqn_lis2=sqn_wsun_best['Strategy'].to_list()
# 使用 map() 和 str() 函数将数字列表转换为字符串列表
sqn_lis2 = list(map(str, sqn_lis2))
sqnb2=dfa.loc[:,sqn_lis2]
#sqnb2=sqnb2.tail(60)
sqnb2_sum = sqnb2.groupby(sqnb2.columns, axis=1).sum().cumsum()
# 计算累积总和
sqnb2_sum_cumsum0 = sqnb2.sum(axis=1).cumsum()

##################################################
##################################################
# 绘制折线图
ax = sqnb1_sum.plot.line(linewidth=0.5, alpha=0.3)
sqnb1_sum_cumsum0.plot.line(ax=ax, linewidth=1, color='red',alpha=0.4)
sqnb2_sum_cumsum0.plot.line(ax=ax, linewidth=1, color='blue',alpha=0.4)
plt.show()
##################################################
##################################################
'''

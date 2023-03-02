import os
import xlwings
import  pandas as pd
import warnings
from sklearn.decomposition import PCA
import difflib
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.optimize import minimize
import numpy as np
import networkx as nx

def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col

# 打開一個新的Excel文件
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=False
wb = xw.Book()  # this will open a new workbook
#wb = xw.Book('Book1.xlsx')
#wb=app.books.open('Book1.xlsx')
sht=wb.sheets(1)
# 從當前工作表中取得第一個儲存格（A2）
cell = sht.range('A2')
# 使用Python中的datetime模組來得到當前月份的資訊
import datetime
#第一步基本資料處理
def step1_f():
    # 取得當前月份的開始日期
    # start_date = datetime.date(datetime.date.today().year, datetime.date.today().month, 1)
    # 將開始日期輸入到A1儲存格中
    dayrng = 90
    start_date = datetime.date.today() - datetime.timedelta(days=dayrng)
    cell.value = start_date
    # 將開始日期往下移動一格，並輸出相對應的日期資訊
    sht.cells(1, 1).value = '日期'
    global last_cell
    for i in range(1, dayrng):
        cell.offset(row_offset=i, column_offset=0).value = start_date + datetime.timedelta(days=i)
        # 找到記錄XLS"Book1.xls"的最後一行
        # 找到最後一行

    last_cell = sht.range('A2').current_region.rows.count

    last_cell = last_cell + 1
    print(last_cell)

    # 先找到目錄中    "出貨通知單開頭    .xlsx"    的檔案____
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
                wb1 = app.books.open(fd + "/" + order_name)
                sht1 = wb1.sheets('交易明細')

                # 將 sheet 內容轉換為 dataframe
                df = sht1.range('B3').options(pd.DataFrame, expand='table').value
                # 移除 column_name 為空的 row
                df.dropna(subset=[r"獲利(¤)"], inplace=True)
                df['日期'] = pd.to_datetime(df['日期']).dt.date

                # 90日期中，一個一個日期字串拉出來 篩選 並取得加總：
                sht.cells(1, c + 1).value = c
                d0 = 1
                for dstr in sht.range('A2:A{}'.format(last_cell)).value:
                    profit = df['獲利(¤)'][(df['日期'] == pd.to_datetime(dstr))].sum()
                    sht.cells(d0 + 1, fn + 1).value = profit
                    d0 = d0 + 1
                wb1.close()
                c = c + 1

            sht.cells(95 + fn, 1).value = c - 1
            sht.cells(95 + fn, 2).value = filename
            fn = fn + 1
            # Quit Excel application #
step1_f()





#第二步，生成Corr_1 Corr_2
def corr_f():
    # 以下是針對關連性做排序
    last_column = sht.range('A2').current_region.columns.count
    rng_all = sht.range('A1:{}'.format(num_to_col(last_column) + str(last_cell)))
    global dfa
    dfa = pd.DataFrame(rng_all.value, columns=rng_all.value[0])
    # 將 '日期' 列轉換為索引

    dfa = df = dfa.drop(dfa.index[0])
    dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
    dfa = dfa.set_index('日期')

    dfa = dfa.rename(columns=lambda x: str(x).strip())
    # 針對策略做相關性測試
    global dfb
    dfb = dfa.corr().round(2)

    sht_corr = wb.sheets.add('corr_1')
    sht_corr.range('A1').options(index=True, header=True).value = dfb
    # 將列按照相關性排序
    global dfc
    dfc = dfb['2.0'].sort_values(ascending=False).index
    # 重新排列 DataFrame 中的列
    global dfd
    dfd = dfa[dfc]
    sht_sortd = wb.sheets.add('re_sortd')
    sht_sortd.range('A1').options(index=True, header=True).value = dfd
    global dfe
    dfe = dfd.corr().round(2)
    sht_corr_2 = wb.sheets.add('corr2')
    sht_corr_2.range('A1').options(index=True, header=True).value = dfe
corr_f()

def MP_mmse_f():
    import numpy as np
    # 讀取交易明細資料
    # 轉換為numpy array
    returns = np.array(dfa)
    mean_returns = np.mean(returns, axis=0)
    # 計算各個策略的波動率
    volatility = np.std(returns, axis=0)
    # 計算各個策略的最大回撤
    cum_returns = np.cumsum(returns, axis=0)
    max_drawdowns = np.zeros(dfa.shape[1])
    for i in range(dfa.shape[1]):
        j = np.argmax(cum_returns[:, i] - np.maximum.accumulate(cum_returns[:, i]))
        if j == 0:
            max_drawdowns[i] = 0
        else:
            max_drawdowns[i] = cum_returns[j, i] - cum_returns[j - 1, i]
    # 印出各個策略的平均回報率、波動率和最大回撤
    print('Mean returns:', mean_returns)
    print('Volatility:', volatility)
    print('Max drawdowns:', max_drawdowns)
    # 計算最優權重
    cov = np.cov(returns, rowvar=False)
    inv_cov = np.linalg.inv(cov)
    ones = np.ones(dfa.shape[1])
    global w
    w = inv_cov @ ones / (ones.T @ inv_cov @ ones)
    # 印出最優權重
    print('Optimal weights:', w)


    dfx = pd.DataFrame({'Strategy': np.arange(1, w.shape[0] + 1),
                        'Weight': w})
    # 使用seaborn.barplot()繪製條形圖
    plt.figure(figsize=(8, 6))
    ax = sns.barplot(x='Strategy', y='Weight', data=dfx)

    # 在條形上添加標籤
    for p in ax.patches:
        ax.annotate(f'{p.get_height():.2f}', (p.get_x() + p.get_width() / 2, p.get_height() + 0.01),
                    ha='center')

    # 設置圖表標題和坐標軸標籤
    ax.set_title('MaxProfit/MMSE Optimal weights')
    ax.set_xlabel('Strategy index')
    ax.set_ylabel('Weight')
    # 顯示圖形
    # plt.show()
    plt.tight_layout()
    sht_pic = wb.sheets.add('MP_MMSE')
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    # 删除临时文件
    os.remove('temp.png')
MP_mmse_f()

def heat_map_f():
    import seaborn as sns
    sns.set(font_scale=1)
    plt.figure(figsize=(8, 6))
    sns.heatmap(dfe, annot=True, cmap='coolwarm')
    plt.tight_layout()
    plt.savefig('temp.png')
    sht_pic = wb.sheets.add('HeatMap')

    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    plt.close()
    # 删除临时文件
    os.remove('temp.png')
    # 召入散點矩陣圖
    import seaborn as sns
    # 绘制散点矩阵图
    # sns.pairplot(dfe, diag_kind='hist')
    # 显示图像
    # plt.show()
    plt.tight_layout()
    plt.close()
    from scipy.cluster import hierarchy
    # dendrogram = hierarchy.dendrogram(hierarchy.linkage(dfe))
heat_map_f()

def networtx_f():
    # 创建图
    G = nx.Graph()

    # 添加节点
    for col in dfe.columns:
        G.add_node(col)
    # 添加边
    for i in range(len(dfe.columns)):
        for j in range(i + 1, len(dfe.columns)):
            if abs(dfe.iloc[i, j]) > 0.3:
                G.add_edge(dfe.columns[i], dfe.columns[j], weight=dfe.iloc[i, j])
    # 绘制网络图
    pos = nx.circular_layout(G)
    # 繪製節點
    nx.draw_networkx_nodes(G, pos, node_size=400, node_color='lightblue', node_shape='o', linewidths=1)
    # 繪製節點標籤
    labels = {node: node for node in G.nodes()}
    nx.draw_networkx_labels(G, pos, labels, font_color='red')
    # 繪製邊
    nx.draw_networkx_edges(G, pos, style="dashed")

    # 繪製邊權重標籤
    nx.draw_networkx_edge_labels(G, pos, edge_labels={(u, v): round(d["weight"], 1) for u, v, d in G.edges(data=True)})
    # 顯示圖形
    # plt.show()
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.savefig('nx.png')
    plt.close()
    sht_pic = wb.sheets.add('NetworkX')
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
networtx_f()

def mmse_f():
    returns = dfa
    pct_df = pd.DataFrame(index=dfa.index, columns=dfa.columns)
    for col in dfa.columns:
        pct_df[col] = dfa[col] / 100000 * 100
    # 给定权重，求组合收益率、标准差、夏普比率
    # 定义无风险收益率
    rf = 0.02
    # 获取股票平均收益率
    mean_returns = dfa.mean()
    # 获取股票收益率的方差协方差矩阵
    cov_matrix = dfa.cov()
    # 定义资产数量
    number_assets = dfa.shape[1]

    # 目标收益率约束条件
    target_return = 0.5

    # 最小化方差的目标函数
    def portfolio_volatility(weights, cov_matrix):
        return np.sqrt(np.dot(weights.T, np.dot(cov_matrix, weights)))

    # 约束条件
    def constraint(weights, mean_returns, target_return):
        return np.sum(mean_returns * weights) - target_return

    # 初始权重向量
    n_assets = len(mean_returns)
    weights_0 = np.ones(n_assets) / n_assets

    # 定义边界条件
    bounds = tuple((0, 1) for _ in range(n_assets))

    # 定义约束条件
    cons = ({'type': 'eq', 'fun': constraint, 'args': (mean_returns, target_return)})

    # 调用 minimize 函数进行投资组合优化
    opt_result = minimize(portfolio_volatility, weights_0, args=(cov_matrix,),
                          method='SLSQP', bounds=bounds, constraints=cons)

    # 输出优化结果
    print(opt_result.x)

    ####################################################################
    ####################################################################
    global wx
    wx = opt_result.x / sum(opt_result.x)
    df = pd.DataFrame({'Strategy': np.arange(1, wx.shape[0] + 1),
                       'Weight': wx})
    # 使用seaborn.barplot()繪製條形圖
    plt.figure(figsize=(8, 6))
    plt.rcParams.update({'font.size': 12})
    ax = sns.barplot(x='Strategy', y='Weight', data=df)

    # 在條形上添加標籤
    for pa in ax.patches:
        ax.annotate(f'{pa.get_height():.2f}', (pa.get_x() + pa.get_width() / 2, pa.get_height() + 0.01),
                    ha='center', color='black', fontsize=10)

    # 設置圖表標題和坐標軸標籤

    ax.set_title('Optimal Least Square Error', fontdict={'fontsize': 14})
    ax.set_xlabel('Strategy index', fontdict={'fontsize': 10})
    ax.set_ylabel('Weight', fontdict={'fontsize': 10})
    # 顯示圖形
    # plt.show()
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.close()
    sht_pic = wb.sheets.add('MMSE Optimal Bar')
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    print("全部策略：" + str(c - 1) + " 支")
mmse_f()

def sqn_f():
    # SQN##################################SQN#################################
    # SQN##################################SQN#################################
    # SQN##################################SQN#################################
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

    # 使用seaborn.barplot()繪製條形圖
    plt.figure(figsize=(8, 6))
    plt.rcParams.update({'font.size': 12})
    ax = sns.barplot(x='Strategy', y='Weight', data=dfx)
    # 在條形上添加標籤
    for pa in ax.patches:
        ax.annotate(f'{pa.get_height():.2f}', (pa.get_x() + pa.get_width() / 2, pa.get_height() + 0.01),
                    ha='center', color='black', fontsize=10)

    # 設置圖表標題和坐標軸標籤

    ax.set_title('SQN OPT Portfolio', fontdict={'fontsize': 18})
    ax.set_xlabel('Strategy index', fontdict={'fontsize': 10})
    ax.set_ylabel('Weight', fontdict={'fontsize': 10})
    # 顯示圖形
    # plt.show()
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.close()
    sht_sqn = wb.sheets.add('SQN Optimal')
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_sqn.pictures.add(pic_path)

    # SQN30##################################SQN30#################################
    # SQN30##################################SQN30#################################
    # SQN30##################################SQN30#################################
    # 读取数据
    dfs30 = dfa.iloc[-31:-1, :]
    # 计算每个策略的平均每笔收益和标准差
    avg_returns30 = dfs30.mean()
    std_returns30 = dfs30.std()

    # 定义SQN函数
    def sqn30(weights30, avg_returns30, std_returns30):
        combined_returns30 = np.dot(weights30, avg_returns30)
        combined_std30 = np.sqrt(np.dot(weights30.T, np.dot(np.cov(dfs30.T), weights30)))
        sqn30 = np.sqrt(len(dfs30)) * combined_returns30 / combined_std30
        return -sqn30  # 目标函数是最小化负的SQN值

    # 定义约束条件
    n_assets30 = len(dfs30.columns)
    constraints30 = [{'type': 'eq', 'fun': lambda x: np.sum(x) - 1}]
    bounds30 = [(0, 1) for i in range(n_assets30)]
    # 初始权重
    weights30 = np.ones(n_assets30) / n_assets30
    # 最小化负的SQN函数，求得最优权重
    result30 = minimize(sqn30, weights30, args=(avg_returns30, std_returns30), method='SLSQP',
                        bounds=bounds30, constraints=constraints30)
    # 打印结果
    print('Optimal weights:', result30.x)
    print('SQN value:', -sqn30(result30.x, avg_returns30, std_returns30))
    ws30 = result30.x
    dfx30 = pd.DataFrame({'Strategy': np.arange(1, ws30.shape[0] + 1),
                          'Weight': ws30})

    #############################################################################
    # 使用seaborn.barplot()繪製條形圖
    plt.figure(figsize=(8, 6))
    plt.rcParams.update({'font.size': 12})
    ax = sns.barplot(x='Strategy', y='Weight', data=dfx30)
    # 在條形上添加標籤
    for pa in ax.patches:
        ax.annotate(f'{pa.get_height():.2f}', (pa.get_x() + pa.get_width() / 2, pa.get_height() + 0.01),
                    ha='center', color='black', fontsize=10)

    # 設置圖表標題和坐標軸標籤
    ax.set_title('SQN30 OPT Portfolio', fontdict={'fontsize': 18})
    ax.set_xlabel('Strategy index', fontdict={'fontsize': 10})
    ax.set_ylabel('Weight', fontdict={'fontsize': 10})
    # 顯示圖形
    # plt.show()
    plt.tight_layout()
    plt.savefig('temp.png')
    plt.close()
    # sht_sqn = wb.sheets.add('SQN Optimal')
    pic_path1 = (os.path.join(os.getcwd(), "temp.png"))
    # 插入图片并设置位置
    sht_sqn.pictures.add(pic_path1, name='MyPic1', left=sht_sqn.range('L1').left, top=sht_sqn.range('L1').top)

    # 畫優化組合的權益曲線########################################################################
    # 畫優化組合的權益曲線########################################################################
    hf = int((w.shape[0]) / 2)
    mmse_df = pd.DataFrame({'Strategy': np.arange(1, w.shape[0] + 1), 'Weight': wx})
    mmse_df = mmse_df.sort_values(by='Weight', ascending=False)
    global mmse7
    mmse7 = mmse_df.iloc[0:hf]
    mmse7 = mmse7[mmse7['Weight'] > 0.01]

    mp_mmse_df = pd.DataFrame({'Strategy': np.arange(1, w.shape[0] + 1), 'Weight': w})
    mp_mmse_df = mp_mmse_df.sort_values(by='Weight', ascending=False)
    global mp_mmse7
    mp_mmse7 = mp_mmse_df.iloc[0:hf]
    mp_mmse7 = mp_mmse7[mp_mmse7['Weight'] > 0.01]

    sqn_df = pd.DataFrame({'Strategy': np.arange(1, ws.shape[0] + 1), 'Weight': ws})
    sqn_df = sqn_df.sort_values(by='Weight', ascending=False)
    global sqn7
    sqn7 = sqn_df.iloc[0:hf]
    sqn7 = sqn7[sqn7['Weight'] > 0.01]
    # 将所有列名中的 ".0" 替换为 ""
    dfa_s = dfa.copy()
    dfa_s.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    s7 = sqn7.Strategy
    global lis_s7
    lis_s7 = s7.tolist()
    lis_s7a = [str(x) for x in lis_s7]
    global s7_df
    s7_df = dfa_s.loc[:, lis_s7a].copy()
    # 将每一行加总并存储为一个新的 Series
    s7_sum = s7_df.sum(axis=1)
    # 将新的 Series 添加到 DataFrame 中
    s7_df['s7 Sum'] = s7_sum
    # 計算B欄逐行加總，並新增一個欄位'C'保存結果
    s7_df['cum_Sum'] = s7_df['s7 Sum'].cumsum()

    dfa_s = dfa.copy()
    dfa_s.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    s7 = sqn7.Strategy
    lis_s7 = s7.tolist()
    lis_s7a = [str(x) for x in lis_s7]
    s7_df = dfa_s.loc[:, lis_s7a].copy()
    # 将每一行加总并存储为一个新的 Series
    s7_sum = s7_df.sum(axis=1)
    # 将新的 Series 添加到 DataFrame 中
    s7_df['s7 Sum'] = s7_sum
    # 計算B欄逐行加總，並新增一個欄位'C'保存結果
    s7_df['cum_Sum'] = s7_df['s7 Sum'].cumsum()

    ##############################################################################
    # SQN30##################################SQN30#################################
    sqn30_df = pd.DataFrame({'Strategy': np.arange(1, ws30.shape[0] + 1), 'Weight': ws30})
    sqn30_df = sqn30_df.sort_values(by='Weight', ascending=False)
    global sqn7_30
    sqn7_30 = sqn30_df.iloc[0:hf]
    sqn7_30 = sqn7_30[sqn7_30['Weight'] > 0.01]

    dfa_s30 = dfa.copy()
    dfa_s30.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    s7_30 = sqn7_30.Strategy
    global lis_s7_30
    lis_s7_30 = s7_30.tolist()
    lis_s7a_30 = [str(x) for x in lis_s7_30]
    global s7_df_30
    s7_df_30 = dfa_s.loc[:, lis_s7a_30].copy()
    # 将每一行加总并存储为一个新的 Series
    s7_sum_30 = s7_df_30.sum(axis=1)
    # 将新的 Series 添加到 DataFrame 中
    s7_df_30['s7_30 Sum'] = s7_sum_30
    # 計算B欄逐行加總，並新增一個欄位'C'保存結果
    s7_df_30['cum_Sum'] = s7_df_30['s7_30 Sum'].cumsum()
    ##############################################################################
    # SQN30##################################SQN30#################################
sqn_f()

def sqn_last10d_f():
    import pandas as pd
    ############################################
    # 存本週最優SQN
    import csv
    # 開啟Excel文件，並選擇要導出的工作表
    ws = wb.sheets['工作表1']
    # 選擇要導出的範圍
    last_row_s = ws.range('A96').current_region.rows.count + 95
    range_to_export = ws.range('A96:B{}'.format(last_row_s))

    # 將範圍轉換為列表
    range_values = range_to_export.value
    global lis_sqn
    lis_sqn = []
    for ai in lis_s7:
        for aj in range_values:
            if ai == aj[0]:
                lis_sqn.append(aj)
    # 將列表轉換為Pandas DataFrame
    dfk = pd.DataFrame(lis_sqn, columns=['num', 'filename'])

    # 將DataFrame保存為Excel文件
    from datetime import datetime
    today = datetime.today()
    date_string = today.strftime('%Y%m%d')
    dfk.to_excel('sqn_{}.xlsx'.format(date_string), index=False)
    # dfkk = pd.read_excel('sqn.xlsx', sheet_name='Sheet1')
    import os
    import fnmatch
    import datetime
    import pandas as pd

    # 設定檔案名稱格式和目錄路徑
    filename_pattern = "sqn_*.xlsx"
    directory = "."

    # 獲取當前日期和10天前的日期
    today = datetime.datetime.today()
    ten_days_ago = today - datetime.timedelta(days=10)
    # 獲取符合條件的檔案列表
    matching_files = fnmatch.filter(os.listdir(directory), filename_pattern)

    # 過濾出10天前的檔案和最舊的檔案
    files_10_days_ago = [f for f in matching_files if
                         datetime.datetime.fromtimestamp(os.path.getctime(f)) < ten_days_ago]
    files_sorted_by_date = sorted(matching_files, key=lambda f: os.path.getctime(f))

    # 獲取需要讀取的檔案名稱
    if len(files_10_days_ago) > 0:
        file_to_read = files_10_days_ago[-1]
    else:
        file_to_read = files_sorted_by_date[0]
    # 讀取檔案為DataFrame
    dfkk = pd.read_excel(file_to_read)
    # 顯示DataFrame
    print(dfkk)

    # 把sqn list 的excel 讀出來
    # 上段完成！
    # 用for 找到dfkk中, wb 工作表一 95-11?  rangevalue的list
    # 再把rangevalue 符合的檔名的序號記到list[]  lis.append
    lis_sqn_old = []
    for bi in dfkk['filename']:
        for bj in range_values:
            matcher = difflib.SequenceMatcher(None, bi, bj[1])
            print(matcher.ratio())
            if matcher.ratio() > 0.95:
                lis_sqn_old.append(bj[0])

    # 最後重dfa中 把這串10天前舊的sqn組合的表現調出來
    dfak = dfa.copy()
    lis_sqn_old = [str(num) for num in lis_sqn_old]
    dfak2 = dfak[lis_sqn_old]
    dfak2['cumsum'] = dfak2.sum(axis=1).cumsum()

    # 做sun及cumsum()後 畫到plot中

    '''
    # 指定要保存的CSV文件名
    with open('sqn.csv', mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(lis_csv[0])
        writer.writerow(lis_csv[1])
    '''
    ############################################

    # 将所有列名中的 ".0" 替换为 ""
    dfa_c = dfa.copy()
    dfa_c.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    m7 = mmse7.Strategy
    lis_m7 = m7.tolist()
    lis_m7a = [str(x) for x in lis_m7]
    m7_df = dfa_c.loc[:, lis_m7a].copy()
    # 将每一行加总并存储为一个新的 Series
    m7_sum = m7_df.sum(axis=1)
    # 将新的 Series 添加到 DataFrame 中
    m7_df['m7 Sum'] = m7_sum
    # 計算B欄逐行加總，並新增一個欄位'C'保存結果
    m7_df['cum_Sum'] = m7_df['m7 Sum'].cumsum()

    # 使用列 'index' 作为 x 轴，列 'cumsum' 作为 y 轴
    plt.figure(figsize=(10, 8))
    # 設置背景顏色
    plt.rcParams['axes.facecolor'] = '#f5deb3'
    # MMSE
    mmse7lis = ('').join(str(mmse7.Strategy.to_list()))

    sns.lineplot(x=m7_df.index, y='cum_Sum', data=m7_df, linestyle="--", color="red", label="MMSE-{}".format(mmse7lis))
    plt.xticks(fontsize=8)
    plt.yticks(fontsize=9)

    p7 = mp_mmse7.Strategy
    lis_p7 = p7.tolist()
    lis_p7a = [str(x) for x in lis_p7]
    p7_df = dfa_c.loc[:, lis_p7a].copy()
    # 将每一行加总并存储为一个新的 Series
    p7_sum = p7_df.sum(axis=1)
    # 将新的 Series 添加到 DataFrame 中
    p7_df['p7 Sum'] = p7_sum
    # 計算B欄逐行加總，並新增一個欄位'C'保存結果
    p7_df['cum_Sum'] = p7_df['p7 Sum'].cumsum()
    # MP_MMSE
    mpmmse7lis = ('').join(str(mp_mmse7.Strategy.to_list()))
    sns.lineplot(x=m7_df.index, y='cum_Sum', data=p7_df, linestyle="-", color="blue",
                 label="MP_MMSE-{}".format(mpmmse7lis))

    dfa2 = dfa.copy()
    dfa_sum = dfa2.sum(axis=1)
    dfa2['cum_Sum'] = dfa_sum.cumsum()
    sns.lineplot(x=m7_df.index, y='cum_Sum', data=dfa2, linestyle=":", color="gray", label="All Strategy")
    plt.xlabel("Date", fontdict={'fontsize': 12})
    plt.ylabel("Total Return", fontdict={'fontsize': 12})
    plt.title("OPT_Portfolio", fontdict={'fontsize': 18})
    plt.legend(loc='upper left', borderaxespad=0., fontsize='large', fancybox=True, edgecolor='navy', framealpha=0.2,
               handlelength=1.5, handletextpad=0.5, borderpad=0.5, labelspacing=0.5)
    # SQN
    sqn7lis = ('').join(str(sqn7.Strategy.to_list()))
    sns.lineplot(x=s7_df.index, y='cum_Sum', data=s7_df, linestyle="-", color="purple", label="SQN-{}".format(sqn7lis))

    global sqn7_30
    sqn7_30lis = ('').join(str(sqn7_30.Strategy.to_list()))
    sns.lineplot(x=s7_df_30.index, y='cum_Sum', data=s7_df_30, linestyle=":", color="purple",
                 label="SQN_30-{}".format(sqn7_30lis))

    ###################################################
    # 上一次的優勝
    '''
    ls_sqn=['1','3','2','6','11','14']
    dfa.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    last_sqn=dfa[ls_sqn]
    last_sqn['cumsum'] = last_sqn.sum(axis=1).cumsum()
    ls_sqn2=('').join(str(ls_sqn))
    '''
    sns.lineplot(x=dfak2.index, y='cumsum', data=dfak2, linestyle="--", color="orange",
                 label="Last_Better-{}".format(lis_sqn_old))
    plt.tight_layout()

    plt.savefig('temp.png')
    plt.savefig('ec.png')
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_ec = wb.sheets.add('Equip Curve')
    sht_ec.pictures.add(pic_path)
    plt.close()
    #########################################################################################
    # 將SQN本週策略序號及名稱加至sht_ec('Equip Curve')
    # 1、先將所有策略放至df
    last_row_s = ws.range('A96').current_region.rows.count + 95
    range_to_export = ws.range('A96:B{}'.format(last_row_s))
    # 將列表轉換為Pandas DataFrame
    dfx = pd.DataFrame(lis_sqn, columns=['num', 'filename'])
    sht_ec.range('P1').options(index=False).value = dfx
sqn_last10d_f()


def pca_f():
    from sklearn.preprocessing import MaxAbsScaler
    #returns = df_
    # 創建MaxAbsScaler對象
    scaler = MaxAbsScaler()
    # 對稀疏數據進行標準化
    data_scaled = scaler.fit_transform(dfa)
    returns = pd.DataFrame(data_scaled, columns=dfa.columns)

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
    plt.figure(figsize=(6, 4))

    # 绘制 sqn7 的策略点
    highlighted_strategies= sqn7.Strategy.index
    colors = ['red' if strategy in highlighted_strategies else 'blue' for strategy in range(len(pca.components_[0]))]
    sizes = [30 if strategy in highlighted_strategies else 5 for strategy in range(len(pca.components_[0]))]
    plt.scatter(pca.components_[0], pca.components_[1], c=colors, s=sizes, marker='.', facecolors='none')
    plt.scatter(pca.components_[0][highlighted_strategies], pca.components_[1][highlighted_strategies], c='red', s=30,
                marker='p')

# 添加坐标轴标签
    plt.title('SQN_Best_{}'.format(list(highlighted_strategies+1)))
    plt.xlabel('PC1')
    plt.ylabel('PC2')

    # 添加每个交易策略的标签
    strategies = list(returns.columns)
    strategies = [x[:-2] if x.endswith('.0') else x for x in strategies]
    for i, strategy in enumerate(strategies):
        if i in highlighted_strategies:
            color = 'blue'
            fs=11
        else:
            color = 'gray'
            fs=8
        plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]), color=color,fontsize=fs)

    # 显示图形
    pp = plt.plot
    plt.tight_layout()
    #plt.show()
    plt.savefig('temp.png')
    plt.close()
    sht_pic = wb.sheets.add('PCA')


    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    sht_pic.pictures.add(pic_path)
    # 删除临时文件
    os.remove('temp.png')





##################################################################
##################################################################
    from sklearn.preprocessing import MaxAbsScaler
    #returns = df_
    # 創建MaxAbsScaler對象
    scaler = MaxAbsScaler()
    # 對稀疏數據進行標準化
    data_scaled = scaler.fit_transform(dfa)
    returns = pd.DataFrame(data_scaled, columns=dfa.columns)
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
    plt.figure(figsize=(6, 4))


    highlighted_strategies_30= sqn7_30.Strategy.index
    colors = ['red' if strategy in highlighted_strategies_30 else 'blue' for strategy in range(len(pca.components_[0]))]
    sizes = [30 if strategy in highlighted_strategies_30 else 5 for strategy in range(len(pca.components_[0]))]
    plt.scatter(pca.components_[0], pca.components_[1], c=colors, s=sizes, marker='.', facecolors='none')
    plt.scatter(pca.components_[0][highlighted_strategies_30], pca.components_[1][highlighted_strategies_30], c='red', s=30,
                marker='p')

    # 添加坐标轴标签
    plt.title('SQN_30_Best_{}'.format(list(highlighted_strategies_30+1)))
    plt.xlabel('PC1')
    plt.ylabel('PC2')

    # 添加每个交易策略的标签
    strategies = list(returns.columns)
    strategies = [x[:-2] if x.endswith('.0') else x for x in strategies]
    for i, strategy in enumerate(strategies):
        if i in highlighted_strategies_30:
            color = 'blue'
            fs=11
        else:
            color = 'gray'
            fs=8
        plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]), color=color,fontsize=fs)



    #for i, strategy in enumerate(strategies):
    #    plt.annotate(strategy, (pca.components_[0][i], pca.components_[1][i]))
    # 显示图形
    plt.tight_layout()
    #plt.show()
    plt.savefig('temp.png')
    plt.close()
    pic_path = (os.path.join(os.getcwd(), "temp.png"))
    #sht_pic.pictures.add(pic_path)
    sht_pic.pictures.add(pic_path, name='MyPic1', left=sht_pic.range('J1').left, top=sht_pic.range('J1').top)
    # 删除临时文件
    os.remove('temp.png')
    # 显示Excel文件

    # 選擇要複製的圖片

    pic_path = (os.path.join(os.getcwd(), "ec.png"))
    sht_pic.pictures.add(pic_path, name='MyPic5', left=sht_pic.range('A19').left, top=sht_pic.range('A19').top,width=500,height=350)
    pic_path = (os.path.join(os.getcwd(), "nx.png"))
    sht_pic.pictures.add(pic_path, name='MyPic6', left=sht_pic.range('J19').left, top=sht_pic.range('J19').top)
    # 删除临时文件
    os.remove('ec.png')
    os.remove('nx.png')

    dfx = pd.DataFrame(lis_sqn, columns=['num', 'filename'])
    sht_pic.range('S1').options(index=False).value = dfx
pca_f()

wb.app.visible = True
plt.close()


wb.save('book.xlsx')
wb.close()

'''
chd=os.getcwd()
with open(chd+"\Sub_group.py", encoding="utf-8") as f:
    exec(f.read())
'''

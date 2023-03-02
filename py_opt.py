import pandas as pd
import numpy as np
from scipy.optimize import minimize, LinearConstraint, Bounds
import xlwings as xw
import string
import pandas as pd
import os
import matplotlib.pyplot as plt
# Read data
# 打開一個新的Excel文件
app=xw.App(visible=True,add_book=False)
app.display_alerts=False
app.screen_updating=True
wb = xw.Book('book.xlsx')

sht=wb.sheets('工作表1')
last_cell = sht.range('A2').current_region.rows.count
# 從當前工作表中取得第一個儲存格（A2）
def num_to_col(num):
    """將數字轉換為Excel欄位的英文命名法"""
    col = ''
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = string.ascii_uppercase[remainder] + col
    return col
cell = sht.range('A1')
last_column=sht.range('A2').current_region.columns.count
rng_all = sht.range('A1:{}'.format(num_to_col(last_column)+str(last_cell)))
dfa =pd.DataFrame(rng_all.value, columns=rng_all.value[0])
dfa = dfa.drop([0])


# 將 '日期' 列轉換為索引
dfa=df = dfa.drop(dfa.index[0])
dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
dfa = dfa.set_index('日期')
#dfa = dfa.drop(dfa.columns[0], axis=1)
dfa.columns
dfa = dfa.rename(columns=lambda x: str(x).strip())
dfa.rename(columns=lambda x: x.replace('.0', ''), inplace=True)

#總共16欄，每欄的第一行+10萬元
dfb=dfa.copy()
# 將每一欄的第一行資料加上10萬元
for col in dfb.columns:
    dfb.iloc[0][col] += 100000
# 計算每一欄的累加和
cumsum_df = dfb.cumsum()


# 獲取每列的權益曲線數據
equity_curves = dfb.cumsum()

# 使用matplotlib繪製曲線圖
fig, ax = plt.subplots()
for col in equity_curves.columns:
    ax.plot(equity_curves.index, equity_curves[col], label=col)
ax.legend()
ax.set_xlabel('日期')
ax.set_ylabel('權益')
ax.set_title('權益曲線')
plt.show()








import numpy as np
import pandas as pd
from pypfopt import expected_returns, risk_models
from pypfopt.efficient_frontier import EfficientFrontier

prices_df = cumsum_df
prices_df = prices_df.drop(prices_df.index[-1])
returns_df = prices_df.pct_change().dropna()
returns_df = returns_df.replace([np.inf, -np.inf], 0)
# 刪除最後一行
mu = expected_returns.mean_historical_return(returns_df)
S = risk_models.sample_cov(returns_df)



#先分成ABC三個dataset
al=['1','3','4','9','12','13','15']
sgA=dfa[al]
sgA['cumsum'] = sgA.sum(axis=1).cumsum()
bl=['10','6','2','7','5']
sgB=dfa[bl]
sgB['cumsum'] = sgB.sum(axis=1).cumsum()
cl=['14','8']
sgC=dfa[cl]
sgC['cumsum'] = sgC.sum(axis=1).cumsum()

#sqn7lis=('').join(str(sgA.Strategy.to_list()))


import seaborn as sns
import matplotlib.pyplot as plt
plt.figure(figsize=(10, 8))
sns.lineplot(x=sgA.index, y='cumsum', data=sgA,linestyle="-",color="red",label="Sub Group A-{}".format(al))
sns.lineplot(x=sgB.index, y='cumsum', data=sgB,linestyle="-",color="blue",label="Sub Group B-{}".format(bl))
sns.lineplot(x=sgC.index, y='cumsum', data=sgC,linestyle=":",color="green",label="Sub Group C-{}".format(cl))
plt.xlabel("Date",fontdict={'fontsize': 16})
plt.ylabel("Total Return",fontdict={'fontsize': 16})
plt.title("SubGroup Portfolio",fontdict={'fontsize': 22})
plt.legend(loc='upper left', borderaxespad=0., fontsize='large',fancybox=True, edgecolor='navy', framealpha=0.2, handlelength=3, handletextpad=1, borderpad=1, labelspacing=.5)
plt.tight_layout()
plt.savefig('temp.png')
pic_path=(os.path.join(os.getcwd(), "temp.png"))
sht_ec = wb.sheets.add('Sub_Group Equip Curve')
sht_ec .pictures.add(pic_path)

plt.close()
wb.save('book.xlsx')
wb.close

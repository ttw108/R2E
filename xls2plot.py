from tkinter import *
import tkinter as tk
from tkinter import simpledialog
import warnings
warnings.filterwarnings("ignore")
import xlwings as xw
import re
import string
import pandas as pd
import win32api

# 获取鼠标当前的位置
x, y = win32api.GetCursorPos()
root = tk.Tk()
root.withdraw()
# 创建一个 Toplevel 窗口
dialog = tk.Toplevel(root)



def proc_data():
    global plt
    # Create an input dialog
    strategies= simpledialog.askstring("Input", "要畫的_策略名稱",parent=dialog)
    dialog.withdraw()
    strategies=[int(i) for i in strategies.split(' ')]
    strategies = list(map(str, strategies))
    #將bookmark開啟
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
    #wb = xw.Book()  # this will open a new workbook
    wb = xw.Book('book.xlsx')
    #wb=app.books.open('Book1.xlsx')
    sht=wb.sheets("工作表1")
    # 從當前工作表中取得第一個儲存格（A2）
    cell = sht.range('A2')

    last_column = sht.range('A2').current_region.columns.count
    last_row = sht.range('A2').current_region.rows.count

    rng_all = sht.range('A1:{}'.format(num_to_col(last_column) + str(last_row)))
    dfa = pd.DataFrame(rng_all.value, columns=rng_all.value[0])

    # 將 '日期' 列轉換為索引

    dfa = dfa.drop(dfa.index[0])
    dfa['日期'] = pd.to_datetime(dfa['日期']).dt.date
    dfa = dfa.set_index('日期')

    dfa = dfa.rename(columns=lambda x: str(x).strip())
    dfa.rename(columns=lambda x: x.replace('.0', ''), inplace=True)
    dfb=dfa.copy()
    # 假設您的DataFrame名稱為dfa
    # 先將第一列加上本錢100000元
    dfa.iloc[0] = dfa.iloc[0] + 100000

    # 對每一欄做cumsum
    dfa = dfa.apply(lambda x: x.cumsum(), axis=0)
    dfa = dfa.drop(dfa.index[-1])

    #將所選的策略加總做cumsum
    dfb['sum'] = dfb[strategies].sum(axis=1)
    dfb = dfb.drop(dfb.index[-1])
    # 做逐列的 cumsum()

    dfb['sum'].iat[0] += 100000
    dfb['cumsum'] = dfb['sum'].cumsum()


    import matplotlib.pyplot as plt
    import seaborn as sns

    # 繪製子策略折線圖
    fig, ax = plt.subplots()
    for col in dfa.columns:
        if col in strategies:
            ax.plot(dfa.index, dfa[col], label=col,lw=0.4)
            ax.annotate(col, xy=(dfa.index[-1], dfa[col].iloc[-1]))
    ax.plot(dfb.index, dfb['cumsum'], label='cumsum',lw=1.3,ls='--')
    ax.annotate('cumsum', xy=(dfb.index[-1], dfb['cumsum'].iloc[-1]))
    # 設定標題、x軸和y軸標籤、圖例等
    ax.set_title('Daily P&L with Cumulative Sum-{}'.format(strategies))
    ax.set_xlabel('Date')
    ax.set_ylabel('Cumulative P&L')
    ax.legend(loc='upper left',fontsize=10)
    ax.xaxis.label.set_fontsize(12)
    ax.yaxis.label.set_fontsize(12)
    ax.xaxis.label.set_color('blue')
    ax.yaxis.label.set_color('blue')
    ax.tick_params(axis='both', labelsize=5)
    def on_key_press(event):
        if event.key == "escape":
            plt.close()
    plt.connect("key_press_event", on_key_press)
    plt.show(block=True)
    x, y = win32api.GetCursorPos()
    dialog.geometry(f"+{x}+{y}")
proc_data()

# 程序開始運行
while True:
    proc_data()







import matplotlib
import matplotlib.pyplot as plt
from matplotlib import ticker
from matplotlib.font_manager import FontProperties
import pandas as pd
import numpy as np
import datetime
from base.RJGraphing import osmkdir
import re
from matplotlib.ticker import Formatter
import matplotlib.dates as mdates
from base import cptj as cj


matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'times']  # 用黑体显示中文，Times New Roman 显示英文
matplotlib.rcParams['axes.unicode_minus'] = False  # 正常显示负号
# This is the global setting of specific fonts(applied to all)
simhei = FontProperties(fname=r"E:\ANo.3\fonts\simhei.TTF", size=12)
TimesNR = FontProperties(fname=r"E:\ANo.3\fonts\TimesNR.TTF", size=12)
Timesbd = FontProperties(fname=r"E:\ANo.3\fonts\timesbd.TTF", size=12)

df = pd.read_excel('C:/Users/ThinkPad/Desktop/to奕泽_标绿_20210722_682样本政策强度.xlsx',
                   sheet_name='682样本')
data = df.copy()

index = 'Quarter'
column = '监管强度指数'
address = "C:/Users/ThinkPad/Desktop/图片"

df1 = df.groupby(by=index, axis=0).sum()[column]
df2 = df.groupby(by=index, axis=0).count()[column]

plt.style.use('classic')
fig, ax = plt.subplots(1, 1, figsize=(16, 11))
ax.grid()
ax2 = ax.twinx()

inds1 = [cj.quarter2date(x) for x in df1.index]
inds2 = [cj.quarter2date(x) for x in df1.index]

ax.plot(inds1, df1, color='black', linewidth=1.6, label='监管强度指数')
ax2.bar(inds2, df2, color='lightgrey', width=60, alpha=0.35, label='监管政策数量')
ax.set_ylim(0, 21)
ax2.set_ylim(0, 90)

# The font properties of legend cannot be set independently
# ax.legend(loc=2, bbox_to_anchor=(.5, .1), labels=['政策强度'], prop=simhei)

# ax.legend(loc='best', labels=['监管强度指数', '监管政策数量'], prop=simhei, shadow=True)
ax.legend(loc='upper left', prop=simhei, shadow=True)
ax2.legend(loc='upper right', prop=simhei, shadow=True)

ax.xaxis.set_major_locator(ticker.MaxNLocator(20))
ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
ax.xaxis.set_major_locator(mdates.YearLocator())

ax.tick_params(axis='x', labelsize=13)

if type(df1.index[0]) is str:
    ax.tick_params(axis='x', rotation=40)
    pass
else:
    ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: "%.0f" % x))

ax.yaxis.set_minor_locator(ticker.AutoLocator())

plt.savefig(address, dpi=300)

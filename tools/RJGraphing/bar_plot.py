import matplotlib
import matplotlib.pyplot as plt
from matplotlib import ticker
from matplotlib.font_manager import FontProperties
import pandas as pd
import numpy as np
import datetime
from RJGraphing import osmkdir
import re
from matplotlib.ticker import Formatter
import matplotlib.dates as mdates
from PolicyAnalysis import cptj as cj
import os

# 设置工作路径
os.chdir('E:\\ANo.3\\FSML\\FinancialSupervision\\tools')

# 设置字体
matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'times']  # 用黑体显示中文，Times New Roman 显示英文
matplotlib.rcParams['axes.unicode_minus'] = False  # 正常显示负号
# This is the global setting of specific fonts(applied to all)
simhei = FontProperties(fname="./fonts/simhei.TTF", size=12)
TimesNR = FontProperties(fname="./fonts/TimesNR.TTF", size=12)
Timesbd = FontProperties(fname="./fonts/timesbd.TTF", size=12)


# 读取数据，建议直接把数据放到 tools 文件夹里
df = pd.read_excel('./Data.xlsx',     sheet_name='Sheet1')
data = df.copy()

# 设置要画的图形
index = 'Quarter'
column = '监管强度指数'
address = "./barplot"

df1 = df.groupby(by=index, axis=0).sum()[column]
df2 = df.groupby(by=index, axis=0).count()[column]

plt.style.use('classic')
fig, ax = plt.subplots(1, 1, figsize=(16, 11))
ax.grid()
# 设置一条次坐标轴
ax2 = ax.twinx()

inds1 = [cj.quarter2date(x) for x in df1.index]
inds2 = [cj.quarter2date(x) for x in df1.index]

# 画图
ax.plot(inds1, df1, color='black', linewidth=1.6, label='监管强度指数')
ax2.bar(inds2, df2, color='lightgrey', width=60, alpha=0.35, label='监管政策数量')

# 设置 y 主轴和 y 次轴的范围
ax.set_ylim(0, 21)
ax2.set_ylim(0, 90)

# The font properties of legend cannot be set independently
# ax.legend(loc=2, bbox_to_anchor=(.5, .1), labels=['政策强度'], prop=simhei)

# ax.legend(loc='best', labels=['监管强度指数', '监管政策数量'], prop=simhei, shadow=True)
# 设置图例
ax.legend(loc='upper left', prop=simhei, shadow=True)
ax2.legend(loc='upper right', prop=simhei, shadow=True)

# 设置 x 轴刻度的格式和排布，这里用年度刻度标记了季度数据，做了一点美化
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

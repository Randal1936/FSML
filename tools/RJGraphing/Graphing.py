import matplotlib
import matplotlib.pyplot as plt
from matplotlib import ticker
from matplotlib.font_manager import FontProperties
import pandas as pd
import numpy as np
import datetime
from RJGraphing import osmkdir

matplotlib.rcParams['font.sans-serif'] = ['SimHei', 'times']     # 用黑体显示中文，Times New Roman 显示英文
matplotlib.rcParams['axes.unicode_minus'] = False     # 正常显示负号
# This is the global setting of specific fonts(applied to all)
simhei = FontProperties(fname=r"E:\ANo.3\fonts\simhei.TTF", size=12)
TimesNR = FontProperties(fname=r"E:\ANo.3\fonts\TimesNR.TTF", size=12)
Timesbd = FontProperties(fname=r"E:\ANo.3\fonts\timesbd.TTF", size=12)


def policy_intensity(data, index, column, how, address="C:/Users/ThinkPad/Desktop/GraphFolder"):
    """
    :param data: pandas.DataFrame, the panel data
    :param index: str, the name of the column that contains the index
    :param column: str, the name of the column to be summarized
    :param how: str, the way used to summarize the data, including ['Sum', 'Mean', 'Count']
    :param address: str, the link where the output graph is stored
    :return: There is no return, but a graph will be stored into the assigned place
    """
    # Use this function to create a new folder (I don't wanna mess up my desktop)
    osmkdir.mkdir(address)
    address = address + '/' + column + '_by_' + index + '_' + how + '_'

    if how in ['sum', 'Sum']:
        data = data.groupby(by=index, axis=0).sum()[column]
        RJ_Plot(data, address, legend=column, xlabel='')

    elif how in ['mean', 'Mean']:
        data = data.groupby(by=index, axis=0).mean()[column]
        RJ_Plot(data, address, legend=column, xlabel=index)

    elif how in ['count', 'Count']:
        data = data.groupby(by=index, axis=0).count()[column]
        RJ_Plot(data, address, legend=column, xlabel=index)

    else:
        pass


def RJ_Plot(data, address, legend, xlabel=any):
    """
    This is my favorite style of plots
    :param data: pd.Series/pd.DataFrame, the input data
    :param address: str, the place assigned to the output figure, e.g. "C:/Users/ThinkPad/Desktop/Figure.png"
    :param legend: the name of the legend in the figure
    :param xlabel: the name of the xlabel in the figure
    :return: There is no return, but a graph will be stored into the assigned place
    """
    plt.style.use('classic')
    fig, ax = plt.subplots(1, 1, figsize=(16, 11))
    ax.grid()
    ax.plot(data, color='black')

    # The font properties of legend cannot be set independently
    # ax.legend(loc=2, bbox_to_anchor=(.5, .1), labels=['政策强度'], prop=simhei)
    ax.legend(loc='best', labels=[legend], prop=simhei, shadow=True)
    ax.xaxis.set_major_locator(ticker.MaxNLocator(20))
    ax.tick_params(axis='x', labelsize=13)
    if type(data.index[0]) is str:
        ax.tick_params(axis='x', rotation=40)
        pass
    else:
        ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda x, pos: "%.0f" % x))
    ax.yaxis.set_minor_locator(ticker.AutoLocator())
    ax.set_xlabel(xlabel, fontproperties=Timesbd, horizontalalignment='center', fontsize=14)

    # Complete the name of the file by adding timestamp
    time_now = datetime.datetime.today()
    time_now = time_now.strftime('%Y%m%d_%H%M')
    address = address + time_now + '.png'

    # Save the figure
    plt.savefig(address, dpi=300)


"""
-------------------------------------------------------------------------------
Execution: Graphing
-------------------------------------------------------------------------------
"""

def main():
    df = pd.read_excel('C:/Users/ThinkPad/Desktop/to奕泽_标绿_20210722_682样本政策强度.xlsx',
                       sheet_name='682样本')
    """
    ---------------监管强度---------------
    """

    # 监管强度(按年求和)
    policy_intensity(df.copy(),
                     index='Year',
                     column='监管强度指数',
                     how='Sum')

    # 监管强度(按年求平均)
    policy_intensity(df.copy(),
                     index='Year',
                     column='监管强度指数',
                     how='Mean')

    # 监管强度(按季求和)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='监管强度指数',
                     how='Sum')

    # 监管强度(按季求平均)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='监管强度指数',
                     how='Mean')

    """
    ---------------政策主体---------------
    """

    # 政策主体(按年求和)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策主体',
                     how='Sum')

    # 政策主体(按年求平均)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策主体',
                     how='Mean')

    # 政策主体(按季求和)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策主体',
                     how='Sum')

    # 政策主体(按季求平均)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策主体',
                     how='Mean')

    """
    ---------------政策基调---------------
    """

    # 政策基调(按年求和)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策基调',
                     how='Sum')

    # 政策基调(按年求平均)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策基调',
                     how='Mean')

    # 政策基调(按季求和)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策基调',
                     how='Sum')

    # 政策基调(按季求平均)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策基调',
                     how='Mean')

    """
    ---------------政策内容---------------
    """

    # 政策内容(按年求和)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策内容',
                     how='Sum')

    # 政策内容(按年求平均)
    policy_intensity(df.copy(),
                     index='Year',
                     column='政策内容',
                     how='Mean')

    # 政策内容(按季求和)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策内容',
                     how='Sum')

    # 政策内容(按季求平均)
    policy_intensity(df.copy(),
                     index='Quarter',
                     column='政策内容',
                     how='Mean')


if __name__ == " __main__":
    main()


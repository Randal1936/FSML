import pandas as pd
import numpy as np
import xlwings as xw
import jieba
import os
import re
from alive_progress import alive_bar
from base import cptj as cj
import time


os.chdir('E:/ANo.3/base')
indifile = 'E:/ANo.3/base/赋分指标清单.xlsx'
indisheet = "被监管业务"
userdict = 'E:/ANo.3/base/businesses.txt'

"""
————————————————————
以下是使用 re 检索+ DFC 映射的数据处理写法
————————————————————
"""


def businesses_re(Data, userdict_link=userdict):
    data = Data.copy()
    # 先获取关键词字典
    n = cj.txt_to_list(userdict_link)

    # 把被监管的业务分类，做成字典映射
    # 首先生成一个列表，标记一下关键词所在的位置
    loc = [(0, 4), (4, 10), (10, 15), (15, 19), (19, 22), (22, 26), (26, 29), (29, 31), (31, 40),
           (40, 41), (41, 42), (42, 43), (43, 44), (44, 45)]

    # 然后遍历列表，按照标记的位置生成关键词切片，把同类的关键词映射到相同的数值
    i = 0
    keymap = {}
    for rank in loc:
        lst = n[rank[0]: rank[1]]
        for item in lst:
            keymap[item] = i
        i += 1

    # 情况一，对全部正文进行检索
    result1 = cj.words_docs_freq(n, data)
    dfc1 = result1['DFC']
    dtm1_class = result1['DTM']
    dtm1_final = cj.dfc_sort_filter(dfc1, keymap, '被监管业务-正文分类统计.xlsx')

    # 情况二，对正文前十句话进行检索
    # 造一个正文栏只包括正文前十句话的样本矩阵
    tf = data
    for i in range(0, data.shape[0]):
        tf.iloc[i, 2] = cj.top_n_sent(10, data.iloc[i, 2])

    result2 = cj.words_docs_freq(n, tf)
    dfc2 = result2['DFC']
    dtm2_class = result2['DTM']
    dtm2_final = cj.dfc_sort_filter(dfc2, keymap, '被监管业务-前十句话分类统计.xlsx')

    # 情况三，仅对标题进行检索
    # 首先把样本弄成一样的格式
    # 建议用这种赋值+循环 iloc 赋值来新建样本
    # 否则会报乱七八糟的错：不能用 slice 来更改原 DataFrame 值啊 blablabla
    tf3 = data
    for i in range(0, data.shape[0]):
        tf3.iloc[i, 2] = data.iloc[i, 1]
    # 生成词频统计结果
    result3 = cj.words_docs_freq(n, tf3)
    dfc3 = result3['DFC']
    dtm3_class = result3['DTM']
    dtm3_final = cj.dfc_sort_filter(dfc3, keymap, '被监管业务-标题分类统计.xlsx')

    dtm_final = pd.concat([dtm1_final, dtm2_final, dtm3_final], axis=1)
    dtm_final.columns = ['被监管业务数（正文）', '被监管业务数（前十句）', '被监管业务数（标题）']

    dtm_aver_class = dtm_final.agg(np.mean, axis=1)
    dtm_aver_class = pd.DataFrame(dtm_aver_class, columns=['被监管业务数'])

    result = {'DTM_aver': dtm_aver_class,  # DTM 1、2、3 被监管业务数求均值
              'DTM_final': dtm_final,  # DTM 1、2、3 被监管业务种类数汇总
              'DTM1_class': dtm1_class,  # 按正文检索得到的 Doc-Term Matrix
              'DTM2_class': dtm2_class,  # 按前十句话检索的 Doc-Term Matrix
              'DTM3_class': dtm3_class}  # 按标题检索得到的 Doc-Term Matrix

    return result


"""
——————————————————————
以下是使用 jieba 检索+ DTM 映射的数据处理写法
——————————————————————
"""


def businesses_jieba(Data, indicator_file=indifile, sheet=indisheet, userdict_link=userdict):
    data = Data.copy()
    # 先获取关键词字典
    n = cj.txt_to_list(userdict_link)

    # 导入指标文件
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(indicator_file)
        sht = wb.sheets[sheet]
        df_indi = sht.used_range.value
        df_indi = pd.DataFrame(df_indi)
        df_indi.columns = df_indi.loc[0]
        df_indi.drop(0, axis=0, inplace=True)
        df_indi.dropna(axis=0, how='all', inplace=True)
    finally:
        app.quit()

    # 生成 Business 分类字典, {'Institution': [keyword1, keyword2, keyword3, ....], ....}
    keymap = {}
    for i in range(df_indi.shape[1]):
        keymap[df_indi.columns[i]] = list(df_indi.iloc[:, i].dropna(''))

    # 情况一，对全部正文进行检索
    dtm1 = cj.jieba_vectorizer(data, user_dict_link=userdict)
    dtm1_result = cj.dtm_sort_filter(dtm1, keymap, '被监管业务-正文分类统计.xlsx')
    dtm1_class = dtm1_result['DTM_class']
    dtm1_final = dtm1_result['DTM_final']

    # 情况二，对正文前十句话进行检索
    # 造一个正文栏只包括正文前十句话的样本矩阵
    tf = data.copy()
    for i in range(0, data.shape[0]):
        tf.iloc[i, 2] = cj.top_n_sent(10, data.iloc[i, 2])

    dtm2 = cj.jieba_vectorizer(tf, user_dict_link=userdict)
    dtm2_result = cj.dtm_sort_filter(dtm2, keymap, '被监管业务-前十句话分类统计.xlsx')
    dtm2_class = dtm2_result['DTM_class']
    dtm2_final = dtm2_result['DTM_final']

    # 情况三，仅对标题进行检索
    # 首先把样本弄成一样的格式
    # 建议用这种赋值+循环 iloc 赋值来新建样本
    # 否则会报乱七八糟的错：不能用 slice 来更改原 DataFrame 值啊 blablabla
    tf3 = data.copy()
    for i in range(0, data.shape[0]):
        tf3.iloc[i, 2] = data.iloc[i, 1]
    # 生成词频统计结果

    dtm3 = cj.jieba_vectorizer(tf3, user_dict_link=userdict)
    dtm3_result = cj.dtm_sort_filter(dtm3, keymap)
    dtm3_class = dtm3_result['DTM_class']
    dtm3_final = dtm3_result['DTM_final']

    dtm_final = pd.concat([dtm1_final, dtm2_final, dtm3_final], axis=1)
    dtm_final.columns = ['被监管业务数（正文）', '被监管业务数（前十句）', '被监管业务数（标题）']

    dtm_aver_class = dtm_final.agg(np.mean, axis=1)
    dtm_aver_class = pd.DataFrame(dtm_aver_class, columns=['被监管业务种类数'])

    result = {'DTM_aver': dtm_aver_class,  # DTM 1、2、3 被监管业务数求均值
              'DTM_final': dtm_final,  # DTM 1、2、3 被监管业务种类数汇总
              'DTM1_class': dtm1_class,  # 按正文检索得到的 Doc-Term Matrix
              'DTM2_class': dtm2_class,  # 按前十句话检索的 Doc-Term Matrix
              'DTM3_class': dtm3_class}  # 按标题检索得到的 Doc-Term Matrix

    return result


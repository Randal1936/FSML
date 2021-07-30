"""
Coding: utf-8
Author: Jesse
Date: 
Code Goal: 
Code Logic: 
"""
from tqdm import tqdm
import pandas as pd
import jieba
import numpy as np
import os
import xlwings as xw
import re
from base import cptj as cj


indifile = 'E:/ANo.3/base/赋分指标清单.xlsx'
indisheet = "被监管机构"
userdict = 'E:/ANo.3/base/institutions.txt'
userdict2 = 'E:/ANo.3/base/institutions.txt'
# 添加关键词词典
jieba.load_userdict(userdict)


"""
————————————————————————
以下是使用 re 检索+ DFC 映射的数据处理写法
————————————————————————
"""


def institutions_re(Data, userdict_link=userdict):
    data = Data.copy()
    # 先获取关键词字典
    n = cj.txt_to_list(userdict_link)

    # 把被监管的业务分类，做成字典映射
    # 首先生成一个列表，标记一下关键词所在的位置
    loc = [(0, 3), (3, 4), (4, 7), (7, 8), (8, 9), (9, 10), (10, 12), (12, 14), (14, 21),
           (21, 22), (22, 23), (23, 24), (24, 25)]

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
    dtm1_final = cj.dfc_sort_filter(dfc1, keymap, '被监管机构-正文分类统计.xlsx')

    # 情况二，对正文前十句话进行检索
    # 造一个正文栏只包括正文前十句话的样本矩阵
    tf = data
    for i in range(0, data.shape[0]):
        tf.iloc[i, 2] = cj.top_n_sent(10, data.iloc[i, 2])

    result2 = cj.words_docs_freq(n, tf)
    dfc2 = result2['DFC']
    dtm2_class = result2['DTM']
    dtm2_final = cj.dfc_sort_filter(dfc2, keymap, '被监管机构-前十句话分类统计.xlsx')

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
    dtm3_final = cj.dfc_sort_filter(dfc3, keymap, '被监管机构-标题分类统计.xlsx')

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
————————————————————————
以下是使用 jieba 分词后检索+ DTM 映射的数据处理写法
————————————————————————
"""


def institutions_jieba(df):
    '''
    :param df: 输入的样本框, {axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4:: freq}
    :return: 返回一个Series, {index=df['id'], values=number of institutions}
    institutions 会对样本进行切词 + 计数统计的处理
    '''

    print("开始检索正文……")
    df = df.copy()  # 防止对函数以外的样本框造成影响

    # 导入指标文件
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(indifile)
        sht = wb.sheets[indisheet]
        df_indi = sht.used_range.value
        df_indi = pd.DataFrame(df_indi)
        df_indi.columns = df_indi.loc[0]
        df_indi.drop(0, axis=0, inplace=True)
    finally:
        app.quit()

    # 生成 Institution 分类字典, {'Institution': [keyword1, keyword2, keyword3, ....], ....}
    keymap = {}
    for i in range(df_indi.shape[1]):
        keymap[df_indi.columns[i]] = list(df_indi.iloc[:, i].dropna(''))

    # 只取样本前50%个数的句子，句子个数不是整数的话就向下取整
    for i in range(df.shape[0]):
        df.iloc[i, 2] = cj.top_n_sent(10, df.iloc[i, 2], percentile=0.5)

    # 得到词向量矩阵
    ff = cj.jieba_vectorizer(df.copy())
    # 生成 Institution 种类数
    ff = cj.dtm_sort_filter(ff, keymap)

    dtm_class = ff['DTM_class']
    dtm_final = ff['DTM_final']
    dtm_final = pd.DataFrame(dtm_final, columns=['被监管机构种类数'])

    result = {'DTM_class': dtm_class,  # 按正文前十句话检索得到的 Doc Term Matrix
              'DTM_final': dtm_final}  # 按正文前十句话检索得到的被监管机构数

    # ff_total.to_excel('E:\Python\MLtext\Z模型建立\\0310新样本\Export_data_2_被监管机构.xlsx')
    return result


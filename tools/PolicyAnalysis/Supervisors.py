"""
Coding: utf-8
Author: Jesse
Code Goal:
完成一级指标：政策主体的指标构建，使用颁布主题清单。
Code Logic:
分成两个部分：颁布主体行政级别以及是否联合发布
颁布主体行政级别部分
"""

import pandas as pd
import jieba
import numpy as np
import os
import xlwings as xw
import re
from base import cptj as cj
from tqdm import tqdm
from numba import jit

# 添加关键词词典
jieba.load_userdict('E:/ANo.3/base/Supervisor.txt')
userdict = 'E:/ANo.3/base/Supervisor.txt'
indifile = 'E:/ANo.3/base/赋分指标清单.xlsx'
opsheet = '颁布主体行政级别'
# 定义一个Dataframe用于判断联合还是单独发布
middle = pd.DataFrame()


"""
————————————————————
以下是使用 re 检索+ DFC 映射的数据处理写法
————————————————————
"""


def score_map(indifile=indifile, opsheet=opsheet):
    # 导入指标文件
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(indifile)
        sht = wb.sheets[opsheet]
        df_indi = sht.used_range.value
        df_indi = pd.DataFrame(df_indi)
        df_indi.drop(0, axis=0, inplace=True)
        df_indi.reset_index(drop=True, inplace=True)

        sr_indi = df_indi[0]
        sr_score = df_indi[1]
        sr_map = dict([(k, v) for k, v in zip(sr_indi, sr_score)])
    finally:
        app.quit()
    return sr_map


def class_map(indifile=indifile, opsheet=opsheet):
    # 导入指标文件
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(indifile)
        sht = wb.sheets[opsheet]
        df_indi = sht.used_range.value
        df_indi = pd.DataFrame(df_indi)
        df_indi.drop(0, axis=0, inplace=True)
        df_indi.reset_index(drop=True, inplace=True)

        cls_indi = df_indi[0]
        cls_id = df_indi[2]
        cls_map = dict([(k, v) for k, v in zip(cls_indi, cls_id)])
    finally:
        app.quit()
    return cls_map


def supervisors_re(Data, userdict_link=userdict):
    """
    :param userdict_link: 关键词清单链接
    :param Data: 输入的样本框, {axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4:: freq}
    :return: 返回一个Series, {index=df['id'], values=level of supervisors}
    supervisors 会对输入的样本进行切词 + 词频统计处理，计算 发文主体+联合发布 的分数
    """
    lst = cj.txt_to_list(userdict_link)
    print('开始检索标题……')
    data = Data.copy()  # 防止对样本以外的样本框造成改动
    sr_map = score_map()
    cls_map = class_map()

    # 接下来对标题进行检索
    data['正文'] = data['标题']
    result_title = cj.words_docs_freq(lst, data)
    point_title = cj.dfc_point_giver(result_title['DFC'], sr_map)
    class_title = cj.dfc_sort_filter(result_title['DFC'], cls_map)

    # 接下来对来源进行检索
    print('开始检索来源……')
    data['正文'] = data['来源']
    result_source = cj.words_docs_freq(lst, data)
    point_source = cj.dfc_point_giver(result_source['DFC'], sr_map)
    class_source = cj.dfc_sort_filter(result_source['DFC'], cls_map)

    two_point = pd.concat([point_title, point_source], axis=1)
    two_class = pd.concat([class_title, class_source], axis=1)

    final_point = pd.DataFrame(two_point.agg(np.max, axis=1), columns=['颁布主体得分'])
    final_class = pd.DataFrame(two_class.agg(np.max, axis=1), columns=['是否联合发布'])
    final_class = final_class.applymap(lambda x: 1 if x > 1 else 0)

    final_point.fillna(0, inplace=True)
    final_class.fillna(0, inplace=True)

    export_data = pd.concat([final_class, final_point], axis=1)
    # ff_export_data.to_excel('Export_data_1_颁布主体+是否联合发布.xlsx')
    return export_data


"""
————————————————————————
以下是使用 jieba 分词后检索+ DTM 映射的数据处理写法
————————————————————————
"""


def point_map(indifile=indifile, opsheet=opsheet):
    # 导入指标文件
    app = xw.App(visible=False, add_book=False)
    app.screen_updating = False
    app.display_alerts = False
    try:
        wb = app.books.open(indifile)
        sht = wb.sheets[opsheet]
        df_indi = sht.used_range.value
        df_indi = pd.DataFrame(df_indi)
        df_indi.drop(0, axis=0, inplace=True)
        df_indi.reset_index(drop=True, inplace=True)

        sr_indi = df_indi[1]
        sr_score = df_indi[2]
        sr_map = [(v, k) for k, v in zip(sr_indi, sr_score)]
        sr_map = list(pd.Series(sr_map).unique())
        sr_map = dict(sr_map)
    finally:
        app.quit()
    return sr_map


def sort_map():
    cls_map = class_map()
    cls_map_reverse = {}
    sorts = list(pd.Series(list(cls_map.values())).unique())
    cls_map_lst = [(v, k) for k, v in cls_map.items()]
    for sort in sorts:
        label = []
        for tup in cls_map_lst:
            if tup[0] == sort:
                label.append(tup[1])
        cls_map_reverse[sort] = label
    return cls_map_reverse


def supervisors_jieba(Data, userdict_link=userdict):
    """
    :param userdict_link: 关键词清单链接
    :param Data: 输入的样本框, {axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4:: freq}
    :return: 返回一个Series, {index=df['id'], values=level of supervisors}
    supervisors 会对输入的样本进行切词 + 词频统计处理，计算 发文主体+联合发布 的分数
    """

    print('开始检索标题……')
    data = Data.copy()  # 防止对样本以外的样本框造成改动
    sr_map = point_map()
    cls_map = sort_map()

    # 接下来对标题进行检索
    data['正文'] = data['标题']
    result_title = cj.jieba_vectorizer(data, user_dict_link=userdict_link, orient=True)
    class_title = cj.dtm_sort_filter(result_title, cls_map)['DTM_final']
    point_title = cj.dtm_point_giver(result_title, cls_map, sr_map)

    # 接下来对来源进行检索
    print('开始检索来源……')
    data['正文'] = data['来源']
    result_source = cj.jieba_vectorizer(data, user_dict_link=userdict_link, orient=True)
    class_source = cj.dtm_sort_filter(result_source, cls_map)['DTM_final']
    point_source = cj.dtm_point_giver(result_source, cls_map, sr_map)

    two_point = pd.concat([point_title, point_source], axis=1)
    two_class = pd.concat([class_title, class_source], axis=1)

    final_point = pd.DataFrame(two_point.agg(np.max, axis=1), columns=['颁布主体得分'])
    final_class = pd.DataFrame(two_class.agg(np.max, axis=1), columns=['是否联合发布'])
    final_class = final_class.applymap(lambda x: 1 if x > 1 else 0)

    final_point.fillna(0, inplace=True)
    final_class.fillna(0, inplace=True)

    export_data = pd.concat([final_class, final_point], axis=1)
    # ff_export_data.to_excel('Export_data_1_颁布主体+是否联合发布.xlsx')
    return export_data


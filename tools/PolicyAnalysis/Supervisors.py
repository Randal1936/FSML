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
import numpy as np
import xlwings as xw
from PolicyAnalysis import cptj as cj


"""
————————————————————
以下是使用 re 检索+ DFC 映射的数据处理写法
————————————————————
"""


class supervisors_re:

    def __init__(self, Data, userdict, indifile, opsheet):
        # 添加关键词词典
        self.userdict = userdict
        self.indifile = indifile
        self.opsheet = opsheet
        self.Data = Data
        self.DTM = None
        self.DFC = None
        self.cls_map = None
        self.sr_map = None
        self.pt_map = None
        self.export = None

        # 定义一个Dataframe用于判断联合还是单独发布
        self.middle = pd.DataFrame()
        self.score_map()
        self.class_map()
        self.supervisors()

    def score_map(self):
        # 导入指标文件
        app = xw.App(visible=False, add_book=False)
        app.screen_updating = False
        app.display_alerts = False
        try:
            wb = app.books.open(self.indifile)
            sht = wb.sheets[self.opsheet]
            df_indi = sht.used_range.value
            df_indi = pd.DataFrame(df_indi)
            df_indi.drop(0, axis=0, inplace=True)
            df_indi.reset_index(drop=True, inplace=True)

            sr_indi = df_indi[0]
            sr_score = df_indi[1]
            sr_map = dict([(k, v) for k, v in zip(sr_indi, sr_score)])
        finally:
            app.quit()
        self.sr_map = sr_map

    def class_map(self):
        # 导入指标文件
        app = xw.App(visible=False, add_book=False)
        app.screen_updating = False
        app.display_alerts = False
        try:
            wb = app.books.open(self.indifile)
            sht = wb.sheets[self.opsheet]
            df_indi = sht.used_range.value
            df_indi = pd.DataFrame(df_indi)
            df_indi.drop(0, axis=0, inplace=True)
            df_indi.reset_index(drop=True, inplace=True)

            cls_indi = df_indi[0]
            cls_id = df_indi[2]
            cls_map = dict([(k, v) for k, v in zip(cls_indi, cls_id)])
        finally:
            app.quit()
        self.cls_map = cls_map

    def supervisors(self):
        """
        :param userdict_link: 关键词清单链接
        :param Data: 输入的样本框, {axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4:: freq}
        :return: 返回一个Series, {index=df['id'], values=level of supervisors}
        supervisors 会对输入的样本进行切词 + 词频统计处理，计算 发文主体+联合发布 的分数
        """
        lst = cj.txt_to_list(self.userdict)
        print('开始检索标题……')
        data = self.Data.copy()  # 防止对样本以外的样本框造成改动

        # 接下来对标题进行检索
        data['正文'] = data['标题']
        result_title = cj.words_docs_freq(lst, data)
        point_title = cj.dfc_point_giver(result_title['DFC'], self.sr_map)
        class_title = cj.dfc_sort_filter(result_title['DFC'], self.cls_map)

        # 接下来对来源进行检索
        print('开始检索来源……')
        data['正文'] = data['来源']
        result_source = cj.words_docs_freq(lst, data)
        self.DFC = result_source['DFC']
        self.DTM = result_source['DTM']

        point_source = cj.dfc_point_giver(self.DFC, self.sr_map)
        class_source = cj.dfc_sort_filter(self.DTM, self.cls_map)

        two_point = pd.concat([point_title, point_source], axis=1)
        two_class = pd.concat([class_title, class_source], axis=1)

        final_point = pd.DataFrame(two_point.agg(np.max, axis=1), columns=['颁布主体得分'])
        final_class = pd.DataFrame(two_class.agg(np.max, axis=1), columns=['是否联合发布'])
        final_class = final_class.applymap(lambda x: 1 if x > 1 else 0)

        final_point.fillna(0, inplace=True)
        final_class.fillna(0, inplace=True)

        export_data = pd.concat([final_class, final_point], axis=1)
        # ff_export_data.to_excel('Export_data_1_颁布主体+是否联合发布.xlsx')
        self.export = export_data


"""
————————————————————————
以下是使用 jieba 分词后检索+ DTM 映射的数据处理写法
————————————————————————
"""

class supervisor_jieba:

    def __init__(self, Data, userdict, indifile, opsheet, stopwords):
        self.userdict = userdict
        self.indifile = indifile
        self.opsheet = opsheet
        self.stopwords = stopwords
        self.Data = Data
        self.DTM = None
        self.DFC = None
        self.cls_map = None
        self.sr_map = None
        self.pt_map = None
        self.export = None

        # 定义一个Dataframe用于判断联合还是单独发布
        self.middle = pd.DataFrame()
        self.point_map()
        self.class_map()
        self.sort_map()
        self.supervisors()

    def class_map(self):
        # 导入指标文件
        app = xw.App(visible=False, add_book=False)
        app.screen_updating = False
        app.display_alerts = False
        try:
            wb = app.books.open(self.indifile)
            sht = wb.sheets[self.opsheet]
            df_indi = sht.used_range.value
            df_indi = pd.DataFrame(df_indi)
            df_indi.drop(0, axis=0, inplace=True)
            df_indi.reset_index(drop=True, inplace=True)

            cls_indi = df_indi[0]
            cls_id = df_indi[2]
            cls_map = dict([(k, v) for k, v in zip(cls_indi, cls_id)])
        finally:
            app.quit()
        self.cls_map = cls_map

    def point_map(self):
        # 导入指标文件
        app = xw.App(visible=False, add_book=False)
        app.screen_updating = False
        app.display_alerts = False
        try:
            wb = app.books.open(self.indifile)
            sht = wb.sheets[self.opsheet]
            df_indi = sht.used_range.value
            df_indi = pd.DataFrame(df_indi)
            df_indi.drop(0, axis=0, inplace=True)
            df_indi.reset_index(drop=True, inplace=True)

            pt_indi = df_indi[1]
            pt_score = df_indi[2]
            pt_map = [(v, k) for k, v in zip(pt_indi, pt_score)]
            pt_map = list(pd.Series(pt_map).unique())
            pt_map = dict(pt_map)
        finally:
            app.quit()
        self.pt_map = pt_map

    def sort_map(self):
        cls_map = self.cls_map
        cls_map_reverse = {}
        sorts = list(pd.Series(list(cls_map.values())).unique())
        cls_map_lst = [(v, k) for k, v in cls_map.items()]
        for sort in sorts:
            label = []
            for tup in cls_map_lst:
                if tup[0] == sort:
                    label.append(tup[1])
            cls_map_reverse[sort] = label
        self.sr_map = cls_map_reverse

    def supervisors(self):
        """
        :param userdict: 关键词清单链接
        :param Data: 输入的样本框, {axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4:: freq}
        :return: 返回一个Series, {index=df['id'], values=level of supervisors}
        supervisors 会对输入的样本进行切词 + 词频统计处理，计算 发文主体+联合发布 的分数
        """

        print('开始检索标题……')
        data = self.Data.copy()  # 防止对样本以外的样本框造成改动

        # 接下来对标题进行检索-----------------------------------------------
        data['正文'] = data['标题']
        result = cj.jieba_vectorizer(data, self.userdict, self.stopwords, orient=True)
        self.title_DTM0 = result.DTM0   # 检索标题得到的原始矩阵
        self.title_features = result.features  # 检索标题得到的词语清单

        result_title = result.DTM

        class_title = cj.dtm_sort_filter(result_title, self.sr_map)['DTM_final']
        point_title = cj.dtm_point_giver(result_title, self.sr_map, self.pt_map)

        # 接下来对来源进行检索-----------------------------------------------
        print('开始检索来源……')
        data['正文'] = data['来源']
        result = cj.jieba_vectorizer(data, self.userdict, self.stopwords, orient=True)
        self.source_DTM0 = result.DTM0  # 检索来源得到的原始矩阵
        self.source_features = result.features  # 检索来源得到的词语清单

        result_source = result.DTM

        class_source = cj.dtm_sort_filter(result_source, self.sr_map)['DTM_final']
        point_source = cj.dtm_point_giver(result_source, self.sr_map, self.pt_map)

        two_point = pd.concat([point_title, point_source], axis=1)
        two_class = pd.concat([class_title, class_source], axis=1)

        final_point = pd.DataFrame(two_point.agg(np.max, axis=1), columns=['颁布主体得分'])
        final_class = pd.DataFrame(two_class.agg(np.max, axis=1), columns=['是否联合发布'])

        final_class = final_class.applymap(lambda x: 2 if x > 1 else 1)

        final_point.fillna(1, inplace=True)
        final_class.fillna(1, inplace=True)

        export_data = pd.concat([final_class, final_point], axis=1)

        self.class_title = class_title  # 检索标题得到的监管主体种类数
        self.class_source = class_source  # 检索来源得到的监管主体种类数
        self.point_title = point_title  # 检索标题得到的监管主体得分
        self.point_source = point_source  # 检索来源得到的监管主体得分

        # ff_export_data.to_excel('Export_data_1_颁布主体+是否联合发布.xlsx')
        self.export = export_data

"""
Coding: UTF-8
Author: Randal
Time: 2021/3/18
E-mail: RandalJin@163.com

Description: This script is designed to prepare all the data required for
 the intensity measure of financial regulation, including level of publishers,
 number of institutions, positive and negative tones, number of supervised
 businesses, number of numerals in the text and number of titles and title levels.

All rights reserved.

"""
import jieba
import pandas as pd
import numpy as np
import os
import xlwings as xw
import datetime
from PolicyAnalysis import cptj as cj
from PolicyAnalysis import Supervisors
from PolicyAnalysis import Institutions
from PolicyAnalysis import NegativeTone
from PolicyAnalysis import Businesses
from PolicyAnalysis import Titles
from PolicyAnalysis import Numerals

"""
——————————————
First Ⅰ - Get Primary Data
——————————————
"""

# 设置项目路径兼数据读取路径
# 建议这里就不要改了，直接把数据文件拿到 tools 文件夹里
os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')

# 导入原始数据
app1 = xw.App(visible=False, add_book=False)
try:
    wb = app1.books.open("调试数据.xlsx")
    sht = wb.sheets['Sheet1']
    df = sht.used_range.value
    df = pd.DataFrame(df)
    df.columns = list(df.loc[0])
    df.drop(0, axis=0, inplace=True)
    df.reset_index(inplace=True, drop=True)
    # for i in range(df.shape[0]):
    #     try:
    #         df.iloc[i, 4] = cj.re_formatter(df.iloc[i, 4])
    #     except ValueError:
    #         pass
    # sht['A1'].value = df
    # wb.save()
finally:
    app1.quit()


"""
——————————————
Second Ⅱ - Data Processing
——————————————
"""

# Resize the DataFrame and drop null values
# column 1: Doc id
# column 2: Title
# column 3: Text
# column 4: Source
# column 5: Freq
# column 6: time
Data = pd.DataFrame(df[['id', '标题', '正文', '来源', '日期']].values, columns=df[['id', '标题', '正文', '来源', '日期']].columns)
Data.dropna(axis=0, how='all', inplace=True)
# id 化为整数
Data['id'] = Data['id'].apply(lambda x: int(x))
Data.insert(4, 'freq', None)

# 1. Level of publishers
# (1) Score on the authoritative level
# (2) Score on the syndication level (joint release or not)
supervisors = Supervisors.supervisor_jieba(Data,
                                           userdict=os.path.abspath('./words_list/Supervisor.txt'),
                                           indifile='./words_list/赋分指标清单.xlsx',
                                           opsheet='颁布主体行政级别',
                                           stopwords='./words_list/stop_words.txt')

supervisors_score = supervisors.export

# 2. Number of institutions (top 10 sentences)
institutions = Institutions.institutions_jieba(Data,
                                                     userdict=os.path.abspath('./words_list/institutions.txt'),
                                                     indifile='./words_list/赋分指标清单.xlsx',
                                                     indisheet="被监管机构",
                                                     stopwords='./words_list/stop_words.txt')

institutions_score = institutions.DTM_final

# 3. Positive and negative tones
# (1) Relative negative tone
# (2) Absolute negative tone
negative_tone = NegativeTone.negative_tone_jieba(Data,
                                                 userdict=os.path.abspath('./words_list/情感词词典_加入政策词汇.txt'),
                                                 posidict='./words_list/正向情感词词典_加入政策词汇.txt',
                                                 negadict='./words_list/负向情感词词典_加入政策词汇.txt',
                                                 stopwords='./words_list/stop_words.txt')

negative_tone_score = negative_tone.tone

# 4. Number of supervised businesses (Average of titles/top 10 sentences/text)
# Sorts of supervised businesses (Average)
business = Businesses.business_jieba(Data,
                                                userdict=os.path.abspath('./words_list/businesses.txt'),
                                                indifile='./words_list/赋分指标清单.xlsx',
                                                indisheet="被监管业务",
                                                stopwords='./words_list/stop_words.txt')

business_score = business.DTM_aver

# 5. Number of titles and title levels
titles = Titles.titles(Data)
titles_score = titles.DTM_final

# 6. Number of numerals in the text
numerals = Numerals.numerals(Data)
numeral_score = numerals.DTM
numeral_score = pd.DataFrame(numeral_score.agg(np.sum, axis=1), columns=['数字个数'])

# 7. Time of issuance
# (1) By year-month-day

year_month_day = pd.DataFrame(Data['日期'].values, index=Data['id'], columns=['日期'])
for i in range(year_month_day.shape[0]):
    try:
        year_month_day.iloc[i, 0] = cj.re_formatter(str(year_month_day.iloc[i, 0])[: 10])
    except ValueError:
        print('Date missing or mis-specified，Doc ' + str(i))

# (2) By years

year = year_month_day.copy()
for i in range(year_month_day.shape[0]):
    try:
        year.iloc[i, 0] = int(year.iloc[i, 0][: 4])
    except TypeError:
        pass
year.columns = ['Year']

# (3) By quarters

quarter = year_month_day.copy()
for i in range(year_month_day.shape[0]):
    quarter.iloc[i, 0] = cj.quarter_finder(quarter.iloc[i, 0])
quarter.columns = ['Quarter']

Data.drop(['freq'], axis=1, inplace=True)  # freq 这一列是 cj 词频统计的前提，所以要到最后才能删去

Data = pd.DataFrame(Data.values, index=Data['id'], columns=Data.columns)

result = pd.concat([Data.iloc[:, 1: ],  # index: id, 0: 标题, 1: 正文, 2: 来源, 3: 年月
                    year,  # 4: 年份
                    quarter,  # 5: 年份-季度，例 2020Q4
                    supervisors_score,  # 6: 是否联合发布, 7: 颁布主体得分
                    institutions_score,  # 8: 被监管机构种类数
                    negative_tone_score,  # 9: 相对情感语调, 10: 绝对情感语调
                    business_score,  # 11: 被监管业务数
                    titles_score,  # 12: 标题层级数, 13: 标题个数
                    numeral_score], axis=1)  # 14: 数字个数（硬性约束个数）

# Before standardization, indexers must be <class str> to comply with the pandas rules
str_id_indexer = [str(indexer) for indexer in Data['id']]
result = pd.DataFrame(result.values, index=str_id_indexer, columns=result.columns)

"""
——————————————
Third Ⅲ - Data Standardization
——————————————
"""

# Data_primary = result.drop(['标题', '正文', '来源', '日期'], axis=1)
#
# # 1、Max_Min_Standardization
# Data_mm_stdizd = pd.DataFrame(Data_primary.values, index=Data_primary.index, columns=Data_primary.columns)
# col = 0
# for column in Data_mm_stdizd.columns:
#     if col > 1:
#         col_max = np.max(Data_mm_stdizd[column])
#         col_min = np.min(Data_mm_stdizd[column])
#         col = str(col)
#         for row in Data_mm_stdizd.index:
#             try:
#                 Data_mm_stdizd.loc[row, column] = (Data_primary.loc[row, column] - col_min)/(col_max - col_min)
#             except TypeError:
#                 pass
#     col = int(col)
#     col += 1
#
#
# # 2、Z-core_Standardization
# Data_zc_stdizd = pd.DataFrame(Data_primary.values, index=Data_primary.index, columns=Data_primary.columns)
# col = 0
# for column in Data_zc_stdizd.columns:
#     if col > 1:
#         col_mean = np.mean(Data_zc_stdizd[column])
#         col_std = np.std(Data_zc_stdizd[column])
#         col = str(col)  # only <class str> is identifiable to DataFrame.loc
#         for row in Data_zc_stdizd.index:
#             try:
#                 Data_zc_stdizd.loc[row, column] = (Data_primary.loc[row, column] - col_mean)/col_std
#             except TypeError:
#                 pass
#     col = int(col)
#     col += 1


"""
————————————
Fourth Ⅳ - Data Export
————————————
"""


time_now = datetime.datetime.today()

# Beware that 'm' and 'd' must be lowercase
time_now = time_now.strftime('%Y%m%d_%H%M')

# 获取一个 overall 分词结果，用于样本去重和样本分筛
overall = cj.jieba_vectorizer(df, userdict='./words_list/BSI.txt',
                                  stopwords='./words_list/stop_words.txt')
overall_DTM_all = overall.DTM
overall_DTM_key = overall.strip_non_keywords(overall_DTM_all)

# 如果 overall DTM 分出的词语超过 16000 个，就做一下行列转置
# 因为 16000 是 excel 表格列数极限
if overall_DTM_all.shape[1] > 16000:
    overall_DTM_all = pd.DataFrame(overall_DTM_all.values.T,
                                   columns=overall_DTM_all.index,
                                   index=overall_DTM_all.columns)

os.chdir('C:/Users/ThinkPad/Desktop/')

name = str(len(Data)) + ' Samples_Export_Data_'+time_now+'.xlsx'
with pd.ExcelWriter(name) as writer:
    # Input the primary data
    result.to_excel(writer, sheet_name="Data")

    # Input supervised businesses DTM (Only Top 10 sentences)
    DTM = business.DTM2_class
    DTM = pd.DataFrame(DTM, index=Data['id']).dropna(axis=0, how='all')
    DTM = pd.concat([year, quarter, DTM], axis=1)
    DTM.to_excel(writer, sheet_name="Businesses")

    # Input supervised institutions DTM (Only Top 10 sentences)
    DTM = institutions.DTM_class
    DTM = pd.DataFrame(DTM, index=Data['id']).dropna(axis=0, how='all')
    DTM = pd.concat([year, quarter, DTM], axis=1)
    DTM.to_excel(writer, sheet_name="Institutions")

    # # Input Max_Min_Standardized Data
    # Data_mm_stdizd.to_excel(writer, sheet_name='Data_mm_stdizd')

    # # Input Z-core_Standardized Data
    # Data_zc_stdizd.to_excel(writer, sheet_name='Data_zc_stdizd')

    # Input titles DFC (Full text)
    DFC = titles.DFC
    DFC.to_excel(writer, sheet_name="titles")

    # Input numerals DFC (Full text)
    DFC = numerals.DFC
    DFC.to_excel(writer, sheet_name="numerals")

    # Input overall DTM
    overall_DTM_all.to_excel(writer, sheet_name="overall_DTM_all")
    overall_DTM_key.to_excel(writer, sheet_name="overall_DTM_key")


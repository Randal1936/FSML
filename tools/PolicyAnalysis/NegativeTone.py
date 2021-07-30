"""
Coding: utf-8
Author: Jesse
Date: 20210301
Code Goal: 使用已构建的正负向情感语调词典和既定情感语调计算方法，对每个文献的情感倾向打分。
Code Logic: 使用jieba切词，统计正负向情感语调词典的词频，然后代入公式计算。
"""
import numpy as np
import pandas as pd
import jieba
import os
import xlwings as xw
import re
from base import cptj as cj
from tqdm import tqdm

posidict = 'E:/ANo.3/base/正向情感词词典_加入政策词汇.txt'
negadict = 'E:/ANo.3/base/负向情感词词典_加入政策词汇.txt'
userdict = 'E:/ANo.3/base/情感词词典_加入政策词汇.txt'
indifile = 'E:/ANo.3/base/赋分指标清单.xlsx'


"""
————————————————————————
以下是使用 jieba 分词后检索+ DTM 映射的数据处理写法
————————————————————————
"""


def negative_tone_jieba(Data, userdict_link=userdict, posidict_link=posidict, negadict_link=negadict):
    data = Data.copy()
    data = cj.jieba_vectorizer(data, user_dict_link=userdict_link)
    positive_tone = cj.dataframe_filter(data, cj.txt_to_list(posidict_link), axis=1)
    negative_tone = cj.dataframe_filter(data, cj.txt_to_list(negadict_link), axis=1)

    data_sum = data.agg(np.sum, axis=1)
    positive_sum = positive_tone.agg(np.sum, axis=1)
    negative_sum = negative_tone.agg(np.sum, axis=1)

    absolute_negative_tone = negative_sum/data_sum
    relative_negative_tone = (negative_sum - positive_sum)/(positive_sum + negative_sum)

    words = pd.concat([positive_sum, negative_sum, data_sum])
    words.columns = ['正向情感词词频', '负向情感词词频', '总词数']
    tone = pd.concat([relative_negative_tone, absolute_negative_tone], axis=1)
    tone.columns = ['相对情感语调', '绝对情感语调']

    result = {'Tone': tone,  # 相对情感语调和绝对情感语调计算结果
              'Words': words}  # 正向情感词词频，负向情感词词频，总词数统计结果

    return result


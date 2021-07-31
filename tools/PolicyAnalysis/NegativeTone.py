"""
Coding: utf-8
Author: Jesse
Date: 20210301
Code Goal: 使用已构建的正负向情感语调词典和既定情感语调计算方法，对每个文献的情感倾向打分。
Code Logic: 使用jieba切词，统计正负向情感语调词典的词频，然后代入公式计算。
"""
import numpy as np
import pandas as pd
from PolicyAnalysis import cptj as cj

"""
————————————————————————
以下是使用 jieba 分词后检索+ DTM 映射的数据处理写法
————————————————————————
"""


class negative_tone_jieba:

    def __init__(self, Data, userdict, posidict, negadict, stopwords):
        self.Data = Data
        self.userdict = userdict
        self.posidict = posidict
        self.negadict = negadict
        self.stopwords = stopwords

        data = Data.copy()
        data = cj.jieba_vectorizer(data, self.userdict, self.stopwords).DTM
        positive_tone = cj.dataframe_filter(data, cj.txt_to_list(self.posidict), axis=1)
        negative_tone = cj.dataframe_filter(data, cj.txt_to_list(self.negadict), axis=1)

        data_sum = data.agg(np.sum, axis=1)
        positive_sum = positive_tone.agg(np.sum, axis=1)
        negative_sum = negative_tone.agg(np.sum, axis=1)

        absolute_negative_tone = negative_sum/data_sum
        relative_negative_tone = (negative_sum - positive_sum)/(positive_sum + negative_sum)

        words = pd.concat([positive_sum, negative_sum, data_sum])
        words.columns = ['正向情感词词频', '负向情感词词频', '总词数']
        tone = pd.concat([relative_negative_tone, absolute_negative_tone], axis=1)
        tone.columns = ['相对情感语调', '绝对情感语调']

        self.tone = tone  # 相对情感语调和绝对情感语调计算结果
        self.words = words  # 正向情感词词频，负向情感词词频，总词数统计结果

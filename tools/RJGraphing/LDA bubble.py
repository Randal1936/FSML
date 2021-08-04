import numpy as np
import pandas as pd
import os
from PolicyAnalysis import cptj as cj
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from sklearn.decomposition import LatentDirichletAllocation
import pyLDAvis.sklearn


os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')


df = pd.read_excel('./调试数据.xlsx')
tf = cj.jieba_vectorizer(df,
                         './words_list/BSI.txt',
                         './words_list/stop_words.txt')

# 获取原装的 CountVectotizer, 在 pyLDAvis 当中要用到
# 毕竟 jieba_vectorizer 就是一个假货，一层包装纸，做不得真
vectorizer = tf.vectorizer

# LatentDirichletAllocation 和 pyLDAvis 的 prepare 函数都只识别 matrix 格式
dtm = tf.DTM0

# 可以设置一个 n 的列表，把 n 写成循环，然后批量绘制气泡图
# 这里以 n = 20 为例
n = 10
lda_tf = LatentDirichletAllocation(n, random_state=0)
lda_tf.fit(dtm)

d = pyLDAvis.sklearn.prepare(lda_tf, dtm, vectorizer)
pyLDAvis.save_html(d, '银保监_lda_'+str(n)+'.html')




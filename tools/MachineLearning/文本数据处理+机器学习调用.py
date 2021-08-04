import pandas as pd
import numpy as np
import jieba
import os
import xlwings as xw
import re
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.metrics import accuracy_score
from sklearn.decomposition import LatentDirichletAllocation
from sklearn.ensemble import RandomForestClassifier
from sklearn.naive_bayes import BernoulliNB
from sklearn.linear_model import LogisticRegression
from sklearn.neural_network import MLPClassifier
import pyLDAvis.sklearn
from matplotlib import pyplot as plt
import time
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
from PolicyAnalysis import cptj as cj

os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')

df = pd.read_excel()


# 朴素贝叶斯法
def NaiveBayesianPred(x):
    X = pd.read_csv('txt_vector.csv')
    y = pd.read_excel('Policy.xlsx')
    y = y['评分']

    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        nb = BernoulliNB()
        nb.fit(X_train, y_train)

        y_pred = nb.predict(X_test)
        y_pred_prob = nb.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print(exp)
    print(stab)
    return 0


# 逻辑回归法
def LogisticRegressionPred(x):
    X = pd.read_csv('txt_vector.csv')
    y = pd.read_excel('Policy.xlsx')
    y = y['评分']

    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        lr = LogisticRegression()
        lr.fit(X_train, y_train)

        y_pred = lr.predict(X_test)
        y_pred_prob = lr.predict_proba(X_test)
        score = accuracy_score(y_test, y_pred)
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print(exp)
    print(stab)
    return 0


# 神经网络法
def NeuralNetworkPred(x):
    X = pd.read_csv('txt_vector.csv')
    y = pd.read_excel('Policy.xlsx')
    y = y['评分']

    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        mlp = MLPClassifier()
        mlp.fit(X_train, y_train)

        y_pred = mlp.predict(X_test)
        y_pred_prob = mlp.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print(exp)
    print(stab)


# 随机森林
def RandomForestPred(x):
    X = pd.read_csv('txt_vector.csv')
    y = pd.read_excel('Policy.xlsx')
    y = y['评分']

    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        rfc = RandomForestClassifier()
        rfc.fit(X_train, y_train)

        y_pred = rfc.predict(X_test)
        y_pred_prob = rfc.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print(exp)
    print(stab)
    return 0


# LDA模型法
def LDAPred(n):
    X = pd.read_csv('txt_vector.csv')
    y = pd.read_excel('Policy_0.xlsx')
    y = y['评分']
    cntVector = CountVectorizer()
    # CountVectorizer 直接生成的是一个csr_matrix对象，可以通过cntTf.toarray()变成词频向量
    cntTf = cntVector.fit_transform(words)
    vocs = cntVector.get_feature_names()

    lda = LatentDirichletAllocation(n_components=2)
    # 计算 doc-topic probability matrix
    docres = lda.fit_transform(cntTf)
    LDA_corpus = np.array(docres)
    print('类别所属概率:\n', LDA_corpus)
    # 每篇文章中对每个特征词的所属概率矩阵：list长度等于分类数量
    # print('主题词所属矩阵：\n', lda.components_)
    # 构建一个零矩阵
    LDA_corpus_one = np.zeros([LDA_corpus.shape[0]])
    # 对比所属两个概率的大小，确定属于的类别
    LDA_corpus_one = np.argmax(LDA_corpus, axis=1)
    # 返回沿轴axis最大值的索引，axis=1代表行；最大索引即表示最可能表示的数字是多少
    print('每个文档所属类别：', LDA_corpus_one)

    # 计算 topic-term probability matrix
    tt_matrix = lda.components_
    id = 0
    for tt_m in tt_matrix:
        tt_dict = [(name, tt) for name, tt in zip(vocs, tt_m)]
        tt_dict = sorted(tt_dict, key=lambda x: x[1], reverse=True)
        # 打印权重值大于0.6的主题词：
        # tt_dict = [tt_threshold for tt_threshold in tt_dict if tt_threshold[1] > 0.6]
        # 打印每个类别前5个主题词：
        tt_dict = tt_dict[:8]
        print('主题%d:' % id, tt_dict)
        id += 1

    score = accuracy_score(LDA_corpus_one, y)
    print(score)


# # NaiveBayesianPred()
# # LogisticRegressionPred()
# # NeuralNetworkPred()
# # # RandomForestPred()
# time_start2 = time.time()
# # LDAPred(5)
# time_end2 = time.time()
# print('LDA training completed' + "Time cost: {:.2f} s".format(time_end2 - time_start2))


ff = pd.DataFrame()
for i, row in df.iterrows():
    if df.iloc[i, 0] not in list(data['id']):
        strip = pd.DataFrame(df.iloc[i, :].values, index=df.columns, columns=[df.iloc[i, 0]])
        ff = ff.append(pd.DataFrame(strip.values.T, index=strip.columns, columns=strip.index))

ff.sort_index(inplace=True)
ff.drop(['id'], axis=1, inplace=True)

app2 = xw.App(visible=False, add_book=False)
try:
    wb = app2.books.open('筛选后数据.xlsx')
    sht = wb.sheets.add('其他样本')
    sht['A1'].value = ff
    wb.save()
finally:
    app2.quit()


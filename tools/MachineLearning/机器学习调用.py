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

# 导入原数据
Sample = pd.read_excel('./调试数据.xlsx', sheet_name='Sheet1')
# 导入分词统计后的 DTM
df = pd.read_excel('./PanelDataSample.xlsx', sheet_name='overall_DTM_key')
df.set_index(df['id'], inplace=True)
df.drop(['id'], inplace=True, axis=1)

X = df
y = Sample['监管强度']

# 朴素贝叶斯法
def NaiveBayesianPred(x):
    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        nb = BernoulliNB()
        nb.fit(X_train, y_train)

        y_pred = nb.predict(X_test)
        y_pred_prob = nb.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        print('{0} times score: {1}'.format(i, score))
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print('NaiveBayes Average Score:' + str(exp))
    print('Volatility of the prediction: ' + str(stab))


# 逻辑回归法
def LogisticRegressionPred(x):
    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        lr = LogisticRegression()
        lr.fit(X_train, y_train)

        y_pred = lr.predict(X_test)
        y_pred_prob = lr.predict_proba(X_test)
        score = accuracy_score(y_test, y_pred)
        print('{0} times score: {1}'.format(i, score))
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print('LogisticRegression Average Score:' + str(exp))
    print('Volatility of the prediction: ' + str(stab))


# 神经网络法
def NeuralNetworkPred(x):
    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        mlp = MLPClassifier()
        mlp.fit(X_train, y_train)

        y_pred = mlp.predict(X_test)
        y_pred_prob = mlp.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        print('{0} times score: {1}'.format(i, score))
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print('Network Average Score:' + str(exp))
    print('Volatility of the prediction: ' + str(stab))


# 随机森林
def RandomForestPred(x):
    list = []
    for i in range(1, x):
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

        rfc = RandomForestClassifier()
        rfc.fit(X_train, y_train)

        y_pred = rfc.predict(X_test)
        y_pred_prob = rfc.predict_proba(X_test)

        score = accuracy_score(y_test, y_pred)
        print('{0} times score: {1}'.format(i, score))
        list = list + [score]
    exp = np.average(list)
    stab = np.var(list)
    print('RandomForest Average Score:' + str(exp))
    print('Volatility of the prediction: ' + str(stab))

# 括号内的 x 为训练次数
NaiveBayesianPred(5)
LogisticRegressionPred(5)
NeuralNetworkPred(5)
RandomForestPred(5)



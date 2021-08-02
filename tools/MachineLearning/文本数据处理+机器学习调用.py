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

jieba.load_userdict('E:\\ANo.3\\base\\add_words_dict.txt')
os.chdir("E:/ANo.3/base")


# def dataprocessing():
global words
os.chdir('E:/ANo.3/base/LDA明细5-70类_1-3次')

# 从excel中提取数据
app1 = xw.App(visible=False, add_book=False)
try:
    wb = app1.books.open("大于等于1次LDA明细.xlsx")
    sht = wb.sheets['LDA70类']
    df = sht.used_range.value
    df = pd.DataFrame(df)

    df.columns = list(df.loc[0])
    df.drop(0, axis=0, inplace=True)

    df.reset_index(inplace=True, drop=True)
    tf = pd.DataFrame(df.iloc[:, 1])
    tf.reset_index(inplace=True, drop=True)
finally:
    app1.quit()

# 读取关键词清单
m = open('add_words_dict.txt', 'r', encoding='utf-8')
n = m.read()
m.close()
n = n.split('\n')
# n = pd.DataFrame(n)
# n.dropna(axis=0, inplace=True)

words = []
for i, row in tf.iterrows():
    try:
        result = row['正文']
        rule = re.compile(u'[^\u4e00-\u9fa5]')
        result = rule.sub('', result)
        words.append(result)
    except TypeError:
        print(i)
    continue

ff = pd.DataFrame(words)

h = int(len(words) / 20)

words = []
t = 0
for i, row in ff.iterrows():
    item = row[0]
    result = jieba.cut(item)
    word = " ".join(result)
    words.append(word)

    t += 1
    if t % h == 0:
        print("\r文本处理进度：{0}{1}%  ".format("■" * int(t / h), int(5 * t / h)), end='')

words = pd.DataFrame(words)
# 取一个words的备份
hw = words
# words.insert(0, "评分", list(df["评分"]))
# words.to_excel("Cut_policy.xlsx")
words = list(words[0])

# CountVectorizer() 可以自动完成词频统计，通过fit_transform生成文本向量和词袋库
vect = CountVectorizer()
X = vect.fit_transform(words)
X = X.toarray()
# 二维ndarray可以展示在pycharm里，但是和DataFrame性质完全不同
# ndarray 没有 index 和 column
features = vect.get_feature_names()
XX = pd.DataFrame(X, index=tf['id'], columns=features)

shell = pd.DataFrame()
for item in lst:
    if item in n:
        shell = pd.concat([shell, XX[item]], axis=1)
shell['Sum'] = shell.sum(axis=1)
# shell.to_excel('RandalTF.xlsx')
# df['Sum'] = shell['Sum']
# 错了可以在这里改名
# shell.rename(index={-1:'sum'}, inplace=True)

# k = pd.DataFrame(words_bag, index=['number'])
# k.to_csv("words_bag.csv", encoding='utf_8_sig')
# # # 面向列的创建方法
# k = pd.DataFrame.from_dict(words_bag)
# pd.DataFrame(X).to_csv("txt_vector.csv", index=False)


# def freq_drawer(tf, shell):
x = shell[['Sum']]
x.insert(0, 'doc-id', x.index)
freq = {'>=0': 0, '>=1': 0, '>=2': 0, '>=3': 0, '>=4': 0, '>=5': 0, '>=6': 0, '>=7': 0, '>=8': 0, '>=9': 0, '>=10': 0, '>=11': 0}

# 功能一是根据关键词频率把文档分类计数，功能二是把分出来的文档写入excel
app2 = xw.App(visible=False, add_book=False)
wb = app2.books.open('样本分筛.xlsx')
try:
    for i, row in x.iterrows():
        if row['Sum'] >= 0:
            freq['>=0'] += 1
        if row['Sum'] >= 1:
            freq['>=1'] += 1
        if row['Sum'] >= 2:
            freq['>=2'] += 1
        if row['Sum'] >= 3:
            freq['>=3'] += 1
        if row['Sum'] >= 4:
            freq['>=4'] += 1
        if row['Sum'] >= 5:
            freq['>=5'] += 1
        if row['Sum'] >= 6:
            freq['>=6'] += 1
        if row['Sum'] >= 7:
            freq['>=7'] += 1
        if row['Sum'] >= 8:
            freq['>=8'] += 1
        if row['Sum'] >= 9:
            freq['>=9'] += 1
        if row['Sum'] >= 10:
            freq['>=10'] += 1
        if row['Sum'] >= 11:
            freq['>=11'] += 1

    # sht = wb.sheets('>=0次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 0].index]
    # sht = wb.sheets('>=1次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 1].index]
    # sht = wb.sheets('>=2次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 2].index]
    # sht = wb.sheets('>=3次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 3].index]
    # sht = wb.sheets('>=4次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 4].index]
    # sht = wb.sheets('>=5次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 5].index]
    # sht = wb.sheets('>=6次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 6].index]
    # sht = wb.sheets('>=7次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 7].index]
    # sht = wb.sheets('>=8次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 8].index]
    # sht = wb.sheets('>=9次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 9].index]
    # sht = wb.sheets('>=10次')
    # sht.range('A1').value = df.loc[x[x['Sum'] >= 10].index]
    # sht = wb.sheets('>=11次')
    # sht.range('A1').value = df.loc[x[x['Sum'] > 10].index]

    wb.save()
finally:
    app2.quit()

# from_dict 创建 纵向的 dataframe
fre = pd.DataFrame.from_dict(freq, orient='index')
# 创建横向的 dataframe
# fre = pd.DataFrame(freq, index=[0])

# fre.drop('0-2', inplace=True, axis=0)
plt.bar(fre.index, fre[0])
plt.title('Histogram of Docs')
plt.ylabel('Number of Docs')
plt.xlabel('Frequency of keywords')
for a, b in zip(fre.index, fre[0]):
    plt.text(a, b + 0.05, '%.0f' % b, ha='center', va='bottom', fontsize=10)
plt.grid(True)
plt.show()

# time_start1 = time.time()
# shell = dataprocessing()
# freq_drawer(tf, shell)
# time_end1 = time.time()
# print('Data processing completed' + "Time cost: {:.2f} s".format(time_end1 - time_start1))


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
    # X = pd.read_csv('txt_vector.csv')
    # y = pd.read_excel('Policy_0.xlsx')
    # y = y['评分']
    # cntVector = CountVectorizer()
    # # CountVectorizer 直接生成的是一个csr_matrix对象，可以通过cntTf.toarray()变成词频向量
    # cntTf = cntVector.fit_transform(words)
    # vocs = cntVector.get_feature_names()
    #
    # lda = LatentDirichletAllocation(n_components=2)
    # # 计算 doc-topic probability matrix
    # docres = lda.fit_transform(cntTf)
    # LDA_corpus = np.array(docres)
    # print('类别所属概率:\n', LDA_corpus)
    # # 每篇文章中对每个特征词的所属概率矩阵：list长度等于分类数量
    # # print('主题词所属矩阵：\n', lda.components_)
    # # 构建一个零矩阵
    # LDA_corpus_one = np.zeros([LDA_corpus.shape[0]])
    # # 对比所属两个概率的大小，确定属于的类别
    # LDA_corpus_one = np.argmax(LDA_corpus, axis=1)
    # # 返回沿轴axis最大值的索引，axis=1代表行；最大索引即表示最可能表示的数字是多少
    # print('每个文档所属类别：', LDA_corpus_one)
    #
    # # 计算 topic-term probability matrix
    # tt_matrix = lda.components_
    # id = 0
    # for tt_m in tt_matrix:
    #     tt_dict = [(name, tt) for name, tt in zip(vocs, tt_m)]
    #     tt_dict = sorted(tt_dict, key=lambda x: x[1], reverse=True)
    #     # 打印权重值大于0.6的主题词：
    #     # tt_dict = [tt_threshold for tt_threshold in tt_dict if tt_threshold[1] > 0.6]
    #     # 打印每个类别前5个主题词：
    #     tt_dict = tt_dict[:8]
    #     print('主题%d:' % id, tt_dict)
    #     id += 1

    # score = accuracy_score(LDA_corpus_one, y)
    # print(score)

    # # 现在开始对 LDA 的分类结果做可视化分析
    # # 计算 topic-term probability matrix
    # tt_m = lda.components_

    # # 计算 doc-topic probability matrix
    # dt_m = lda.fit_transform(cntTf)

    # 写入记事本的一个小代码，为了看每个 doc 到底有多长(记得要写上‘w’开权限）
    # with open('words[0].txt', 'w') as f:
    #     f.write(cc)
    #     f.close()

    # pyLDAvis.save_html(d, 'lda_pass10.html')  # 将结果保存为该html文件
    # # 接下来获取第三个参数：每个文本的词数统计
    # doc_len = []
    # for item in words:
    #     cut_item = item.split(' ')
    #     doc_len.append(len(cut_item))

    # 接下来整一个 words_bag 的 array

    # 接下来整一个词频统计

    # 以下是大佬写的全自动lda训练代码
    print('LDA Training now......')
    tf_vectorizer = CountVectorizer(strip_accents='unicode',
                                    stop_words='english',
                                    lowercase=True,
                                    token_pattern=r'\b[\u4e00-\u9fa5]{2,}\b',
                                    max_df=0.5,
                                    min_df=10)
    dtm_tf = tf_vectorizer.fit_transform(words)
    tfidf_vectorizer = TfidfVectorizer(**tf_vectorizer.get_params())
    dtm_tfidf = tfidf_vectorizer.fit_transform(words)
    lda_tf = LatentDirichletAllocation(n, random_state=0)
    lda_tf.fit(dtm_tf)
    # for TFIDF DTM
    lda_tfidf = LatentDirichletAllocation(n, random_state=0)
    lda_tfidf.fit(dtm_tfidf)

    d = pyLDAvis.sklearn.prepare(lda_tf, dtm_tf, tf_vectorizer)

    pyLDAvis.save_html(d, '银保监_lda_'+str(n)+'.html')


# # NaiveBayesianPred()
# # LogisticRegressionPred()
# # NeuralNetworkPred()
# # # RandomForestPred()
# time_start2 = time.time()
# # LDAPred(5)
# time_end2 = time.time()
# print('LDA training completed' + "Time cost: {:.2f} s".format(time_end2 - time_start2))

# 按照 id 筛选一下样本
f = r'E:/ANo.3/base/id_total.dta'
data = pd.read_stata(f)
# df.insert(0, 'id', df.index)
tf = pd.DataFrame()
for item in list(data['id']):
    tf = tf.append(pd.DataFrame(df[df['id'] == item].values))
tf.index = tf['id']
tf.columns = df.columns
tf.sort_index(inplace=True)
tf.drop(['id'], axis=1, inplace=True)
tf.to_excel('筛选后数据.xlsx')

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


# 这里比较了三种循环之间的差异，总体来说差距不小
# 同时执行一亿次，前一种弹性较小，每次运行都在16-17s之间
# 后两种都在20s-50s不等，而且随着实验次数的增加，耗时大大加长
# 总体来说，第三种优于第二种，调用已经编译好的函数，归根结底，都要比重新编译快得多

t1 = time.time()
for i in range(0, 100000000):
    x = 1
    y = x
t2 = time. time()
print(f'Time cost:  {t2 - t1}')
#
#
t1 = time.time()
for i in range(0, 100000000):
    def inner_func():
        x = 1
        y = x
        return y
    m = inner_func()
t2 = time. time()
print(f'Time cost:  {t2 - t1}')


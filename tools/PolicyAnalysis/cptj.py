"""
Coding: UTF-8
Author: Randal
Time: 2021/2/20
E-mail: RandalJin@163.com

Description: This is a simple toolkit for data extraction of text.
The most important function in the script is about word frequency statistics.
Using re, I generalized the process in words counting, regardless of  any preset
word segmentation. Besides, many interesting functions, like getting top sentences are built here.

All rights reserved.

"""

import xlwings as xw
import pandas as pd
import numpy as np
import os
import re
from alive_progress import alive_bar
from alive_progress import show_bars, show_spinners
import jieba
import datetime
from sklearn.feature_extraction.text import CountVectorizer, TfidfVectorizer
import math


class jieba_vectorizer(CountVectorizer):

    def __init__(self, tf, userdict, stopwords, orient=False):
        """
        :param tf: 输入的样本框，{axis: 1, 0: id, 1: 标题, 2: 正文, 3: 来源, 4: freq}
        :param stopwords: 停用词表的路径
        :param user_dict_link: 关键词清单的路径
        :param orient: {True: 返回的 DTM 只包括关键词清单中的词，False: 返回 DTM 中包含全部词语}
        :return: 可以直接使用的词向量样本
        """
        self.userdict = userdict
        self.orient = orient
        self.stopwords = stopwords
        jieba.load_userdict(self.userdict)  # 载入关键词词典

        tf = tf.copy()  # 防止对函数之外的原样本框造成改动
        print('切词中，请稍候……')
        rule = re.compile(u'[^\u4e00-\u9fa5]')  # 清洗所有样本，只保留汉字
        for i in range(0, tf.shape[0]):
            try:
                tf.iloc[i, 2] = rule.sub('', tf.iloc[i, 2])
            except TypeError:
                print('样本清洗Error: doc_id = ' + str(i))
            continue

        if self.stopwords is not None:
            stopwords = txt_to_list(self.stopwords)  # 载入停用词表
        else:
            stopwords = []

        # 开始切词

        words = []
        items = range(0, len(tf))
        with alive_bar(len(items), force_tty=True, bar='circles') as bar:
            for i, row in tf.iterrows():
                item = row['正文']
                result = jieba.cut(item)
                # 同时过滤停用词
                word = ''
                for element in result:
                    if element not in stopwords:
                        if element != '\t':
                            word += element
                            word += " "
                words.append(word)
                bar()

        # CountVectorizer() 可以自动完成词频统计，通过fit_transform生成文本向量和词袋库
        # 如果需要换成 tfidfVectorizer, 把下面三行修改一下就可以了
        vect = CountVectorizer()
        X = vect.fit_transform(words)
        self.vectorizer = vect

        matrix = X
        X = X.toarray()
        # 二维ndarray可以展示在pycharm里，但是和DataFrame性质完全不同
        # ndarray 没有 index 和 column
        features = vect.get_feature_names()
        XX = pd.DataFrame(X, index=tf['id'], columns=features)

        self.DTM0 = matrix
        self.DTM = XX
        self.features = features

        # # 下面是之前走的弯路，不足一哂
        # words_bag = vect.vocabulary_
        # # 字典的转置（注意只适用于vk一一对应的情况，1v多k请参考setdefault)
        # bag_words = dict((v, k) for k, v in words_bag.items())
        #
        # # 字典元素的排列顺序不等于字典元素值的排列顺序
        # lst = []
        # for i in range(0, len(XX.columns)):
        #     lst.append(bag_words[i])
        # XX.columns = lst


        if orient:
            dict_filter = txt_to_list(self.userdict)
            for word in features:
                if word not in dict_filter:
                    XX.drop([word], axis=1, inplace=True)
            self.DTM_key = XX

    def get_feature_names(self):
        return self.features

    def strip_non_keywords(self, df):
        ff = df.copy()
        dict_filter = txt_to_list(self.userdict)
        for word in self.features:
            if word not in dict_filter:
                ff.drop([word], axis=1, inplace=True)
        return ff


def make_doc_freq(word, doc):
    """
    :param word: 指的是要对其进行词频统计的关键词
    :param doc: 指的是要遍历的文本
    :return: lst: 返回字典，记录关键词在文本当中出现的频次以及上下文
    """
    # 使用正则表达式进行匹配, 拼接成pattern

    # re.S表示会自动换行
    # finditer是findall的迭代器版本，通过遍历可以依次打印出子串所在的位置
    it = re.finditer(word, doc, re.S)
    # match.group()可以返回子串，match.span()可以返回索引
    lst = []
    for match in it:
        lst.append(match.span())
    freq = dict()
    freq['Frequency'] = len(lst)
    # 将上下文结果也整理为一个字典
    context = dict()
    for i in range(0, len(lst)):
        # 将span的范围前后各扩展不多于10个字符，得到上下文
        try:
            # 为了划出适宜的前后文范围，需要设定索引的最大值和最小值
            # 因此要比较span+10和doc极大值，span-10和doc极小值
            # 最大值在两者间取小，最小值在两者间取大
            MAX = min(lst[i][1] + 10, len(doc))
            MIN = max(0, lst[i][0] - 10)
            # 取得上下文
            context[str(i)] = doc[MIN: MAX]
        except IndexError:
            print('IndexError: ' + word)
    freq['Context'] = context
    return freq


def make_info_freq(name, pattern, doc):
    """
    :param name: 指的是对其进行词频统计的形式
    :param pattern: 指的是对其进行词频统计的正则表达式
    :param doc: 指的是要遍历的文本
    :return: lst: 返回字典，记录关键词在文本当中出现的频次以及上下文
    注：该函数返回字典中的context元素为元组：（关键词，上下文）
    """
    # 使用正则表达式进行匹配, 拼接成pattern
    # re.S表示会自动换行
    # finditer是findall的迭代器版本，通过遍历可以依次打印出子串所在的位置
    it = re.finditer(pattern[0], doc, re.S)
    # match.group()可以返回子串，match.span()可以返回索引
    cls = pattern[1]
    lst = []
    for match in it:
        lst.append(match.span())
    freq = dict()
    freq['Frequency'] = len(lst)
    freq['Name'] = name
    # 将上下文结果也整理为一个字典
    context = dict()
    for i in range(0, len(lst)):
        # 将span的范围前后各扩展不多于10个字符，得到上下文
        try:
            # 为了划出适宜的前后文范围，需要设定索引的最大值和最小值
            # 因此要比较span+10和doc极大值，span-10和doc极小值
            # 最大值在两者间取小，最小值在两者间取大
            MAX = min(lst[i][1] + 10, len(doc))
            MIN = max(0, lst[i][0] - 10)
            # 取得匹配到的关键词，并做掐头去尾处理
            word = match_cut(doc[lst[i][0]: lst[i][1]], cls)
            # 将关键词和上下文打包，存储到 context 条目中
            context[str(i)] = (word, doc[MIN: MAX])
        except IndexError:
            print('IndexError: ' + name)
    freq['Context'] = context
    return freq


def make_docs_freq(word, docs):
    """
    :param word: 指的是要对其进行词频统计的关键词
    :param docs: 是要遍历的文本的集合，必须是pandas DataFrame的形式，至少包含id列 (iloc: 0)，正文列 (iloc: 2) 和预留出的频次列 (iloc: 4)
    :return: 返回字典，其中包括“单关键词-单文本”的词频字典集合，以及计数结果汇总
    """
    freq = dict()
    # 因为总频数是通过"+="的方式计算，不是简单赋值，所以要预设为0
    freq['Total Frequency'] = 0
    docs = docs.copy()  # 防止对函数之外的原样本框造成改动

    for i in range(0, len(docs)):
        # 对于每个文档，都形成一个字典，字典包括关键词在该文档出现的频数和上下文
        # id需要在第0列，正文需要在第2列
        freq['Doc' + str(docs.iloc[i, 0])] = make_doc_freq(word, docs.iloc[i, 2])
        # 在给每个文档形成字典的同时，对于总概率进行滚动加总
        freq['Total Frequency'] += freq['Doc' + str(docs.iloc[i, 0])]['Frequency']
        docs.iloc[i, 4] = freq['Doc' + str(docs.iloc[i, 0])]['Frequency']

    # 接下来建立一个DFC(doc-freq-context)统计面板，汇总所有文档对应的词频数和上下文
    # 首先构建(id, freq)的字典映射
    xs = docs['id']
    ys = docs['freq']
    # zip(迭代器)是一个很好用的方法，建议多用
    id_freq = {x: y for x, y in zip(xs, ys)}

    # 新建一个空壳DataFrame，接下来把数据一条一条粘贴进去
    data = pd.DataFrame(columns=['id', 'freq', 'word', 'num', 'context'])
    for item in xs:
        doc = freq['Doc' + str(item)]
        num = doc['Frequency']
        context = doc['Context']
        for i in range(0, num):
            strip = {'id': item, 'freq': id_freq[item], 'word': word, 'num': i, 'context': context[str(i)]}
            # 默认orient参数等于columns
            # 如果字典的值是标量，那就必须传递一个index，这是规定
            strip = pd.DataFrame(strip, index=[None])
            # df的append方法只能通过重新赋值来进行修改
            data = data.append(strip)
    data.set_index(['id', 'freq', 'word'], drop=True, inplace=True)
    freq['DFC'] = data
    return freq


def make_infos_freq(name, pattern, docs):
    """
    :param name: 指的是对其进行词频统计的形式
    :param pattern: 指的是对其进行词频统计的（正则表达式, 裁剪方法）
    :param docs: 是要遍历的文本的集合，必须是pandas DataFrame的形式，至少包含id列(iloc: 0)和正文列(iloc: 2)
    :return: 返回字典，其中包括“单关键词-单文本”的词频字典集合，以及计数结果汇总
    """
    freq = dict()
    # 因为总频数是通过"+="的方式计算，不是简单赋值，所以要预设为0
    freq['Total Frequency'] = 0
    docs = docs.copy()  # 防止对函数之外的原样本框造成改动

    items = range(0, len(docs))
    with alive_bar(len(items), force_tty=True, bar='circles') as bar:
        for i in items:
            # 对于每个文档，都形成一个字典，字典包括关键词在该文档出现的频数和上下文
            # id需要在第0列，正文需要在第2列
            # pattern 要全须全尾地传递进去，因为make_info_freq两个参数都要用
            freq['Doc' + str(docs.iloc[i, 0])] = make_info_freq(name, pattern, docs.iloc[i, 2])
            # 在给每个文档形成字典的同时，对于总概率进行滚动加总
            freq['Total Frequency'] += freq['Doc' + str(docs.iloc[i, 0])]['Frequency']
            docs.iloc[i, 4] = freq['Doc' + str(docs.iloc[i, 0])]['Frequency']
            bar()

    # 接下来建立一个DFC(doc-freq-context)统计面板，汇总所有文档对应的词频数和上下文
    # 首先构建(id, freq)的字典映射
    xs = docs['id']
    ys = docs['freq']
    # zip(迭代器)是一个很好用的方法，建议多用
    id_freq = {x: y for x, y in zip(xs, ys)}

    # 新建一个空壳DataFrame，接下来把数据一条一条粘贴进去
    data = pd.DataFrame(columns=['id', 'freq', 'form', 'word', 'num', 'context'])
    for item in xs:
        doc = freq['Doc' + str(item)]
        num = doc['Frequency']
        # 从（关键词，上下文）中取出两个元素
        context = doc['Context']
        for i in range(0, num):
            # context 中的关键词已经 match_cut 完毕，不需要重复处理
            strip = {'id': item, 'form': name, 'freq': id_freq[item], 'word': context[str(i)][0],
                     'num': i, 'context': context[str(i)][1]}
            # 默认orient参数等于columns
            # 如果字典的值是标量，那就必须传递一个index，这是规定
            strip = pd.DataFrame(strip, index=[None])
            # df的append方法只能通过重新赋值来进行修改
            data = data.append(strip)
    data.set_index(['id', 'freq', 'form', 'word'], drop=True, inplace=True)
    freq['DFC'] = data
    print(name + '  Completed')
    return freq


def words_docs_freq(words, docs):
    """
    :param words: 表示要对其做词频统计的关键词清单
    :param docs: 是要遍历的文本的集合，必须是pandas DataFrame的形式，至少包含id列、正文列、和频率列
    :return: 返回字典，其中包括“单关键词-多文本”的词频字典集合，以及最终的DFC(doc-frequency-context)和DTM(doc-term matrix)
    """
    freqs = dict()
    # 与此同时新建一个空壳DataFrame，用于汇总DFC
    data = pd.DataFrame()
    # 新建一个空壳，用于汇总DTM(Doc-Term-Matrix)
    dtm = pd.DataFrame(None, columns=words, index=docs['id'])
    # 来吧，一个循环搞定所有
    items = range(len(words))
    with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
        for word in words:
            freq = make_docs_freq(word, docs)
            freqs[word] = freq
            data = data.append(freq['DFC'])
            for item in docs['id']:
                dtm.loc[item, word] = freq['Doc' + str(item)]['Frequency']
            bar()
    # 记得要sort一下，不然排序的方式不对（应该按照doc id来排列）
    data.sort_index(inplace=True)

    freqs['DFC'] = data
    freqs['DTM'] = dtm
    return freqs


def infos_docs_freq(infos, docs):
    """
    :param docs: 是要遍历的文本的集合，必须是pandas DataFrame的形式，至少包含id列和正文列
    :param infos: 指的是正则表达式的列表，格式为字典，key是示例，如“（1）”，value 是正则表达式，如“（[0-9]）”
    :return: 返回字典，其中包括“单关键词-多文本”的词频字典集合，以及最终的DFC(doc-frequency-context)和DTM(doc-term matrix)
    """
    freqs = dict()
    # 与此同时新建一个空壳DataFrame，用于汇总DFC
    data = pd.DataFrame()
    # 新建一个空壳，用于汇总DTM(Doc-Term-Matrix)
    dtm = pd.DataFrame(None, columns=list(infos.keys()), index=docs['id'])
    # 来吧，一个循环搞定所有
    items = range(len(infos))
    with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
        for k, v in infos.items():
            freq = make_infos_freq(k, v, docs)
            freqs[k] = freq
            data = data.append(freq['DFC'])
            for item in docs['id']:
                dtm.loc[item, k] = freq['Doc' + str(item)]['Frequency']
            bar()
    # 记得要sort一下，不然排序的方式不对（应该按照doc id来排列）
    data.sort_index(inplace=True)

    freqs['DFC'] = data
    freqs['DTM'] = dtm
    return freqs


def massive_pop(infos, doc):
    """
    :param infos: List，表示被删除内容对应的正则表达式
    :param doc: 表示正文
    :return: 返回一个完成删除的文本
    """
    for info in infos:
        doc = re.sub(info, '', doc)
    return doc


def massive_sub(infos, doc):
    """
    :param infos: Dict, 表示被替换内容对应的正则表达式及替换对象
    :param doc: 表示正文
    :return: 返回一个完成替换的文本
    """
    for v, k in infos:
        doc = re.sub(v, k, doc)
    return doc


# 接下来取每个样本的前n句话(或者不多于前n句话的内容)，再做一次进行对比
# 取前十句话的原理是，对！？。等表示语义结束的符号进行计数，满十次为止
def top_n_sent(n, doc, percentile=1):
    """
    :param n: n指句子的数量，这个函数会返回一段文本中前n句话，若文本内容不多于n句，则全文输出
    :param word: 指正文内容
    :param percentile: 按照分位数来取句子时，要输入的分位，比如一共有十句话，取50%分位就是5句
    如果有11句话，向下取整也是输出5句
    :return: 返回字符串：前n句话
    """
    info = '[。？！]'
    # 在这个函数体内，函数主体语句的作用域大于循环体，因此循环内的变量相当于局部变量
    # 因此想在循环外直接返回，就会出现没有定义的错误，因此可以做一个全局声明
    # 但是不建议这样做，因为如果函数外有一个变量恰巧和局部变量重名，那函数外的变量也会被改变
    # 因此还是推荐多使用迭代器，把循环包裹成迭代器，可以解决很多问题
    # 而且已经封装好的迭代器，例如re.findall_iter，就不用另外再去写了，调用起来很方便
    # 如下，第一行代码的作用是用列表包裹迭代器，形成一个生成器的列表
    # 每个生成器都存在自己的 Attribute
    re_iter = list(re.finditer(info, doc))
    # max_iter 是 re 匹配到的最大次数
    max_iter = len(re_iter)

    # 这一句表示，正文过于简短，或者没有标点，此时直接输出全文
    if max_iter == 0:
        return doc

    # 考虑 percentile 的情况，如果总共有11句，就舍弃掉原来的 n，直接改为总句数的 percentile 对应的句子数
    # 注意是向下取整
    if percentile != 1:
        n = math.ceil(percentile * max_iter)
        # 如果匹配到至少一句，循环自然结束，输出结果
        if n > 0:
            return doc[0: re_iter[n - 1].end()]
        # 如果正文过于简短，或设定的百分比过低，一句话都凑不齐，此时直接输出第一句
        elif n == 0:
            return doc[0: re_iter[0].end()]

    # 如果匹配到的句子数大于 n，此时只取前 n 句
    if max_iter >= n:
        return doc[0: re_iter[n - 1].end()]
    # 如果匹配到的句子不足 n 句，直接输出全部内容
    elif 0 < max_iter < n:
        return doc[0: re_iter[-1].end()]

    # 为减少重名的可能，尽量在函数体内减少变量的使用


def dtm_sort_filter(dtm, keymap, name=None):
    """
    :param dtm: 前面生成的词频统计矩阵：Doc-Term-Matrix
    :param keymap: 字典，标明了  类别-关键词列表  两者关系
    :param name: 最终生成 Excel 文件的名称（需要包括后缀）
    :return: 返回一个字典，字典包含两个 pandas.DataFrame: 一个是表示各个种类是否存在的二进制表，另一个是最终的种类数
    """
    dtm = dtm.applymap(lambda x: 1 if x != 0 else 0)
    strips = {}
    for i, row in dtm.iterrows():
        strip = {}
        for k, v in keymap.items():
            strip[k] = 0
            for item in v:
                try:
                    strip[k] += row[item]
                except KeyError:
                    pass
        strips[i] = strip
    dtm_class = pd.DataFrame.from_dict(strips, orient='index')
    dtm_class = dtm_class.applymap(lambda x: 1 if x != 0 else 0)
    dtm_final = dtm_class.agg(np.sum, axis=1)
    result = {'DTM_class': dtm_class, 'DTM_final': dtm_final}
    return result


def dtm_point_giver(dtm, keymap, scoremap, name=None):
    """
    :param dtm: 前面生成的词频统计矩阵：Doc-Term-Matrix
    :param keymap: 字典，{TypeA: [word1, word2, word3, ……], TypeB: ……}
    :param scoremap: 字典，标明了  类别-分值 两者关系
    :param name: 最终生成 Excel 文件的名称（需要包括后缀）
    :return: 返回一个 pandas.DataFrame，表格有两列，一列是文本id，一列是文本的分值（所有关键词的分值取最高）
    """
    dtm = dtm.applymap(lambda x: 1 if x != 0 else 0)

    # 非 keymap 中词会被过滤掉
    strips = {}
    for i, row in dtm.iterrows():
        strip = {}
        for k, v in keymap.items():
            strip[k] = 0
            for item in v:
                try:
                    strip[k] += row[item]
                except KeyError:
                    pass
        strips[i] = strip

    dtm_class = pd.DataFrame.from_dict(strips, orient='index')
    dtm_class = dtm_class.applymap(lambda x: 1 if x != 0 else 0)

    # 找到 columns 对应的分值
    keywords = list(dtm_class.columns)
    multiplier = []
    for keyword in keywords:
        multiplier.append(scoremap[keyword])

    # DataFrame 的乘法运算，不会改变其 index 和 columns
    dtm_score = dtm_class.mul(multiplier, axis=1)
    # 取一个最大值来赋分
    dtm_score = dtm_score.agg(np.max, axis=1)
    return dtm_score


def dfc_sort_filter(dfc, keymap, name=None):
    """
    :param dfc: 前面生成的词频统计明细表：Doc-Frequency-Context
    :param keymap: 字典，标明了  关键词-所属种类  两者关系
    :param name: 最终生成 Excel 文件的名称（需要包括后缀）
    :return: 返回一个 pandas.DataFrame，表格有两列，一列是文本id，一列是文本中所包含的业务种类数
    """
    # 接下来把关键词从 dfc 的 Multi-index 中拿出来（这个index本质上就是一个ndarray)
    # 拿出来关键词就可以用字典进行映射
    # 先新建一列class-id，准备放置映射的结果
    dfc.insert(0, 'cls-id', None)

    # 开始遍历
    for i in range(0, len(dfc.index)):
        dfc.iloc[i, 0] = keymap[dfc.index[i][2]]

    # 理论上就可以直接通过 excel 的分类计数功能来看业务种类数了
    # 失败了，excel不能看种类数，只能给所有值做计数，因此还需要借助python的unique语句
    # dfc.to_excel('被监管业务统计.xlsx')

    # 可以对于每一种index做一个计数，使用loc索引到的对象是一个DataFrame
    # 先拿到一个doc id的列表
    did = []
    for item in dfc.index.unique():
        did.append(item[0])
    did = list(pd.Series(did).unique())

    # 接下来获得每一类的结果，注：多重索引的取值值得关注
    uni = {}
    for item in did:
        uni[item] = len(dfc.loc[item, :, :]['cls-id'].unique())

    # 把生成的字典转换为以键值行索引的 DataFrame
    uni = pd.DataFrame.from_dict(uni, orient='index')
    uni.fillna(0, axis=1, inplace=True)
    # uni.to_excel(name)
    return uni


def dfc_point_giver(dfc, keymap, name=None):
    """
    :param dfc: 前面生成的词频统计明细表：Doc-Frequency-Context
    :param keymap: 字典，标明了  关键词-分值 两者关系
    :param name: 最终生成 Excel 文件的名称（需要包括后缀）
    :return: 返回一个 pandas.DataFrame，表格有两列，一列是文本id，一列是文本的分值（所有关键词的分值取最高）
    """
    dfc.insert(0, 'point', None)

    # 开始遍历
    for i in range(0, len(dfc.index)):
        dfc.iloc[i, 0] = keymap[dfc.index[i][2]]

    # 可以对于每一种index做一个计数，使用loc索引到的对象是一个DataFrame
    # 先拿到一个doc id的列表
    did = []
    for item in dfc.index.unique():
        did.append(item[0])
    did = list(pd.Series(did).unique())

    # 接下来获得每一类的结果，注：多重索引的取值值得关注
    uni = {}
    for item in did:
        uni[item] = max(dfc.loc[item, :, :]['point'].unique())

    # 把生成的字典转换为以键值行索引的 DataFrame
    uni = pd.DataFrame.from_dict(uni, orient='index')
    uni.fillna(0, axis=1, inplace=True)
    # uni.to_excel(name)
    return uni


def dfc_sort_counter(dfc, name=None):
    """
    :param dfc: 前面生成的词频统计明细表：Doc-Frequency-Context
    :param name: 最终生成 Excel 文件的名称（需要包括后缀）
    :return: 返回一个 pandas.DataFrame，表格有两列，一列是文本id，一列是文本中所包含的业务种类数
    """
    # 可以对于每一种index做一个计数，使用loc索引到的对象是一个DataFrame
    dfc.insert(0, 'form', None)
    for i in range(0, dfc.shape[0]):
        dfc.iloc[i, 0] = dfc.index[i][2]

    # 先拿到一个doc id的列表
    did = []
    for item in dfc.index.unique():
        did.append(item[0])
    did = list(pd.Series(did).unique())

    # 接下来获得每一类的结果，注：多重索引的取值值得关注
    uni = {}
    for item in did:
        uni[item] = len(dfc.loc[item, :, :, :]['form'].unique())

    # 把生成的字典转换为以键值行索引的 DataFrame
    uni = pd.DataFrame.from_dict(uni, orient='index')
    uni.fillna(0, axis=1, inplace=True)
    # uni.to_excel(name)
    return uni


# 定义一个大写中文数字转阿拉伯数字的函数
CN_NUM = {
    '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '零': 0,
    '壹': 1, '贰': 2, '叁': 3, '肆': 4, '伍': 5, '陆': 6, '柒': 7, '捌': 8, '玖': 9, '貮': 2, '两': 2,
}

CN_UNIT = {
    '十': 10,
    '拾': 10,
    '百': 100,
    '佰': 100,
    '千': 1000,
    '仟': 1000,
    '万': 10000,
    '萬': 10000,
    '亿': 100000000,
    '億': 100000000,
    '兆': 1000000000000,
}


def chinese_to_arabic(cn:str) -> str:
    unit = 0  # current
    ldig = []  # digest
    for cndig in reversed(cn):
        if cndig in CN_UNIT:
            unit = CN_UNIT.get(cndig)
            if unit == 10000 or unit == 100000000:
                ldig.append(unit)
                unit = 1
        else:
            dig = CN_NUM.get(cndig)
            if unit:
                dig *= unit
                unit = 0
            ldig.append(dig)
    if unit == 10:
        ldig.append(10)
    val, tmp = 0, 0
    for x in reversed(ldig):
        if x == 10000 or x == 100000000:
            val += tmp * x
            tmp = 0
        else:
            tmp += x
    val += tmp
    val = str(val)
    return val


def match_cut(matchobj, cls):
    """
    :param matchobj: 输入匹配好的字符串
    :param cls: 这是正则表达式的匹配方法，遵循掐头去尾的原则。
    0：原样奉还，1：掐一去一，2：掐一去二，3：掐二去一，4：掐一去零,  5：掐零去一
    :return: 返回一个处理好的字符串
    """
    if cls == 0:
        return matchobj
    if cls == 1:
        return matchobj[1: -1]
    if cls == 2:
        return matchobj[1: -2]
    if cls == 3:
        return matchobj[2, -1]
    if cls == 4:
        return matchobj[1:]
    if cls == 5:
        return matchobj[: -1]


def re_formatter(file_time):
    """
    这个函数是修改数据格式的一个小工具，会将输入的时间统一由YYYY-DD-MM修改为YYYY/MM/DD
    :param file_time: 输入的时间, 纯字符串格式
    :return: 返回的时间，datetime格式(%YYYY/%MM/%DD)
    """
    # 首先把空格和空行删去
    file_time = str(file_time)
    file_time = file_time.strip('\n\r\t ')
    file_time = file_time[:10]
    # 将字符串转为datetime格式
    file_time = datetime.datetime.strptime(file_time, '%Y-%m-%d')
    # 将datetime转为字符串
    file_time = datetime.datetime.strftime(file_time, '%Y/%m/%d')
    return file_time


def year_month_finder(file_time):
    """
   这个函数是修改数据格式的一个小工具，会将输入的时间统一由YYYY/DD/MM修改为YYYY年M月
    :param file_time: 输入的时间, 纯字符串格式
    :return: 返回的时间，格式为(%YYYY年%M月)
    """
    # 首先把空格和空行删去
    file_time = str(file_time)
    file_time = file_time.strip('\n\r\t ')
    # 将字符串转为datetime格式
    try:
        file_time = datetime.datetime.strptime(file_time, '%Y/%m/%d')
        file_time_year = datetime.datetime.strftime(file_time, '%Y')
        file_time_month = datetime.datetime.strftime(file_time, '%m')
        file_time = file_time_year + 'M' + file_time_month
    except ValueError:
        pass
    return file_time


def month_day_finder(file_time):
    """
   这个函数是修改数据格式的一个小工具，会将输入的时间统一由YYYY/DD/MM修改为M月D日
    :param file_time: 输入的时间, 纯字符串格式
    :return: 返回的时间，格式为(%M%月/%D%日)
    """
    # 首先把空格和空行删去
    file_time = str(file_time)
    file_time = file_time.strip('\n\r\t ')
    # 将字符串转为datetime格式
    try:
        file_time = datetime.datetime.strptime(file_time, '%Y/%m/%d')
        file_time_month = datetime.datetime.strftime(file_time, '%m')
        file_time_day = datetime.datetime.strftime(file_time, '%d')
        file_time = file_time_month + '/' + file_time_day
    except ValueError:
        pass
    return file_time


quarter_map = {'1': '1', '2': '1', '3': '1', '4': '2', '5': '2', '6': '2',
               '7': '3', '8': '3', '9': '3', '10': '4', '11': '4', '12': '4',
               '01': '1', '02': '1', '03': '1', '04': '2', '05': '2', '06': '2',
               '07': '3', '08': '3', '09': '3'}


def quarter_finder(file_time):
    """
    这个函数是修改数据格式的一个小工具，会将输入的时间统一由YYYY/DD/MM修改为YYYY年Q季度
    :param file_time: 输入的时间, 纯字符串格式
    :return: 返回的时间，格式为(%YYYY年%Q季度)
    """
    # 首先把空格和空行删去
    file_time = str(file_time)
    file_time = file_time.strip('\n\r\t ')
    # 将字符串转为datetime格式
    try:
        file_time = datetime.datetime.strptime(file_time, '%Y/%m/%d')
        file_time_year = datetime.datetime.strftime(file_time, '%Y')
        file_time_month = datetime.datetime.strftime(file_time, '%m')
        file_time_month = quarter_map[file_time_month]
        file_time = file_time_year + 'Q' + file_time_month
    except ValueError:
        pass
    return file_time


def dataframe_filter(df, keywords, axis, status):
    """
    :param df: 要过滤的样本框
    :param keywords: list, 关键词清单: 要保留/剔除的 index 或 columns
    :param axis: integer, 过滤的方向，axis=0 为竖直方向，axis=1 为水平方向
    :param status: integer, 过滤的方式，status=0 为只保留清单中的数据，status=1 为只去除清单中的数据
    :return: 过滤好的样本框
    """
    df = df.copy()
    # 如果是只保留清单中的数据
    if status == 0:
        if axis == 1:
            lst = df.columns
            items = range(len(lst))
            with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
                for item in lst:
                    if item not in keywords:
                        df.drop([item], axis=1, inplace=True)
                    bar()
        elif axis == 0:
            lst = df.index
            items = range(len(lst))
            with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
                for item in lst:
                    if item not in keywords:
                        df.drop(item, axis=0, inplace=True)
                    bar()
    # 如果是只去除清单中的数据
    elif status == 1:
        if axis == 1:
            lst = df.columns
            items = range(len(lst))
            with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
                for item in lst:
                    if item in keywords:
                        df.drop([item], axis=1, inplace=True)
                    bar()
        elif axis == 0:
            lst = df.index
            items = range(len(lst))
            with alive_bar(len(items), force_tty=True, bar='blocks') as bar:
                for item in lst:
                    if item in keywords:
                        df.drop(item, axis=0, inplace=True)
                    bar()
    else:
        raise KeyError('Axis not properly set')
    return df


def main():
    # 主程序运行，开始！
    # 读取文件
    os.chdir('E:/ANo.3/base')
    app1 = xw.App(visible=False, add_book=False)
    wb = app1.books.open('20210116样本分筛.xlsx')
    try:
        sht = wb.sheets('0次')
        df = pd.DataFrame(sht.used_range.value)
        df.columns = list(df.loc[0])
        df.drop(0, axis=0, inplace=True)
        df.reset_index(drop=True, inplace=True)
        # 记得把id变成整数，否则之后索引字符化后会有麻烦
        df['id'] = df['id'].apply(lambda x: int(x))
        # 提前插入一列频率，方便遍历时记录结果
        df.insert(4, 'freq', None)

        # 获取关键词清单
        f = open('add_words_dict.txt', 'r', encoding='UTF-8')
        n = f.read()
        f.close()
        n = n.split('\n')

        freqs = words_docs_freq(n, df)
        wb = app1.books.add()
        sht = wb.sheets('Sheet1')
        sht.name = '词频统计表'
        sht['A1'].value = freqs['DTM']
        sht = wb.sheets.add('词频统计明细')
        sht['A1'].value = freqs['DFC']
        wb.save('20210126简单词频统计.xlsx')
    finally:
        app1.quit()


def list_to_txt(name, content, sep = '\n'):
    """
    :param name: 指要创建的 txt 文件名称， 不存在的话会自动创建
    :param content: 指要写入的内容，此处必须为列表
    :param sep: 指列表的分隔符，默认为换行符
    :return: 无返回值
    这个函数可以把一个列表写入txt, 元素之间自动以换行符形式进行粘连
    """
    # 首先把列表使用分隔符进行粘连
    content = sep.join(content)
    # 使用with open方式打开，可以保证结束操作的同时关闭句柄
    with open(name, 'w', encoding='UTF-8') as f:
        f.write(content)


def txt_to_list(path, sep = '\n'):
    """
    :param path: 指要读取的 txt 文件所在位置， 不存在的话会报错，path 支持列表格式
    :param sep: 指列表的分隔符，默认为换行符
    :return: 返回一个列表
    这个函数可以把一个 txt 转化为列表 t, 导入的元素之间需要以制定换行符形式分隔开
    如果同时传入多个txt，则自动拼接在同一个列表当中
    """
    if path is not list:
        path = [path]
    lst = []
    for link in path:
        # 使用with open方式打开，可以保证结束操作的同时关闭句柄
        with open(link, 'r', encoding='UTF-8') as f:
            content = str(f.read())
        content = content.split(sep)
        try:
            content.remove('\n')
            content.remove('\t')
            content.remove('\r')
            content.remove('')
            content.remove(' ')
            content.remove('  ')
        except ValueError:
            pass
        lst.extend(content)
    return lst


def doc_filter(df, target):
    """
    :param df: 要进行样本筛选的数据框
    :param target: 符合要求的id列表
    :return: 只保留符合要求id对应样本的数据框
    """
    for doc_id in df['id']:
        if doc_id not in target:
            df.drop(df[df['id'] == doc_id].index, inplace=True)
    return df


def vec_cos(arr1, arr2):
    """
    :param arr1:  the first np.array
    :param arr2:  the second np.array
    :return:  cos(θ) (θ: the angle between two vectors)

    This measures the similarity between two vectors: the closer the value is to 1, the more similar between two vectors
    """
    num = np.sum(arr1 * arr2)
    den = np.sqrt(sum(arr1**2) * sum(arr2**2))
    sim = num/den
    return sim


def cos_rank(matrix, keymap=None, threshold=0.9):
    """
    :param matrix: (np.array) it contains all the keywords frequency vectors
    :param keymap: (dictionary) you can pass a preset dictionary if you wanna keep the original vector id
    :param threshold: (float) you can pass a threshold of cosθ above which the id will be recorded
    :return df: (pd.DataFrame) the rank of similarity

    This function splits the vectors into pairs and calculate the cos similarity pair by pair.
    Then it will generate a DataFrame that ranks the pairs from 1 to 0
    """
    pairs = []
    for i in range(0, len(matrix), 1):
        for t in range(i+1, len(matrix), 1):
            pairs.append((i, t))
    sim = []
    if keymap is None:
        for pair in pairs:
            sim.append((pair, vec_cos(matrix[pair[0]], matrix[pair[1]])))
    else:
        for pair in pairs:
            sim.append(((keymap[pair[0]], keymap[pair[1]]), vec_cos(matrix[pair[0]], matrix[pair[1]])))

    df = pd.DataFrame(sim)
    df.columns = ['id', 'Similarity']
    df.sort_values('Similarity', axis=0, ascending=False, inplace=True)

    # You can choose desired filtering condition
    df = df[df['Similarity'] > threshold]

    return df


def quarter2date(date):
    # 将字符串 date 转化为 datetime 格式
    quarter = int(re.findall('Q[0-9]', date)[0][1])
    month = {1: 1, 2: 4, 3: 7, 4: 10}
    rule = re.compile('Q[0-9]')
    year = rule.sub('', date)
    return datetime.date(int(year), month[quarter], 1)


if __name__ == '__main__':
    main()

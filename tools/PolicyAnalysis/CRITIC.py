"""
Coding: UTF-8
Author: Victor Xu, Randal Jin
Time: 2021/5/11
E-mail: viktor.p.xu@yandex.ru, RandalJin@163.com

Description: If you are interested in empirically economic research,
CRITIC weighting will be a useful tool to construct a new variable.
(To finish the school paper, we coded it to test our skills, and hopefully it will be of use to you.)

"""

import pandas as pd
import numpy as np
import os

os.chdir('C:/Users/ThinkPad/Desktop')


def segment(df, para, std):
    seg = []
    if std:
        for i in para:
            seg.append((i, mmstd(np.array(df[i]))))
    else:
        for i in para:
            seg.append((i, np.array(df[i])))
    return dict(seg)


def r(x, y):
    x_mean = np.mean(x)
    y_mean = np.mean(y)
    x_std = np.std(x)
    y_std = np.std(y)
    zx = (x - x_mean) / x_std
    zy = (y - y_mean) / y_std
    r = np.sum(zx * zy) / (len(x))
    return r


def weigh(seg):
    tem = []
    result = []
    p_sum = 0
    for k, v in seg.items():
        s = np.std(v)
        rel = 0
        for k1, v1 in seg.items():
            rel += (1 - r(v, v1))
        if rel == 0:
            rel = 1
        else:
            pass
        p = s * rel
        tem.append((k, p, s, rel))
        p_sum += p
    for i in range(len(tem)):
        w = tem[i][1] / p_sum
        result.append((tem[i][0], w))
    return [dict(result), tem]


def mmstd(x):
    x_min = np.min(x)
    x_max = np.max(x)
    x_new = (x - x_min) / (x_max - x_min)
    return list(x_new)


def multiply(seg, weight):
    result = dict()
    for k, v in seg.items():
        result[k + '_w'] = list(np.array(v) * weight[k])
    return result


def agg(dic):
    title = ''
    res = []
    for v in range(len(list(dic.values())[0])):
        agg = 0
        title = ''
        for k in list(dic.keys()):
            agg += dic[k][v]
            title += k
        res.append(agg)
    ri = {title: res}
    return ri


def critic(df, para, std=True):
    """
    :param df: pd.DataFrame, the panel data you use CRITIC method with
    :param para: dictionary, the way we split variables into different sets for the weighting
    :param std:  whether or not to standardize the input data (max_min_standardization)
    :return: df: pd.DataFrame, the DataFrame with CRITIC weight and combined variable inserted
    """

    for v, k in para.items():
        seg = segment(df, k, std)
        weight = weigh(seg)[0]
        mul = multiply(seg, weight)
        add = agg(mul)
        for var in k:
            df.insert(df.shape[1], var + '_w', weight[var])
        df.insert(df.shape[1], v, list(add.values())[0])
    return df


# Usage Sample

# -------------------------------------------------------------------
#                 |    GDP                  |
#   Econ      ---------------------------------------------------------
#   Index    |    Inflation           |
#                  --------------------------------------------------------
#                 |   Employment    |
# -------------------------------------------------------------------

# para = {'Econ Index': ['GDP', 'Inflation', 'Employment']}
#
# f = 'C:/Users/ThinkPad/Desktop/data.xlsx'
# df = pd.read_excel(f)
# df = critic(df, para, std=True)

para1 = {"政策主体": ['是否联合发布', '颁布主体得分'],
         "覆盖范围广度": ['被监管业务种类数', '被监管机构种类数'],
         "政策基调": ['相对情感语调', '绝对情感语调'],
         "措施细化深度": ['标题级别数', '标题总个数', '数字个数'],
         "政策内容": ['覆盖范围广度', '措施细化深度'],
         "政策强度": ["政策基调", "政策主体", "政策内容"]}

df = pd.read_excel('C:/Users/ThinkPad/Desktop/20210517数据.xlsx', sheet_name='原始数据')
df = critic(df, para1, std=True)  # std=True 表示每次赋权产生的新指标都自动做一次最大最小值标准化
df. to_excel('C:/Users/ThinkPad/Desktop/CRITIC赋权结果.xlsx')

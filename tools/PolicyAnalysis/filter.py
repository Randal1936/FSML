import pandas as pd
import numpy as np
import xlwings as xw
from PolicyAnalysis import cptj as cj
import os
import datetime

os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')

# 导入原始数据，方便之后查看疑似重复文本
Sample = pd.read_excel('./调试数据.xlsx', sheet_name='Sheet1')
Sample.set_index(Sample['id'], inplace=True)
Sample.drop(['id'], inplace=True, axis=1)

# 导入原始分词结果（词项越多，鉴别相似性的准确度越高）
df = pd.read_excel('./PanelDataSample.xlsx', sheet_name='overall_DTM_all')
df.set_index(df['id'], inplace=True)
doc_id = df['id']
df.drop(['id'], inplace=True, axis=1)

rank = [x for x in range(0, df.shape[0])]
keymap = [(m, n) for m, n in zip(rank, doc_id)]
# 建立一个字典，保留文档的 id 序号
keymap = dict(keymap)

# DataFrame 转化为 ndarray
matrix = df.__array__()

# 调用 cptj 中计算余弦值的函数，默认余弦值(相似度)大于 0.9 时记录在案
result = cj.cos_rank(matrix, keymap, threshold=0.9)
pairs = result['id']
ids = []
for pair in pairs:
    ids.append(pair[0])
    ids.append(pair[1])
ids = list(pd.Series(ids).unique())

# 将有重复嫌疑的样本提到前面来，得到新的样本
simi_sample = cj.dataframe_filter(Sample, ids, axis=0, status=0)
rest_sample = cj.dataframe_filter(Sample, ids, axis=0, status=1)
new_sample = pd.concat([simi_sample, rest_sample], axis=0)

# 将样本写入 excel 表格，且有重复嫌疑的样本标黄
app1 = xw.App(visible=False, add_book=False)
try:
    wb = app1.books.add()
    sht = wb.sheets.add('Data')
    sht['A1'].value = new_sample
    # 将嫌疑区域标为黄色
    sht.range('A2:F{0}'.format(len(simi_sample)+2)).color = 255,255,0
    time_now = datetime.datetime.today()
    time_now = time_now.strftime('%Y%m%d_%H%M')
    wb.save('NewData_'+time_now+'.xlsx')
finally:
    app1.quit()

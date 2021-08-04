import pandas as pd
import numpy as np
import xlwings as xw
from PolicyAnalysis import cptj as cj
import os

os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')

# 导入原始分词结果（词项越多，鉴别相似性的准确度越高）
df = pd.read_excel('./PanelDataSample.xlsx', sheet_name='overall_DTM_all')




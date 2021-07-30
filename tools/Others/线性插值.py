import numpy as np
from scipy import interpolate
import pylab as pl
import xlwings as xw
import pandas as pd
import os
import statsmodels.api as sm
import statsmodels.formula.api as smf

os.chdir('C:/Users/ThinkPad/Desktop')

app = xw.App(visible=False, add_book=False)
try:
    wb = app.books.open('数据搜集20210420.xlsx')
    sht = wb.sheets('杠杆率（年）')
    df = pd.DataFrame(sht.used_range.value)
finally:
    app.quit()

cols = list(df.loc[0])
cols[0] = 'year'
df.columns = cols
df.drop(0, axis=0, inplace=True)
df.dropna(axis=1, how='all', inplace=True)
df['year'] = df['year'].apply(int)

x = list(df.index)
x_H = list(df.index)
del x_H[0: 6]
x_NFC = list(df.index)
del x_NFC[0: 6]
y_H = list(df['Households and NPISHs'])
y_NFC = list(df['Non-financial corporations'])
y_H.remove(None)
y_NFC.remove(None)

pl.plot(x,y,"ro")

# 这个外推完全只看最近的一个点，有点太不靠谱了
for kind in ["slinear"]:#插值方式
    #"nearest","zero"为阶梯插值
    #slinear 线性插值
    #"quadratic","cubic" 为2阶、3阶B样条曲线插值
    f=interpolate.interp1d(x_H,y_H,kind=kind, fill_value='extrapolate')
    # ‘slinear’, ‘quadratic’ and ‘cubic’ refer to a spline interpolation of first, second or third order)
    ynew=f(x)
    pl.plot(x,ynew,label=str(kind))
pl.legend(loc="lower right")
pl.show()


# 接下来尝试 OLS 外推
# 法一，使用smf + 公式
y = df['Households and NPISHs']
x = df['year']

model = smf.ols('Households~year', data=df, missing='drop').fit()
print(model.summary())

# 法二，使用 sm + np 列表



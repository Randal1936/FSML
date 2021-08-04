from matplotlib import pyplot as plt
import xlwings as xw
import pandas as pd


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

    sht = wb.sheets('>=0次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 0].index]
    sht = wb.sheets('>=1次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 1].index]
    sht = wb.sheets('>=2次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 2].index]
    sht = wb.sheets('>=3次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 3].index]
    sht = wb.sheets('>=4次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 4].index]
    sht = wb.sheets('>=5次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 5].index]
    sht = wb.sheets('>=6次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 6].index]
    sht = wb.sheets('>=7次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 7].index]
    sht = wb.sheets('>=8次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 8].index]
    sht = wb.sheets('>=9次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 9].index]
    sht = wb.sheets('>=10次')
    sht.range('A1').value = df.loc[x[x['Sum'] >= 10].index]
    sht = wb.sheets('>=11次')
    sht.range('A1').value = df.loc[x[x['Sum'] >=11].index]

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

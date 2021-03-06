
## 输出数据<!-- {docsify-ignore} -->


输出数据非常简单，使用 pandas 自带的 to_excel 功能即可

```python
df.to_excel('筛选后数据.xlsx')
```

如果想要一次写入多个 excel 表格，可以操作如下：

```python
with pd.ExcelWriter("excel 样例.xlsx") as writer:
	data.to_excel(writer, sheet_name="这是第一个sheet")
	data.to_excel(writer, sheet_name="这是第二个sheet")
	data.to_excel(writer, sheet_name="这是第三个sheet")
```

如果是用 xlwings 写入后直接保存，只需要：

```python
app = xw.App(visible=False, add_book=False)
wb = app.books.add()
sht = wb.books.add('Data')
sht['A1'].value = df
wb.save('C:/Users/ThinkPad/Desktop/Data.xlsx') # 输入工作簿保存路径
app.quit() # 一定要退出 app
```

> [!WARNING]
> 在输出含有 MultiIndex 的 DataFrame 时，xlwings 会报错，此时需要借助 VBA 的 api 来完成导入，[参考此处](https://stackoverflow.com/questions/38305346/xlwings-vs-pandas-native-export-with-multi-index-dataframes-how-to-reconcile)

```python
filename = 'format_excel_export.xlsx'
s.to_excel(filename)

outpath = os.path.join(os.path.abspath(os.path.dirname(__file__)), filename)
os.path.sep = r'/'
wb = xw.Workbook(outpath)

xw.Range('Sheet1', 'A13').value = s
```

xlwings 项目负责人 2016 年说会考虑改善这里的功能，但是 issue 打开之后再也没有了消息，估计是鸽掉了，因此**这里还是使用 pandas 的导出功能更好**，参考 KnowPolicy Alpha v1.0.py 末尾的数据导出


> [!TIP]
> 如果想查看指标计算过程当中产生的数据，可以在指标计算结束后，先获取对应数据，再按照上述方法导出

**目前可获得的数据：**

- supervisors.class_title # 检索标题得到的监管主体种类数
- supervisors.class_source # 检索来源得到的监管主体种类数
- supervisors.point_title # 检索标题得到的监管主体得分
- supervisors.point_source # 检索来源得到的监管主体得分



- institutions.DTM_class  # 按正文前十句话检索得到的 DTM
- institutions.DTM_final # 按正文前十句话检索得到的被监管机构数



- negative_tone.tone # 相对情感语调和绝对情感语调计算结果
- negative_Tone.words # 正向情感词词频，负向情感词词频，总词数统计结果



- business.DTM_aver  # DTM 1、2、3 被监管业务数求均值
- business.DTM_final  # DTM 1、2、3 被监管业务种类数汇总
- business.DTM1_class  # 按正文检索得到的 Doc-Term Matrix
- business.DTM2_class  # 按前十句话检索的 Doc-Term Matrix
- business.DTM3_class  # 按标题检索得到的 Doc-Term Matrix



- titles.DTM  # 正文检索获得的 Document Term Matrix
- titles.DFC  # 正文检索获得的 Document Frequency Context
- titles.DTM_final  # 标题个数和标题种类数



- numeral.DTM  # 正文检索获得的 Document Term Matrix
- numeral.DFC  # 正文检索获得的 Document Frequency Context
 

如果还有其他想要获得的数据，可以再观察一下代码的工作流，然后对指标计算的代码稍作修改，返回数据的操作涉及到[ Python 的类与继承](Python?id=类与继承)


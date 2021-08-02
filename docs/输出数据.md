
## 输出数据<!-- {docsify-ignore} -->


输出数据非常简单，使用 pandas 自带的 to_excel 功能即可

```python
df.to_excel('筛选后数据.xlsx')
```

如果是用 xlwings 写入后直接保存，只需要：

```python
wb = wb.save('C:/Users/ThinkPad/Desktop/Data.xlsx') # 输入工作簿保存路径
app.quit() # 一定要退出 app
```

如果想查看指标计算过程当中产生的数据，可以在指标计算结束后，先获取对应数据，再按照上述方法导出

目前可获得的数据：

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
 




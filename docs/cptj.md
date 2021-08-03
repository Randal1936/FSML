
### jieba_vectorizer

### make_doc_freq(word, doc)

### make_info_freq(name, pattern, doc)

### make_docs_freq(word, docs)

### make_infos_freq(name, pattern, docs)

### words_docs_freq(words, docs)



### infos_docs_freq(infos, docs)



### top_n_sent(n, doc, percentile=1)

**原理：**找到表示语义终结的标点符号：[!?。]，确定其位置，然后以此为位置索引，以句子为单位将文本拆开
拆解句子有两种方式：
- n ： 设定只要前 n 个句子，
- percentile： 设定只要前多少百分比个句子，比如 11 个句子取前 50%，11 × 50% = 5.5，向下取整，得到 5 个句子

```python
    info = '[。？！]'
    re_iter = list(re.finditer(info, doc))
    # max_iter 是 re 匹配到的最大次数
    max_iter = len(re_iter)
    # 考虑 percentile 的情况，如果总共有11句，就舍弃掉原来的 n，直接改为总句数的 percentile 对应的句子数
    # 注意是向下取整
    if percentile != 1:
        n = math.ceil(percentile * max_iter)
        if n > 0:
            return doc[0: re_iter[n - 1].end()]
        elif n == 0:
            return doc

    if max_iter >= n:
        return doc[0: re_iter[n - 1].end()]
    # 这一句表示，如果匹配到的次数大于0小于10，循环自然结束，也输出结果
    elif 0 < max_iter < n:
        return doc[0: re_iter[-1].end()]
    # 这一句表示，正文过于简短，一个结束符号都没有，直接输出全文
    elif max_iter == 0:
        return doc
    # 为减少重名的可能，尽量在函数体内减少变量的使用
```

> [!NOTE]
> - percentile 默认值为 1，此时以句子个数 n 为准切割内容，但是若 percent 小于 1，则舍弃句子个数，以 percentile 为准
> - 前 50% 的句子并非包括一半字数的内容，而是前 50% 个句子（5句 in 10句）

### dtm_sort_filter(dtm, keymap, name=None)

![dtm_sort_filter](dtm_sort_filter.jpg)



### dtm_point_giver(dtm, keymap, scoremap, name=None)

![dtm_point_giver](dtm_point_giver.jpg)


### dfc_sort_filter(dfc, keymap, name=None)
dfc_sort_filter


### dfc_point_giver(dfc, keymap, name=None)


### dfc_sort_counter(dfc, name=None)


### match_cut(matchobj, cls)


### re_formatter(file_time)


### year_month_finder(file_time)


### month_day_finder(file_time)


### quarter_finder(file_time)


### dataframe_filter(df, keywords, axis)


### list_to_txt(name, content, sep = '\n')


### txt_to_list(path, sep = '\n')


### doc_filter(df, target)


### vec_cos(arr1, arr2)


### cos_rank(matrix, keymap=None)





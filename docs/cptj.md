
### 文本向量化 Python 类：jieba_vectorizer

jieba_vectorizer(df, userdict, stopwords, orient=False)

- df: 输入数据
- userdict: 用户自定义词典
- stopwprds:
- orient:

这个函数是对 jieba 分词流程的简单封装（封装是为了方便调用），只需要输入数据、用户自定义词典和停用词词典，就可以自动完成文本向量化的全过程

在 [数据处理 > 文本向量化](TextVect) 有详细介绍

需要注意的是，本项目每个指标的分词流程彼此独立，各自调用了 jieba_vectorizer，也就是说，本项目的指标都各自独立地做了一次分词，而非统一分词后再计算指标，这样做主要出于三个原因：

- **指标的异质性：**有的指标适宜分词处理(如监管机构、被监管业务)，有的指标不适宜分词处理（如标题、数字）
- **计算过程的独立性：**由于指标较多，而且后续有可能拓展其他指标，因此有必要将各个指标的计算过程彼此分隔，以便于调试和改写
- **用户词典和分词目标相匹配：**用户词典也不能解决所有的问题，过于复杂的用户词典容易引起词与词之间的冲突，比如 "信托" 和 "信托公司"，"证券" 和 "证券公司"，两字强调业务属性，四字强调机构属性，如果他们都在用户词典当中，那么实际的切分结果是相当随机。但实际上，在计算业务种类数时，我们需要前者，在计算机构种类数时，我们需要后者，所以建议分头建立词典并各自分词

### re 词频统计函数(关键词版)


#### make_doc_freq

make_doc_freq(word, doc)


#### make_docs_freq

make_docs_freq(word, docs)


#### words_docs_freq

words_docs_freq(words, docs)


### re 词频统计函数(正则表达式版)

#### make_info_freq

make_info_freq(name, pattern, doc)


#### make_infos_freq

make_infos_freq(name, pattern, docs)


#### infos_docs_freq

infos_docs_freq(infos, docs)




### 计类函数


#### dtm_sort_filter

dtm_sort_filter(dtm, keymap, name=None)

![dtm_sort_filter](dtm_sort_filter.jpg)



#### dfc_sort_filter

dfc_sort_filter(dfc, keymap, name=None)


#### dfc_sort_counter

dfc_sort_counter(dfc, name=None)



### 赋分函数

#### dtm_point_giver

dtm_point_giver(dtm, keymap, scoremap, name=None)

![dtm_point_giver](dtm_point_giver.jpg)


#### dfc_point_giver

dfc_point_giver(dfc, keymap, name=None)



### 日期格式函数

#### re_formatter

re_formatter(file_time)

#### year_month_finder

year_month_finder(file_time)


#### month_day_finder

month_day_finder(file_time)


#### quarter_finder

quarter_finder(file_time)


#### quarter2date

quarter2date(date)


### 其他辅助函数


#### top_n_sent

top_n_sent(n, doc, percentile=1)

**原理：**找到表示语义终结的标点符号：[!?。]，确定其位置，然后以此为位置索引，以句子为单位将文本拆开
拆解句子有两种方式：

- n ： 设定只要前 n 个句子，
- percentile： 设定只要前多少百分比个句子，比如 11 个句子取前 50%，11 × 50% = 5.5，向下取整，得到 5 个句子

```python
    info = '[。？！]'
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
```

> [!NOTE]
> - percentile 默认值为 1，此时以句子个数 n 为准切割内容，但是若 percent 小于 1，则舍弃句子个数，以 percentile 为准
> - 前 50% 的句子并非包括一半字数的内容，而是前 50% 个句子（5句 in 10句）


#### match_cut

match_cut(matchobj, cls)

- matchobj: 匹配好的字符串
- cls: 正则表达式类型

正则表达式匹配到的结果并不能直接使用，原因是我们在首尾规定的字符也被匹配到了，比如，我们想要夹在汉字中间的阿拉伯数字，可以这样编写正则表达式：

```python
[\u4e00-\u9fa5][0-9]+[\u4e00-\u9fa5]'
```

但匹配到的结果是这样的："于30天", ""


#### dataframe_filter

dataframe_filter(df, keywords, axis)


#### list_to_txt

list_to_txt(name, content, sep = '\n')


#### txt_to_list

txt_to_list(path, sep = '\n')


#### doc_filter

doc_filter(df, target)


#### vec_cos

vec_cos(arr1, arr2)


#### cos_rank

cos_rank(matrix, keymap=None)



# 文本向量化<!-- {docsify-ignore} -->

在处理数据时可以使用本项目的支持包 cptj，cptj 集成了很多有用的函数，同时文本向量化的下述四个步骤也都封装在了这个库中，可以一键调用

```python
from PolicyAnalysis import cptj as cj
dtm = cj.jieba_vectorizer(data, userdict, stopwords).DTM
```
输出的结果是 DTM (DTM: Document Term Matrix，词矩阵) ，每行代表一个文件，每列代表一个词项，单元格数值表示该文件中对应关键词出现的频次

![DTM示例](DTM示例.jpg)

如果想获取 sklearn.CountVectorizer 产生的原始矩阵，可以令 matrix = jieba_vectorizer.DTM0

原始矩阵经过两步转化可以得到上述 dtm 的结果

```python
vect = cj.jieba_vectorizer(data, userdict, stopwords)
matrix = vect.DTM0
matrix = matrix.toarray()
features = vect.get_feature_names()
dtm = pd.DataFrame(matrix, index=tf['id'], columns=features)
```

以下为文本向量化的详细步骤，感兴趣的读者可以自行阅读或按自身需求修改

#### 1. 获取关键词清单

> 说明：os.chdir 规定了程序的工作目录在什么地方，文件的查、读、写、存都是在这个路径上进行，而 ./ 就指代这个工作目录，加上后面的部分就组成了完整的路径，这是相对路径的写法，较为简洁
- '/'表示根目录
- './'表示当前目录
- '../'表示上一级目录

```python
# os.getcwd() 可以查看当前的工作目录
os.chdir('E:/ANo.3/FSML/FinancialSupervision/tools')
cj.txt_to_list('./words_list/add_words_dict.txt', sep='\n')
```

> 如果程序不支持相对路径，或者出于其他原因想要把相对路径修改为绝对路径，如'E:/ANo.3/FSML/FinancialSupervision/tools/words_list/add_words_dict.txt'，可以操作如下:

```python
abs_path = os.path.abspath('./words_list/add_words_dict.txt')
cj.txt_to_list(abs_path, sep='\n')
```

#### 2. 文本清洗

本项目中的文本清洗指的是把含有英文字母、阿拉伯数字、汉字、标点符号和其他特殊字符的文本转化为仅含汉字的文本，这一需求可以用[正则表达式](https://www.runoob.com/regexp/regexp-syntax.html)轻松实现

```python
words = []  # 新建一个列表用来存放清洗后的正文
for i, row in tf.iterrows():
    # tf 是上文获取的样本(pandas.DataFrame)
    # 这个循环可以逐行遍历 DataFrame
    try:
        result = row['正文']  # 逐行获取正文，毕竟清洗的就是正文
        rule = re.compile(u'[^\u4e00-\u9fa5]')  # 编写一个正则表达式，方括号中'^'表示'非'，'\u4e00-\u9fa5'表示所有汉字，合起来就表示汉字之外所有的其他字符
        result = rule.sub('', result)  # 使用正则表达式匹配文本中的内容，并将被匹配到的内容替换为空字符串 ''
        words.append(result)  # 清洗后的正文装入列表中
    except TypeError:  
        # 有的样本是空值，会因为无法清洗导致报错，所以要留一手
        print(i) # 报错之后就输出样本所在行，快速锁定问题样本
    continue

ff = pd.DataFrame(words) # 列表转化为 DataFrame
```

#### 3. jieba 分词

```python
h = int(len(words) / 20)  # 一个简易进度条的准备：获取样本长度并等分为20份 （每份为 5% ）

words = []  # 准备一个空列表
t = 0  # 准备计数变量 t
for i, row in ff.iterrows():
    item = row[0]
    result = jieba.cut(item)  # jieba 切词得到一个列表
    word = " ".join(result)  # 使用空格" "将列表所有元素粘连起来，因为向量化要用到的 sklearn.CountVectorizer() 要求词与词之间以空格作为分隔,如 "互联网金融服务" > "互联网 金融 服务"
    words.append(word)  # 把切词后的结果装入列表

    t += 1  # 每切完一个文件，计数变量加一
    if t % h == 0:  # 打印进度条：t 每超过一份的量就打印一个方块，\r 表示每次都是从头打印，end = '' 表示不换行
        print("\r文本处理进度：{0}{1}%  ".format("■" * int(t / h), int(5 * t / h)), end='')
```

本项目后期采用了更方便的进度条，来自 [alive_progress](https://github.com/rsalmei/alive-progress) 程序包

#### 4. 分词向量化

```python
# CountVectorizer() 可以自动完成词频统计，通过 fit_transform 生成文本向量和词袋库
vect = CountVectorizer()
X = vect.fit_transform(words)  # 返回结果是一个压缩后的矩阵
X = X.toarray()  # 可以将矩阵变回我们可以操作的格式（numpy.ndarray）

# 二维 ndarray 看起来和 DataFrame 相似，但性质完全不同，ndarray 没有 index 和 column
features = vect.get_feature_names()  # 获取列标题中编号所对应的词语
XX = pd.DataFrame(X, index=tf['id'], columns=features)  # 将 ndarray 转换为 pandas.DataFrame ，更换 index 和 columns
```

这里使用了 sklearn 程序包里的 CountVectorizer 函数自动完成了分词结果的向量化，也是目前的主要操作方法，优点是简单易行，但是缺点是必须先分词再完成词频统计和向量转化，分词的方式直接和词频统计的结果直接挂钩，容易产生疏漏

因此我们基于 re 从头完成了一组词频统计的函数，放在 [cptj](支持包cptj.md) 支持包当中


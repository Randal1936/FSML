# FSML Toolkits

本文档是一套论文数据处理程序的使用指南，隶属于孙莎老师领导下的金融监管强度研究项目，仅供内部参考，希望对后续参与研究的老师和同学们有所帮助

## 简介

### 功能介绍
- 自动处理原始政策文本
    - 数据导入
    - 文本分词
    - 词频统计
    - 指标计算
- 批量呈现数据结果
    - 按年或者按季度对某一项指标进行加总、平均、计数处理
    - 批量绘制并保存图像

### 项目结构
```text
tools
├─ PolicyAnalysis —— 指标计算工具
│    ├─ Businesses.py
│    ├─ CRITIC.py
│    ├─ Institutions.py
│    ├─ KnowPolicy Alpha v1.0.py
│    ├─ NegativeTone.py
│    ├─ Numerals.py
│    ├─ Supervisors.py
│    ├─ Titles.py
│    └─ cptj.py
├─ RJGraphing —— 绘图工具
│    ├─ Filter.py
│    ├─ Graphing.py
│    ├─ bar_plot.py
│    ├─ main.py
│    ├─ osmkdir.py
│    ├─ simhei.ttf
│    └─ 词云.py
├─ Samples —— 数据样例
├─ words_list —— 关键词清单
│      ├─ BSI.txt
│      ├─ Supervisor.txt
│      ├─ add_words_dict.txt
│      ├─ businesses.txt
│      ├─ institutions.txt
│      ├─ stop_words.txt
│      ├─ 正向情感词词典_加入政策词汇.txt
│      └─ 负向情感词词典_加入政策词汇.txt
|
├─ MachineLearning —— 机器学习模型
│    └─ 文本数据处理+机器学习调用.py
└─ Others
    └─ 线性插值.py
```


## 快速开始

### 环境配置
请确保电脑上已安装 python 3，并完成相关配置
- 推荐环境: Anaconda
- 推荐 Python 版本：3.7 及以上
- 所需 Package:
    - pandas
    - numpy
    - xlwings
    - sklearn
    - mglearn
    - jieba
    - alive_progress

### 基本使用
#### 1.指标计算工具
- 整理关键词清单
- 打开 PolicyAnalysis > KnowPolicy Alpha v1.0.py
- 修改样本文件所在路径
- 修改面板数据保存路径
- 整体运行程序 (Pycharm shortcut: ctrl + shift + F10)


#### 2.批量绘图工具
- 打开 RJGraphing > Graphing.py
- 设置面板数据读取路径
- 设置图形绘制方式
    - 分类字段: index （'Year', 'Quarter'）
    - 汇总变量：column (任意数值变量)
    - 汇总方式: how ('Sum', 'Mean', 'Count')
- 设置图像保存路径
- 整体运行程序








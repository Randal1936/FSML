# FSML Toolkits

本文档是一套论文数据处理程序的使用指南，隶属于孙莎老师领导下的金融监管强度研究项目，仅供内部参考，希望对后续参与研究的老师和同学们有所帮助

## 简介

### 功能介绍
- 对原始政策文本进行处理，包括导入、分词、词频统计、指标计算
- 导入面板数据，绘制简易图像，目前可以按年或者按季度对某一项指标进行加总、平均、计数处理后输出图像


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
python 3

- pandas
- numpy
- xlwings
- sklearn
- mglearn
- jieba
- alive_progress

### 基本使用







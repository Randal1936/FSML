## 导入数据<!-- {docsify-ignore} -->

项目代码统一使用 [xlwings](https://docs.xlwings.org/zh_CN/latest/quickstart.html) 导入和导出数据

##### xlwings 的优点：
- 延展性好
- 更灵活，可实现精细化操作（循环写入，调整单元格格式）

导入数据流程：创建实例 > 打开工作表 > 选择工作簿 > 读取数据 > 调整格式 > 关闭实例

```python
app1 = xw.App(visible=False, add_book=False)  
# 定义一个实例 app1
# visible=False 表示程序操作 excel 表格的过程不可见，想看表格实时变动的话就设置 visible=True
# add_book=False 表示不新建工作簿，如果不设置的话，由于 add_book 默认为 True，哪怕打开已存在的文件，xlwings 还是会新建一个工作簿

try:
    # 为什么要使用 try···finally··· 格式？
    # 因为 app 用完后应该及时关闭，忘记关闭 app 就退出 IDE 可能导致进程堵塞，会导致电脑的 excel 软件异常。一个常见的情况是程序运行报错后，由于疏忽大意就会忘记关闭 app，因此影响软件工作。
    # 而 try···finally··· 专为应付突发情况（文件不存在，名字打错了等），无论是否报错，finally 后的语句一定会执行，这样就可以保证会自动关闭 app。

    wb = app1.books.open("调试数据.xlsx")  # 打开工作簿
    sht = wb.sheets['Sheet1']  # 定位工作表

    df = sht.used_range.value  # 获取 Sheet 内全部数据
    df = pd.DataFrame(df)  # 将获取的数据转换为 pandas.DataFrame
    df.columns = list(df.loc[0])  # 获取数据首行作为变量名
    df.drop(0, axis=0, inplace=True)  # 删除首行(相当于 Stata 的 firstrow)
    df.reset_index(inplace=True, drop=True)  # 
finally:
    app1.quit()  # 关闭实例

tf = df.copy()  # copy 样本的原因是万一后续程序出错，不至于重头再来，尤其适用于逐行运行/调试的时候，建议每一步都先备份，因为有的步骤会相当耗时
```

更简单的数据导入方法是使用 pandas 自带的数据读取函数

##### pandas 函数优点：
- 简单
- 读取速度快

```python
df = pd.read_excel('C:/Users/ThinkPad/Desktop/Data.xlsx', sheet_name='Sheet1')
```

其他的操作 excel 的 package 还包括 [xlrd](https://xlrd.readthedocs.io/en/latest/)、[openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html) 等

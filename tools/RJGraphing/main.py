from RJGraphing.Filter import Branch
import pandas as pd
import numpy as np

df = pd.read_excel('C:/Users/ThinkPad/Desktop/20210705调试_汇总政策强度.xlsx',
                   sheet_name='Sheet24')
df.sort_values(['Year', 'id'], inplace=True)

Target = pd.read_excel('C:/Users/ThinkPad/Desktop/20210705调试_汇总政策强度.xlsx',
                       sheet_name='Target')

Target.set_index('Unnamed: 0', inplace=True, drop=True)
Target.drop('Target Line', inplace=True, axis=1)
Target = np.array(Target)

Sample_pool = [(year, order) for year, order in zip(df['Year'], df['id'])]
Sample_keys = list(df['Year'].unique())
Sample = {}
for key in Sample_keys:
    Sample[key] = []
for item in Sample_pool:
    Sample[item[0]].append(item[1])


jyz = Branch(df, Target, 'Year', 'pmg强度', 25, "C:/Users/ThinkPad/Desktop/GraphFolder/Filtered_data_",  Sample)



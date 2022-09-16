# -*-coding: utf-8 -*-
import pandas as pd

path = 'C:\\Users\\13602\\Desktop\\Wind 数据变换\\'
data = pd.read_excel(path + '工作簿1.xlsx').iloc[:-2, :]
data.set_index(['证券代码', '证券简称'], inplace=True)
group_name = data.columns.to_list()
group_2019y = [i for i in group_name if ('2019年报' in i)]
group_2020y = [i for i in group_name if ('2020年报' in i)]
group_2021h = [i for i in group_name if ('2021中报' in i)]
data_2019y = data.loc[:, group_2019y]
data_2020y = data.loc[:, group_2020y]
data_2021h = data.loc[:, group_2021h]

data.pivot_table()
exit()

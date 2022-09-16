# -*-coding: utf-8 -*-
import os
import shutil
import pandas as pd
path = 'D:\\SynologyDrive\\AY【通用文件】\\26 底稿归档\\3 需补充底稿\\新建文件夹\\'
df = pd.read_excel('D:\\SynologyDrive\\AY【通用文件】\\26 底稿归档\\3 需补充底稿\\工作簿2.xlsx')
path2 = 'D:\\按项目公司名称分类\\'
print(df)

# path = r"D:\SynologyDrive\公司债（含可交换债）\第一部分 公司债券承销业务尽职调查文件\第一章节 发行人基本情况调查\1-3 发行人对其他企业的重要权益投资\1-3-1 对发行人有重要影响的子公司\1-3-1-4 诚信信息查询文件\3 深圳锦弘劭晖投资有限公司"
# path = r"F:\1 上海大陆期货有限公司"
# path = r'D:\SynologyDrive\项目管理\2021.08.11 公司债 - 中山证券私募债\2021.09.01 盖章扫描版'
# path = r'D:\SynologyDrive\项目管理\2021.08.11 公司债 - 中山证券私募债\version\pdf'
n=0
file_name = os.listdir(path)
for i,j in zip(df['行标签'],df['计数项:项目公司名称']):
    for root, dirs, files in os.walk(path):
        for file in files:
            src_file = os.path.join(root, file)
            if i in src_file:
                sub_name = src_file.split('\\')
                print(sub_name)
                print(src_file)
                target_path = path2+j+'\\'+'{}-{}-'.format(sub_name[-3],sub_name[-2])+file
                try:
                    os.makedirs(path2+j)
                except:
                    pass
                shutil.copy(src_file, target_path)
                n+=1
                # print(src_file)
                # print(target_path)
        # shutil.copy(path+i+'\\', path+j+'\\')

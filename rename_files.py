# -*-coding: utf-8 -*-
import os
import shutil
import pandas as pd
path = "C:\\Users\\13602\\Documents\\大学资料\\post graduate\\实习\\中山证券\\20200730\\奥创二期贷款移交表\\"
path1 = "C:\\Users\\13602\\Documents\\大学资料\\post graduate\\实习\\中山证券\\20200730\\贷款移交表\\"
# 重命名文件
# path = 'D:\\SynologyDrive\\ZS01a【原稿】中山私募债\\6 销售与发行 - 21中山02\\5-1 发行前准备的文件\\2 To 资本市场部\\2021.12.08 - 21中山02缴款通知书\\缴款通知书\\'
# files = os.listdir(path)
# df = pd.read_excel("D:\\SynologyDrive\\ZS01a【原稿】中山私募债\\6 销售与发行 - 21中山02\\5-1 发行前准备的文件\\2 To 资本市场部\\2021.12.08 - 21中山02缴款通知书\\副本要素表.xlsx")
# for i,j,z in zip(df['机构简称'][-6:], df['申购数量（亿元）'][-6:], files[-6:]):
#        os.rename(path+z,path+"缴款通知书-{}-{}亿元.pdf".format(i, j))
# exit()
# 添加后缀
path = 'D:\\SynologyDrive\\ZS01a【归档】中山私募债\\2021.11.15 发行稿-底稿\\发行稿底稿盖章版\\'
files = os.listdir(path)
for f in files:
    new_name = f[:-9]+'【0630】'+'.pdf'
    os.rename(path+f, path+new_name)
exit()
# 读取文件夹名称，重新命名另一个文件夹
# file_name_path = 'D:\\SynologyDrive\\ZS01a【归档】中山私募债\\2021.11.15 发行稿-底稿\\21中山01发行阶段底稿-PDF\\新建文件夹\\'
# needed_chg_path = 'D:\\SynologyDrive\\ZS01a【归档】中山私募债\\2021.11.15 发行稿-底稿\\发行稿底稿盖章版\\新建文件夹\\'
# file_name = os.listdir(file_name_path)
# needed_chg_file_name = os.listdir(needed_chg_path)
# for i,j in zip(file_name, needed_chg_file_name):
#     os.rename(needed_chg_path+j, needed_chg_path+i)

# for f in pre_name:
#     child_name = os.listdir(path1_1+f)
#     for j in child_name:
#         subchild_name = os.listdir(path1_1+f+'\\'+j)
#         if len(subchild_name) == 0:
#             pass
#         else:
#             z_split = subchild_name[0].split('-')
#         if len(z_split) == 1:
#             last_name = f
#             first_name = z_split[0][:-4]
#             full_name = first_name+'-'+last_name+'.pdf'
#         if len(z_split) >1:
#             first_name = z_split[0]
#             mid_name = f
#             last_name = z_split[1]
#             full_name = z_split[0]+'-'+f+'-'+z_split[1]
#         os.rename(path1_1+f+'\\'+j+'\\'+subchild_name[0], path1_1+f+'\\'+j+'\\'+full_name)

        # for z in subchild_name:
        #     z_split = z.split('-')
        #     if len(z_split) == 3:
        #         first_name = z_split[0]+'-'+z_split[1]
        #     else:
        #         first_name = z_split[0]
        #     last_name = f.split('-')
        #     last_name = '-'.join(last_name[:-1])+z_split[-1]
        #     os.rename(path1_1 +f+'\\'+j+'\\'+z, path1_1 +f+'\\'+j+'\\'+ first_name+'-'+last_name)

exit()
# print(pre_name)
s = '按揭受理凭证'
loc = ['28座-一单元-401', '', '24座-一单元-302', '26座-三单元-101',
       '24座-二单元-101', '13座-一单元-4101', '14-b座一单元-2904',
       '14-b座一单元-2504', '']
loc = ['27座-2单元-401', '26座-3单元-402',
       '14座-2单元-102', '21座-2单元-401',
       '26座-1单元-301', '22座-2单元-201',
       '24座-3单元-502']

# new_name = [j + i for i, j in zip(pre_name, loc)]
# print(new_name)
#
# new_names = [s + i for i in new_name]
# for i, j in zip(pre_name, new_names):
#     os.rename(path + i, path + j)
# print(os.listdir(path))

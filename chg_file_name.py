# -*-coding: utf-8 -*-
import os
import shutil
import pandas as pd
path = 'F:\\中山证券\\Work\\ACYS\\土地核查2\\2 土地核查\\'
file_path = 'F:\\中山证券\\Work\\ACYS\\土地核查2\\2 土地核查\\网络核查\\'

file_name = os.listdir(file_path)
df = pd.read_excel(path+'土地核查.xlsx')
for i,j in zip(df['序号'],df['开发单位名称']):
    for root, dirs, files in os.walk(file_path):
        for file in files:
            src_file = os.path.join(root, file)
            if j in src_file:
                sub_name = src_file.split('\\')
                print(sub_name)
                print(src_file)
                target_path = path + '土地核查' + '\\' + '{} {}'.format(i, j)
                try:
                    os.makedirs(target_path)
                except:
                    pass
                shutil.copy(src_file, target_path)

exit()
# df2 = pd.read_excel('C:\\Users\\13602\\Desktop\\奥创ABS二期土核-CYK\\土地核查项目-奥创二期-CYK.xlsx',sheet_name='李康')
for i,j in zip(df['序号'],df['开发单位名称']):
    try:
        os.makedirs(os.path.join(file_path, '{} {}\\'.format(i, j)))
        shutil.move(os.path.join(file_path,'{} {}.docx'.format(i,j)), os.path.join(file_path,'{} {}\\{} {}.docx'.format(i,j,i,j)))
    except:
        pass
exit()
print(df.head())
for i,j in zip(df['序号'],df['开发单位名称']):
    try:
        # os.chdir(os.path.join(file_path, '{} {}'.format(i, j)))
        # os.rename(os.path.join(file_path, j+'.docx'), os.path.join(file_path,'{} {}.docx'.format(i,j)))
        try:
            os.makedirs(os.path.join(file_path,'{} {}\\'.format(i,j)))
            print(os.path.join(file_path,'{} {}\\'.format(i,j)))
        except FileExistsError:
            exit()
        shutil.move(os.path.join(file_path,'{} {}.docx'.format(i,j)), os.path.join(file_path,'{} {}\\{} {}.docx'.format(i,j,i,j)))
    except FileNotFoundError:
        pass
"""删除文件名的数字"""
# file_name = os.listdir(file_path)
# for i in file_name:
#     old = i
#     for j in range(0,10):
#         if str(j) in i:
#             i = i.replace(str(j),'')
#     try:
#         os.rename(os.path.join(file_path,old),os.path.join(file_path,i))
#     except FileExistsError:
#         pass
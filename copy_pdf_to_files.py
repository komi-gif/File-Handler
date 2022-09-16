# -*-coding: utf-8 -*-
import os
import shutil
target_file_path ="D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字【终稿】\\pdf_modified【终稿】\\"
save_palce_path = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\2 尽职调查阶段\\"
target_file = os.listdir(target_file_path)
for i in target_file:
    print(i)
print("开始复制")
save_place = [i[0] for i in os.walk(save_palce_path)]
for file_name in target_file:
    sign = file_name.split(' ')[0]
    for file_path in save_place:
        if sign in file_path:
            print(file_name)
            shutil.copy(target_file_path+file_name, file_path+'\\'+file_name)
            break


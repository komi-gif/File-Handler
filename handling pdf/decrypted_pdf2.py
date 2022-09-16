# -*-coding: utf-8 -*-
import pikepdf
import os

path = "F:\\Study\\1 Exams\\保代考试\\4 考试资料\\2022年\\233网校讲义\\"
files = os.listdir(path)
for name in files:
    pdf = pikepdf.open(path+name,password='')
    pdf.save(path+name[:-4]+"【可编辑】.pdf")
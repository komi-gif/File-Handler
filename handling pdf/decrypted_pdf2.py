# -*-coding: utf-8 -*-
import pikepdf
import os

path = "add your folder path"
files = os.listdir(path)
for name in files:
    pdf = pikepdf.open(path+name,password='')
    pdf.save(path+name[:-4]+"【可编辑】.pdf")

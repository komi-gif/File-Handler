# -*-coding: utf-8 -*-

import PyPDF2
import os

# 新建save_path 文件夹
file_path = "add your pdf folder"
save_path = "add your pdf folder"

file_name = os.listdir(file_path)

for name in file_name:

    file1 = open(file_path+name, 'rb')
    reader1 = PyPDF2.PdfFileReader(file1)
    writer = PyPDF2.PdfFileWriter()
    for pageIndex in range(reader1.numPages-1):
        writer.addPage(reader1.getPage(pageIndex))
    newFile = open(save_path+name, "wb")
    writer.write(newFile)
    newFile.close()
    file1.close()


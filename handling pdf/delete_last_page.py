# -*-coding: utf-8 -*-

import PyPDF2
import os

# 新建save_path 文件夹
file_path = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字文件\\3 PDF_Modified\\"
save_path = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字文件\\4 PDF_for_Print\\"
# newFileName = "/Users/weiyang/Desktop/NewTest.pdf"
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


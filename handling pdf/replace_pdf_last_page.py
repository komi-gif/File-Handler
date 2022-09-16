# -*-coding: utf-8 -*-
import PyPDF2
import os
file_path1 = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字文件\\3 PDF_Modified - 副本\\"
fileName2 = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字【终稿】\\项目组签字页v2.pdf"
# 新建save_path 文件夹
save_path = "D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字【终稿】\\pdf_modified【终稿】\\"
# newFileName = "/Users/weiyang/Desktop/NewTest.pdf"
file_name = os.listdir(file_path1)
file_2 = open(fileName2, 'rb')
reader2 = PyPDF2.PdfFileReader(file_2)
print(reader2.numPages)
page_num = reader2.numPages
for name, i in zip(file_name,range(reader2.numPages)):
    print(file_path1+name)
    file1 = open(file_path1+name, 'rb')
    reader1 = PyPDF2.PdfFileReader(file1)
    writer = PyPDF2.PdfFileWriter()
    for pageIndex in range(reader1.numPages-1):
        writer.addPage(reader1.getPage(pageIndex))

    writer.addPage(reader2.getPage(i))

    newFile = open(save_path+name, "wb")
    writer.write(newFile)
    newFile.close()
    file1.close()
file_2.close()

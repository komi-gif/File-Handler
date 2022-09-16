# -*-coding: utf-8 -*-
import glob
import win32com.client
import os

word = win32com.client.Dispatch("Word.Application")
print('my check point 1')

pdfs_path = "D:\\SynologyDrive\\HYS01a【原稿】火焰山股\\1 项目管理\\3 资料整理\\1 客户提供的资料\\2022.08.22 初步尽调资料\\3 8.20补充资料\\"  # folder where the .pdf files are

# stored
reqs_path = pdfs_path

for i, doc in enumerate(glob.iglob(pdfs_path + "*.pdf")):
    print('my check point 2')
    print(doc)
    filename = doc.split('\\')[-1]
    in_file = os.path.abspath(doc)
    print(in_file)
    wb = word.Documents.Open(in_file)
    out_file = os.path.abspath(reqs_path + filename[0:-4] + ".docx".format(i))
    print("outfile\n", out_file)
    wb.SaveAs2(out_file, FileFormat=16)  # file format for docx
    print("success...")
    wb.Close()

word.Quit()

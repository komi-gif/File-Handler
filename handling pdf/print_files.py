# -*-coding: utf-8 -*-
import tempfile
import win32api
import win32print
import os

path = 'D:\\SynologyDrive\\FY01a【归档】物业尾款一期\\【纸质归档】材料整理\\项目组签字文件\\4 PDF_for_Print\\'
filenames = os.listdir(path)
for name in filenames:
    open(path+name,"r")
    win32api.ShellExecute (
      0,
      "print",
      path+name,
      #
      # If this is None, the default printer will
      # be used anyway.
      #
      '/d:"%s"' % win32print.GetDefaultPrinter (),
      ".",
      0
    )
    print(name)


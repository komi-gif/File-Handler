# -*-coding: utf-8 -*-
import tempfile
import win32api
import win32print
import os

path = 'add your folder path'
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


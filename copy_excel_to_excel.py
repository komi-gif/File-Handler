# -*-coding: utf-8 -*-
import os
from win32com.client import Dispatch

path = "D:\\SynologyDrive\\风险处置\\2 奥园\\9-2 邮件及回复 to 原始权益人\\发出邮件\\2022.09.01 【请确认】奥园集团未偿还债务\\底稿-2022.09.01\\"
save_path = "D:\\SynologyDrive\\风险处置\\2 奥园\\9-2 邮件及回复 to 原始权益人\\发出邮件\\2022.09.01 【请确认】奥园集团未偿还债务\\"
excel = Dispatch('Excel.Application')
files = [
    "司法案件.xls",
    "被执行人.xls",
    "裁判文书.xls",
    "开庭公告.xls",
    "股权冻结.xls",
    "法院公告.xls",
    "立案信息.xls",
    "送达公告.xls",

]
excel.Visible = False
excel.DisplayAlerts = False
wb_new = excel.Workbooks.Add()
for i in files:
    filename = path + i
    wb = excel.Workbooks.Open(filename)
    wb.Sheets(i[:-4]).Copy()
    wb.Worksheets(i[:-4]).Move(Before=wb_new.Worksheets("Sheet1"))

wb_new.SaveAs(save_path + "奥园集团司法案件、被执行人、裁判文书等情况【企查查】-2022.0.30.xlsx")
wb_new.Close()

# -*-coding: utf-8 -*-
from win32com.client import Dispatch
xlapp = Dispatch('Excel.Application')
xlapp.Visible = True
excel_path = r'D:\SynologyDrive\FY01a【归档】物业尾款一期\2 尽职调查阶段\2-6 基础资产基本情况调查\2-6-1 基础资产一般要求\2-6-1-3 基础资产现金流状况\2-6-1-3-7 ' \
             r'基础资产现金流压力测试参数、依据及合理性分析\\基础资产物业管理费现金流预测 - FY.xlsx '

'''
    1、将所有列调整为一页
    2、打印表格标题
    3、设置页眉为sheet name， 设置页脚为“第？页，共？页”字样
    4、设置表格为水平居中
'''
xlbook = xlapp.Workbooks.Open(excel_path, UpdateLinks=False, ReadOnly=False)
sheet_names = [sht.Name for sht in xlbook.Worksheets]
for name in sheet_names:
    ws = xlbook.Worksheets(name)
    ws.Activate()
    ws.PageSetup.PrintTitleRows = "$1:$1"
    ws.PageSetup.CenterFooter = "第 &P 页，共 &N 页"
    ws.PageSetup.CenterHeader = "&A"
    ws.PageSetup.Zoom = False
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.CenterHorizontally = 1
    ws.PageSetup.Orientation = 2  # 页面横向

xlbook.Save()
xlbook.Close()
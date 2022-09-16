import os
import docx
import pandas as pd
from docx.document import Document
from win32com.client import Dispatch
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import numpy as np
import time


def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document order.
    Each returned value is an instance of either Table or Paragraph. *parent*
    would most commonly be a reference to a main Document object, but
    also works for a _Cell object, which itself can contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def read_table(table):
    return [[cell.text for cell in row.cells] for row in table.rows]


def read_word(word_path):
    content = []
    doc = docx.Document(word_path)
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            # print("text", [block.text])
            content.append(["text", block.text])
        elif isinstance(block, Table):
            # print("table", read_table(block))
            content.append(["table", read_table(block)])
    return content


# def get_target_array(my_path):
#     files = os.listdir(my_path)
#     for name in files:
#         my_content = read_word(my_path + name)
#         my_array = [item[1] for item in my_content]
#         return files, my_content, my_array


def get_target_table(my_path, start_title, end_title):
    target = []
    files = os.listdir(my_path)
    for name in files:
        my_content = read_word(my_path + name)
        my_array = [item[1] for item in my_content]

        position = [i for i, s in enumerate(my_array) if start_title in s][1]
        position2 = [i for i, s in enumerate(my_array) if end_title in s][1]

        if position2 - position > 1:
            my_list = my_content[position:position2]
            for j in my_list:
                if j[0] == 'table':
                    my_table = j[1:]
                    my_table = sum(my_table, [])
                    # my_table = list(map(list, zip(*my_table)))  # 列表转置
                    my_table.insert(0, [name.split('-')[0]] * len(my_table))  # 每一行前面加入项目公司名称
                    target.append(my_table)
    if len(target) > 1:
        my_array = np.array(target, dtype='object')
        my_array = my_array.reshape(len(my_array), 1)
        my_df = pd.DataFrame()
        for lis in my_array:
            df_lis = pd.DataFrame(lis[0][1:])
            df_lis.insert(0, '项目公司名称', lis[0][0])
            my_df = my_df.append(df_lis, ignore_index=True)

    elif len(target) == 1:
        my_df = pd.DataFrame(target[0][1:])
        my_df.insert(0, '项目公司名称', target[0][0])
    else:
        my_df = pd.DataFrame()

    return my_df


if __name__ == '__main__':
    T1 = time.time()
    # print("文件路径示例：C:/Users/Desktop/项目公司/")
    # path = input("请输入企查查专业报告WORD版文件夹路径：")
    # save_path = input("请输入导出的文件Excel保存路径：")
    path = "D:/SynologyDrive/AYRK01【奥园风险处置】/7 舆情跟踪/项目公司是否涉及诉讼、仲裁或其他纠纷、行政处罚的核查/ACYS/2022.09.01/企查查报告/"  # 项目根目录
    save_path = 'D:/SynologyDrive/AYRK01【奥园风险处置】/7 舆情跟踪/项目公司是否涉及诉讼、仲裁或其他纠纷、行政处罚的核查/ACYS/2022.09.01/' + '项目公司网络核查.xlsx'
    sifa = get_target_table(path, '4.1司法案件', '4.2被执行人信息')
    print("1、已处理完成司法案件信息")
    zhixing = get_target_table(path, '4.2被执行人信息', '4.3失信被执行人')
    print("2、已处理完成被执行人信息")
    shixing = get_target_table(path, '4.3失信被执行人', '4.4限制高消费')
    print("3、已处理完成失信被执行人信息")
    caipan = get_target_table(path, '4.8裁判文书', '4.9法院公告')
    print("4、已处理完成裁判文书信息")
    kaiting = get_target_table(path, '4.10开庭公告', '4.11送达公告')
    print("5、已处理完成开庭公告信息")
    jingying = get_target_table(path, '5.1经营异常', '5.2严重违法')
    print("6、已处理完成经营异常信息")
    xingzheng = get_target_table(path, '5.4行政处罚', '5.5环保处罚')
    print("7、已处理完成行政处罚信息")
    writer = pd.ExcelWriter(save_path)
    sifa.to_excel(writer, '4.1司法风险', encoding='utf-8', index=False, header=False)
    zhixing.to_excel(writer, '4.2被执行人信息', encoding='utf-8', index=False, header=False)
    shixing.to_excel(writer, '4.3失信被执行人', encoding='utf-8', index=False, header=False)
    caipan.to_excel(writer, '4.8裁判文书', encoding='utf-8', index=False, header=False)
    kaiting.to_excel(writer, '4.10开庭公告', encoding='utf-8', index=False, header=False)
    jingying.to_excel(writer, '5.1经营异常', encoding='utf-8', index=False, header=False)
    xingzheng.to_excel(writer, '5.4行政处罚', encoding='utf-8', index=False, header=False)
    writer.save()
    writer.close()
    writer.handles = None
# 修改格式
    xlapp = Dispatch("Excel.Application")
    xlapp.visible = True
    xlbook = xlapp.Workbooks.Open(save_path, UpdateLinks=False, ReadOnly=False)
    sheet_names = [sht.Name for sht in xlbook.Worksheets]
    for name in sheet_names:
        ws = xlbook.Worksheets(name)
        ws.Activate()
        ws.Cells.Font.Name = 'Times New Roman'
        ws.Cells.Font.Size = 9
        ws.Cells.VerticalAlignment = 2
        ws.UsedRange.Borders.LineStyle = 1
        ws.Columns.WrapText = True
        ws.UsedRange.ColumnWidth = 25
        ws.Cells.EntireRow.AutoFit()
        try:
            ws.Range("1:1").AutoFilter(2, "序号")
            ws.UsedRange.Interior.ColorIndex = 15
            ws.UsedRange.Cells.Font.Bold = True
            ws.AutoFilterMode = False
        except:
            pass
        xlapp.Range("B2").Select()
        xlapp.ActiveWindow.FreezePanes = True
    xlbook.Save()
    xlbook.Close()
    T2 = time.time()
    print('程序运行时间:%fs秒' % ((T2 - T1)))
"""
# 数据->分列->完成
    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("J3").Select
"""
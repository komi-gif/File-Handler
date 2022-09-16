# -*-coding: utf-8 -*-
import glob
from PyPDF2 import PdfFileMerger
from pdfrw import PdfReader
import fitz
import ftfy
import os
from win32com.client import Dispatch, constants, gencache, DispatchEx


class PDFConverter:
    def __init__(self, pathname, export='.'):
        self._handle_postfix = ['doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx', 'png', 'jpg']
        self._filename_list = list()
        self._export_folder = os.path.join(os.path.abspath('..'), '../pdfconver')
        if not os.path.exists(self._export_folder):
            os.mkdir(self._export_folder)
        self._enumerate_filename(pathname)

    def _enumerate_filename(self, pathname):
        '''
        读取所有文件名
        '''
        full_pathname = os.path.abspath(pathname)
        if os.path.isfile(full_pathname):
            if self._is_legal_postfix(full_pathname):
                self._filename_list.append(full_pathname)
            else:
                raise TypeError('文件 {} 后缀名不合法！仅支持如下文件类型：{}。'.format(pathname, '、'.join(self._handle_postfix)))
        elif os.path.isdir(full_pathname):
            for relpath, _, files in os.walk(full_pathname):
                for name in files:
                    filename = os.path.join(full_pathname, relpath, name)
                    if self._is_legal_postfix(filename):
                        self._filename_list.append(os.path.join(filename))
        else:
            raise TypeError('文件/文件夹 {} 不存在或不合法！'.format(pathname))

    def _is_legal_postfix(self, filename):
        return filename.split('.')[-1].lower() in self._handle_postfix and not os.path.basename(filename).startswith(
            '~')

    def run_conver(self):
        '''
        进行批量处理，根据后缀名调用函数执行转换
        '''
        print('需要转换的文件数：', len(self._filename_list))
        for filename in self._filename_list:
            postfix = filename.split('.')[-1].lower()
            funcCall = getattr(self, postfix)
            print('原文件：', filename)
            funcCall(filename)
        print('转换完成！')

    def doc(self, filename):
        '''
        doc 和 docx 文件转换
        '''
        name = os.path.basename(filename).split('.doc')[-1] + '.pdf'
        exportfile = filename.split('.doc')[0]+'.pdf'
        print('保存 PDF 文件：', exportfile)
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        w = Dispatch("Word.Application")
        doc = w.Documents.Open(filename)
        doc.ExportAsFixedFormat(exportfile, constants.wdExportFormatPDF,
                                Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)

        # w.Quit(constants.wdDoNotSaveChanges)

    def docx(self, filename):
        self.doc(filename)

    def xls(self, filename):
        '''
        xls 和 xlsx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = filename.split('.xls')[0]+'.pdf'
        xlApp = DispatchEx("Excel.Application")
        xlApp.Visible = False
        xlApp.DisplayAlerts = 0
        books = xlApp.Workbooks.Open(filename, False)
        for ws in books.Worksheets:
            ws.PageSetup.Zoom = False
            ws.PageSetup.FitToPagesWide = 1
        books.ActiveSheet.ExportAsFixedFormat(0, exportfile)
        books.Close(False)
        print('保存 PDF 文件：', exportfile)
        xlApp.Quit()

    def xlsx(self, filename):
        self.xls(filename)

    def ppt(self, filename):
        '''
        ppt 和 pptx 文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = filename.split('.ppt')[0]+'.pdf'
        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)
        p = Dispatch("PowerPoint.Application")
        ppt = p.Presentations.Open(filename, False, False, False)
        ppt.ExportAsFixedFormat(exportfile, 2, PrintRange=None)
        print('保存 PDF 文件：', exportfile)
        p.Quit()

    def pptx(self, filename):
        self.ppt(filename)

    def png(self, filename):
        '''
        png 和 jpg文件转换
        '''
        name = os.path.basename(filename).split('.')[0] + '.pdf'
        exportfile = os.path.join(self._export_folder, name)
        doc = fitz.open()
        # for img in sorted(glob.glob(os.path.join(filename, "*.png"))):  # 读取图片，确保按文件名排序
        imgdoc = fitz.open(filename)  # 打开图片
        pdfbytes = imgdoc.convertToPDF()  # 使用图片创建单页的 PDF
        imgpdf = fitz.open("pdf", pdfbytes)
        doc.insertPDF(imgpdf)  # 将当前页插入文档
        doc.save(filename.split('.')[0]+'.pdf')  # 保存pdf文件
        doc.close()

    def jpg(self, filename):
        self.png(filename)


if __name__ == "__main__":
    # 支持文件夹批量导入
    pathname = 'D:\\SynologyDrive\\ZS01a【归档】中山私募债\\第1部分 公司债券承销业务尽职调查文件\\第1章节 发行人基本情况调查\\1-1 历史沿革与股权结构\\1-1-1 发行人工商登记文件\\'


    # 也支持单个文件的转换
    # pathname = 'test.doc'

    pdfConverter = PDFConverter(pathname)
    pdfConverter.run_conver()
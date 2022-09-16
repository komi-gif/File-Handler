# -*-coding: utf-8 -*-

from PyPDF4 import PdfFileReader, PdfFileWriter, PdfFileMerger
import fitz
import os
from win32com.client import Dispatch
from win32com.client import constants
from win32com.client import gencache
from win32com.client import DispatchEx
import win32print


"""转换为PDF"""

printers = win32print.EnumPrinters(2)
print(printers)
# (8388608, 'Microsoft Print to PDF,Microsoft Print To PDF,', 'Microsoft Print to PDF', '')
my_printer = 'Microsoft Print to PDF'
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

    # def pdf(self, filename):
    #     pdf_reader = PdfFileReader(filename, 'rb')
    #     if pdf_reader.isEncrypted:
    #         print(filename)
    #         win32api.ShellExecute(0, "print", filename, my_printer, ".", 0)
    #     # if filename in locals():
    #     #     filename.close()


    def doc(self, filename):
        '''
        doc 和 docx 文件转换
        '''
        name = os.path.basename(filename).split('.doc')[-1] + '.pdf'
        exportfile = filename.split('.doc')[0] + '.pdf'
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
        exportfile = filename.split('.xls')[0] + '.pdf'
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
        exportfile = filename.split('.ppt')[0] + '.pdf'
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
        doc.save(filename.split('.')[0] + '.pdf')  # 保存pdf文件
        doc.close()

    def jpg(self, filename):
        self.png(filename)


"""合并PDF"""


def get_reader(filename, password):
    try:
        old_file = open(filename, 'rb')
        print('run  jiemi1')
    except Exception as err:
        print('文件打开失败！' + str(err))
        return None

    # 创建读实例
    pdf_reader = PdfFileReader(old_file, 'rb')

    # 解密操作
    if pdf_reader.isEncrypted:
        if password is None:
            print('%s文件被加密，需要密码！' % filename)
            return None
        else:
            if pdf_reader.decrypt(password) != 1:
                print('%s密码不正确！' % filename)
                return None
    if old_file in locals():
        old_file.close()
    return pdf_reader


def decrypt_pdf(filename, password, decrypted_filename=None):
    """
    将加密的文件及逆行解密，并生成一个无需密码pdf文件
    :param filename: 原先加密的pdf文件
    :param password: 对应的密码
    :param decrypted_filename: 解密之后的文件名
    :return:
    """
    # 生成一个Reader和Writer
    print('run  jiemi')
    pdf_reader = get_reader(filename, password)
    if pdf_reader is None:
        return
    if not pdf_reader.isEncrypted:
        print('文件没有被加密，无需操作！')
        return
    pdf_writer = PdfFileWriter()

    pdf_writer.appendPagesFromReader(pdf_reader)

    if decrypted_filename is None:
        decrypted_filename = "".join(filename[:-4]) + '_' + 'decrypted' + '.pdf'

    # 写入新文件
    pdf_writer.write(open(decrypted_filename, 'wb'))


# 获取pdf页数的函数
def getPdfPages(filePath):
    try:
        reader = PdfFileReader(filePath, strict=False)
    except:
        print(filePath)
        pass

    pageNum = reader.getNumPages()
    return pageNum


# 存储所有pdf文件路径的函数
def loadAllFilesPath(rootPath, filePaths):
    # 分别代表根目录、文件夹、文件
    for root, dirs, files in os.walk(rootPath):
        # 遍历文件
        for file in files:
            # 识别pdf文件
            if file.endswith('.pdf') or file.endswith('.PDF'):
                # 获取文件绝对路径
                filePath = os.path.join(root, file)
                # 文件路径添加进列表
                filePaths.append(filePath)


while True:
    print('请详细阅读下列说明：\n1、文件路径示例：D:\\SynologyDrive\\  \n2、如果是加密的PDF，需使用“Microsoft Print to PDF”生成非加密文档，并删除原始文件\n'
          '3、适用文件格式：pdf, doc, docx, ppt, pptx, xls, xlsx, png, jpg\n4、代码逻辑：将所有非pdf文件转化并保存为pdf，再按顺序合并所有pdf文件\n'
          '5、请在备份文件夹操作,否则原始文件夹会被修改\n6、生成的PDF文件请在“合并完成的文件”查看，建议使用pdf编辑器添加页码\n7、仅Windows适用')
    # 输入根目录的文件夹
    # rootPath = 'C:\\Users\\13602\\Desktop\\土地核查\\201-269\\'
    rootPath = input('请输入需要合并的文件夹路径：')

    pathname = rootPath
    pdfConverter = PDFConverter(pathname)
    pdfConverter.run_conver()

    # 存储所有pdf文件路径的列表
    pdfPaths = []
    # 存储所有pdf文件的页码和索引
    pageCount = 0
    pageIndex = []
    # 找到所有pdf文件路径，并存储所有pdf文件路径到列表中
    loadAllFilesPath(rootPath, pdfPaths)

    # 创建pdf合成器对象
    fileMerger = PdfFileMerger(strict=False)
    # 合并pdf文件
    n = 0
    for pdfPath in pdfPaths:
        print(pdfPath)
        try:
            fileMerger.append(pdfPath)
        except FileNotFoundError:
            print("文件已损坏：", pdfPath)
            pass
        page = getPdfPages(pdfPath)
        pageIndex.append('页码索引：' + str(pageCount + 1) + '~' + str(pageCount + page))
        pageCount += page
        n += 1

    # 创建保存新文件的文件夹
    newDirPath = os.path.join(rootPath, '合并完成的文件')
    if not os.path.exists(newDirPath):
        os.mkdir(newDirPath)
    # 保存合并完成的pdf文件
    fileMerger.write(os.path.join(newDirPath, '合并完成的pdf文件.pdf'))
    # 保存合并pdf文件的顺序
    with open(os.path.join(newDirPath, '合并pdf文件的顺序.txt'), 'w', encoding='utf-8') as f:
        for pdfPath, pdfInd in zip(pdfPaths, pageIndex):
            f.write(pdfPath + '\t\t' + pdfInd + '\n')

    # pdf文件合并成功！
    print(n, '个pdf文件合并成功，请查看文件夹“合并完成的文件”')

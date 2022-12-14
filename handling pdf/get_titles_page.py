# -*-coding: utf-8 -*-

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
import pandas as pd


def parsePDFtoTXT(pdf_path):
    fp = open(pdf_path, 'rb')
    parser = PDFParser(fp)
    document = PDFDocument(parser)
    rsrcmgr = PDFResourceManager()
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    for i, page in enumerate(PDFPage.create_pages(document)):
        interpreter.process_page(page)
        layout = device.get_result()
        print(layout)
        output = str(layout)
        for x in layout:
            if (isinstance(x, LTTextBoxHorizontal)):
                text = x.get_text()
                output += text
    with open(path+'pdfoutput.txt', 'a', encoding='utf-8') as f:
        f.write(output)


def get_word_page(word_list):
    f = open(path+'pdfoutput.txt', encoding='utf-8')
    text_list = f.read().split(' ')
    n = len(text_list)
    for w in word_list:
        page_list = []
    for i in range(1, n):
        if w in text_list[i]:
            page_list.append(i)
        with open(path+'pdfoutput.txt', 'a', encoding='utf-8') as f:
            f.write(w + str(page_list) + '\n')


if __name__ == '__main__':
    path = 'add your file path'
    parsePDFtoTXT(path + 'add your pdf name')
    df = pd.read_excel(path+'content.xlsx')
    words = df['title'].tolist()
    get_word_page(words)

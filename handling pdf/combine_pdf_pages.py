# -*-coding: utf-8 -*-
from PyPDF3 import PdfFileWriter, PdfFileReader
from PyPDF3.pdf import PageObject

pdf_filenames = []

input1 = PdfFileReader(open(pdf_filenames[0], "rb"), strict=False)
input2 = PdfFileReader(open(pdf_filenames[1], "rb"), strict=False)

page1 = input1.getPage(0)
page2 = input2.getPage(0)

total_width = page1.mediaBox.upperRight[0] + page2.mediaBox.upperRight[0]
total_height = max([page1.mediaBox.upperRight[1], page2.mediaBox.upperRight[1]])

new_page = PageObject.createBlankPage(None, total_width, total_height)

# Add first page at the 0,0 position
new_page.mergePage(page1)
# Add second page with moving along the axis x
new_page.mergeTranslatedPage(page2, page1.mediaBox.upperRight[0], 0)

output = PdfFileWriter()
output.addPage(new_page)
output.write(open("result.pdf", "wb"))

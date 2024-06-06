import  docx2pdf
from  pdf2docx import Converter
docx2pdf.convert("test1.docx","OutputFiles/test2.pdf")
pdfFile = "OutputFiles/UniconDoc43.pdf"
docxFile = "DocxOutputFiles/sample.docx"
cv = Converter(pdfFile)
cv.convert(docxFile)
cv.close()
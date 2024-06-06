import sys
from PyQt5 import QtWidgets, uic
from PyQt5.uic import loadUi
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import docx2pdf
import time
from  pdf2docx import Converter
from PyQt5.QtWidgets import QMessageBox
import os
class MainWindow(QDialog):
    def __init__(self):
        super(MainWindow,self).__init__()
        uic.loadUi(r"C:\Desktop\UniConDoc\ConverterUI.ui",self)
        self.browse.clicked.connect(self.DocxbrowseFiles)
        self.pdfbrowse.clicked.connect(self.PdfbrowseFiles)
        self.converter.clicked.connect(self.doc2pdf)
        self.converter_pdf.clicked.connect(self.pdf2doc)
        self.filename = self.findChild(QLineEdit,"LineEdit")
        self.pdffile = self.findChild(QLineEdit,"LineEdit_2")
        self.docxLabel = self.findChild(QLabel,"docxloading")
        self.pdflabel = self.findChild(QLabel,"pdfloading")
    def DocxbrowseFiles(self):
        self.fname = QFileDialog.getOpenFileName(self,'Open File','C:','Docx Files(*.docx)')
        print(self.fname[0])
        self.filename.setText(self.fname[0])



    def PdfbrowseFiles(self):
        self.pdffname = QFileDialog.getOpenFileName(self, 'Open File', 'C:', 'PDF Files(*.pdf)')
        print(self.pdffname[0])
        self.pdffile.setText(self.pdffname[0])

    def pdf2doc(self):
        if self.pdffile.text() == "":
            ret = QMessageBox.question(self, 'UniConDoc Asks',
                                       f"Please Select any File ",
                                       QMessageBox.Yes)
            if ret == QMessageBox.Yes:
                return 0;
        date = time.strftime("%S")
        print(date)
        print(self.pdffname[0])
        self.pdflabel.setText("Converting to Docx.....")
        ret = QMessageBox.question(self, 'UniConDoc Asks',
                                   f"Are You Sure to Convert Your Docx file{self.pdffname[0]} to PDF",
                                   QMessageBox.Yes | QMessageBox.No)
        if ret == QMessageBox.Yes:

            print('Button QMessageBox.Yes clicked.')
            pdfFile = f"{self.pdffname[0]}"
            docxFile = f"DocxOutputFiles/UniconDocx{date}.docx"
            cv = Converter(pdfFile)
            cv.convert(docxFile)
            cv.close()
            self.pdflabel.setText("Successfully Converted to Docx")


    def doc2pdf(self):
        if self.filename.text() == "":
            ret = QMessageBox.question(self, 'UniConDoc Asks',
                                       f"Please Select any File ",
                                       QMessageBox.Yes)
            if ret == QMessageBox.Yes:
                return 0;

        date = time.strftime("%S")
        print(date)
        print(self.fname[0])
        self.converter.setText("Converting ....")
        self.docxLabel.setText("Converting to PDF.....")

        ret = QMessageBox.question(self, 'UniConDoc Asks', f"Are You Sure to Convert Your Docx file{self.fname[0]} to PDF",
                                   QMessageBox.Yes | QMessageBox.No)

        if ret == QMessageBox.Yes:

            print('Button QMessageBox.Yes clicked.')

            docx2pdf.convert(f"{self.fname[0]}", f"OutputFiles/UniconDoc{date}.pdf")

            QMessageBox.about(self, " UniCornDoc Success ", f"Congrats ! Your File has been converted from docx to PDF and has been saved in the Directory UniconDoc{date}")
            self.filename.setText("")
            self.docxLabel.setText("SuccessFully Converted !!!!!!")
            self.converter.setText("Convert to Pdf ")



        elif ret==QMessageBox.No:
            return 0;





app = QApplication(sys.argv)
mainwindow = MainWindow()
widget = QtWidgets.QStackedWidget()
widget.setFixedWidth(640)
widget.setFixedHeight(477)
widget.addWidget(mainwindow)
widget.show()
sys.exit(app.exec_())

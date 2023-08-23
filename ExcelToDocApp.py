import sys, os, io
if hasattr(sys, 'frozen'):
    os.environ['PATH'] = sys._MEIPASS + ";" + os.environ['PATH']
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QPushButton, QListWidget, \
QVBoxLayout, QHBoxLayout, QGridLayout, \
QDialog, QFileDialog, QMessageBox, QAbstractItemView
from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QIcon

from pathlib import Path
import pandas as pd  
from docxtpl import DocxTemplate  
from docx import Document
from docxcompose.composer import Composer

from datetime import datetime 
from dateutil import parser

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

excelFiles =[]

class ListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent=None)
        self.setAcceptDrops(True)
        self.setStyleSheet('''font-size:25px''')
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            return super().dragEnterEvent(event) 

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            return super().dragMoveEvent(event) 

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()

            #excelFiles = []

            for url in event.mimeData().urls():
                if url.isLocalFile():
                    if url.toString().endswith('.xlsx'):   ### MOI DOI DUOI THANH .csv
                        excelFiles.append(str(url.toLocalFile()))
                        print("Excel Files +1")      ###Bug MARKER
            self.addItems(excelFiles)
        else:
            return super().dropEvent(event) 
    
class output_field(QLineEdit):
    def __init__(self):
        super().__init__()
        self.height = 55
        self.setStyleSheet('''font-size: 30px;''')
        self.setFixedHeight(self.height)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls:
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls:
            event.setDropAction(Qt.CopyAction)
            event.accept()

        if event.mimeData().urls():
            self.setText(event.mimeData().urls()[0].toLocalFile())
            excelFiles.append(str.url.toLocalFile())
        else:
            event.ignore()

class button(QPushButton):
    def __init__(self, label_text):
        super().__init__()
        self.setText(label_text)
        self.setStyleSheet('''
        font-size: 30px;
        width: 180px;
        height: 50;
        ''')

class AppDemo(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Biopig Excel To Doc App')
        self.resize(1800, 800)
        self.initUI()
        
    def initUI(self):
        mainLayout = QVBoxLayout()
        outputFolderRow = QHBoxLayout()
        buttonLayout = QHBoxLayout()

        self.outputFile = output_field()
        outputFolderRow.addWidget(self.outputFile)

        # browse button
        self.buttonBrowseOutputFile = button('&Save To')
        self.buttonBrowseOutputFile.clicked.connect(self.populateFileName)
        self.buttonBrowseOutputFile.setFixedHeight(self.outputFile.height)
        outputFolderRow.addWidget(self.buttonBrowseOutputFile)
        
        self.excelListWidget = ListWidget(self)

        """
        Buttons
        """
        self.buttonDeleteSelect = button('&Delete')
        self.buttonDeleteSelect.clicked.connect(self.deleteSelected) ### DA COMMENT PROBLEM LAI
        buttonLayout.addWidget(self.buttonDeleteSelect, 1, Qt.AlignRight)

        self.buttonGo = button('&Go')
        self.buttonGo.clicked.connect(self.convertExcelToDoc)  ## CHINH LAI FUNCTION CHO NAY 
        buttonLayout.addWidget(self.buttonGo)

        self.buttonClose = button('&Close')
        self.buttonClose.clicked.connect(QApplication.quit)
        buttonLayout.addWidget(self.buttonClose)

        # reset button
        self.buttonReset = button('&Reset')
        self.buttonReset.clicked.connect(self.clearQueue)
        buttonLayout.addWidget(self.buttonReset)
        
        
        mainLayout.addLayout(outputFolderRow)
        mainLayout.addWidget(self.excelListWidget)
        mainLayout.addLayout(buttonLayout)
        self.setLayout(mainLayout)
        
    def deleteSelected(self):
        for item in self.excelListWidget.selectedItems():
            self.excelListWidget.takeItem(self.excelListWidget.row(item))

    def clearQueue(self):
        self.excelListWidget.clear()
        self.outputFile.setText('')

    def dialogMessage(self, message):
        dlg = QMessageBox(self)
        dlg.setWindowTitle('EXCEL TO DOC APP')
        dlg.setIcon(QMessageBox.Information)
        dlg.setText(message)
        dlg.show()

    def populateFileName(self):
        path = self._getSaveFilePath()
        if path:
            self.outputFile.setText(path)
            
    def _getSaveFilePath(self):
        file_save_path, _ = QFileDialog.getSaveFileName(self, 'Save Doc file', os.getcwd(), 'Doc file (*.docx)')
        return file_save_path

    def convertExcelToDoc(self):
        if not self.outputFile.text():
            self.populateFileName()
            return 
    
        if self.excelListWidget.count()>0:
            try:
                df=pd.read_excel(excelFiles[0],sheet_name="Sheet1")    
                df.to_csv('file.csv') ### lUU Y: NEU NHAN FILE XLSX THI SE TAO FILE MOI CSV DE PHU HOP DINH DANG 
                df =pd.read_csv('file.csv')
                # print(self.excelListWidget.item(0))
                print(excelFiles[0])
                # df=pd.read_csv(excelFiles[0]) ### TEST VOI FILE XLSX
                output_doc =Document()
                composer = Composer(output_doc)
                for record in df.to_dict(orient="records"):
                    single_doc = DocxTemplate("Mau-ly-lich-heo-giong-new.docx")
                    # TRY TO CHANGE DATE FORMAT 
                        #if bool(parser.parse(record.####values())) ==True:
                            #record= datetime.strptime(record,'%Y-%m-%d').strftime('%d/%m/%Y')
                        #print(record)
                        #print(len(record))
                    single_doc.render(record)
                    #single_doc.add_page_break()
                    composer.append(single_doc)
                output_doc.save(self.outputFile.text())
                
                self.dialogMessage('Convert Excel to Doc Complete')
                print("Converted")
            
            except Exception as e: 
                self.dialogMessage(e)
                print('exception')
                
        else:
            self.dialogMessage('Queue is empty')        
            self.dialogMessage(f'There are {self.excelListWidget.count()} files here')    

app = QApplication(sys.argv)
app.setStyle('fusion')

ExcelToDocApp = AppDemo()
ExcelToDocApp.show()

sys.exit(app.exec_())

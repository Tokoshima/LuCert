from qtUI import Ui_MainWindow
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook, load_workbook
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt5.QtCore import QDir
import gettex
import docx
import openpyxl
import sys,os
import shutil



class MyWindow(QMainWindow, Ui_MainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        #self.initUI()
        #self.setupUi(self)

        self.setupUi(self)


        self.btnExc.clicked.connect(self.exc_clicked)
        self.btnUpload.clicked.connect(self.upload_clicked)

    def exc_clicked(self):
        self.update()

        #/home/louwste()r3000/Documents/LuCert
        cmbCol_selected = self.cmbCol.currentText()
        os.chdir(os.path.abspath('Lists'))
        #os.chdir(os.listdir(FILE_PATH))
        #/home/louwster3000/Documents/LuCert/Lists
        wb = load_workbook("List.xlsx")
        #print("Cek1")
        source = wb["Sheet1"]
        if self.rdbSep.isChecked():
            print("Is Checked")
        for cell in source['%c' % cmbCol_selected]:

            #/home/louwster3000/Documents/LuCert/Lists
            name =cell.value
            print(cell.value)
            if name is None:
                print("ERROR: Cell is invalid")
            # if not (isinstance(name, str))
            #     print("Cell is not a Name!")
            os.chdir(os.path.abspath('..'))
            #/home/louwster3000/Documents/LuCert

            doc = docx.Document('masterDoc.docx')
            p1 = doc.add_paragraph('%s' % name)
            p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER


            if os.path.exists(os.path.abspath("Certs")):
                os.chdir(os.path.abspath("Certs"))
                #/home/louwster3000/Documents/LuCert/Certs
                doc.save('%s.docx' % cell.value)



                #os.chdir(os.path.abspath('Lists'))
                print(os.getcwd())
            else:
                print('failed')
        os.chdir(os.path.abspath('..'))
        #print(os.getcwd())

    def upload_clicked(self):
        print(os.getcwd())
        xlsxListDir = QFileDialog.getOpenFileName(self)
        xlsxListDir = xlsxListDir[0]
        shutil.copyfile(xlsxListDir,os.getcwd()+'/Lists/List.xlsx')
        #os.chdir(os.listdir(FILE_PATH))


        if not xlsxListDir.endswith('.xlsx'):
            print("ERROR")


    #     dialog = QFileDialog
    #     dialog.setFileMode(QFileDialog.AnyFile)
    #     dialog.setFilter(QDir.Files)
    #
    #     if dialog.exec_():
    #         file_name = dialog.selectedFiles()
    #
    # if file_name[0].endswith('.py'):
    #     with open(file_name[0], 'r') as f:
    #         data = f.read()
    #         self.textEditor.setPlainText(data)
    #         f.close()
    # else:
    #     pass
    # def initUI(self):
    #     self.setGeometry(200, 200, 300, 300)
    #     self.setWindowTitle("Tech With Tim")
    #
    #     self.label = QtWidgets.QLabel(self)
    #     self.label.setText("my first label!")
    #
    #     self.label.move(50,50)
    #     self.b1 = QtWidgets.QPushButton(self)
    #     self.b1.setText("click me!")
    #     self.b1.clicked.connect(self.button_clicked)
    #
    # def update(self):
    #     self.label.adjustSize()
#
# def window():
#     app = QApplication(sys.argv)
#     win = MyWindow()
#     win.show()
#
#
#     sys.exit(app.exec_())
#
# window()

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = MyWindow()
    #MainWindow = QtWidgets.QMainWindow()
    # ui = Ui_MainWindow()
    # ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


    #name =""
    #doc = docx.Document('masterDoc.docx')
    #p1 = doc.add_paragraph('%s' % name)
    #p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    #doc.save('Out.docx')
    #print(gettex.getText('Out.docx'))

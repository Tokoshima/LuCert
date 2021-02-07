from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook, load_workbook
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow
import gettex
import docx
import openpyxl
import sys,os



class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow,self).__init__()
        self.initUI()

    def button_clicked(self):
        self.label.setText("you pressed the button")
        self.update()

        os.chdir(os.path.abspath('Lists'))
        wb = load_workbook("List.xlsx")
        #print("Cek1")
        source = wb["Sheet1"]
        for cell in source['%c' % col]:
            #print("Before chdir masterdoc")
            #print(os.getcwd())

            os.chdir(os.path.abspath('..'))
            #print("After chdir masterdoc")
            #print(os.getcwd())
            print(cell.value)
            name =cell.value
            doc = docx.Document('masterDoc.docx')
            p1 = doc.add_paragraph('%s' % name)
            p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER


            if os.path.exists(os.path.abspath("Certs")):
                os.chdir(os.path.abspath("Certs"))
                doc.save('%s.docx' % cell.value)
            else:
                print('failed')


    def initUI(self):
        self.setGeometry(200, 200, 300, 300)
        self.setWindowTitle("Tech With Tim")

        self.label = QtWidgets.QLabel(self)
        self.label.setText("my first label!")

        self.label.move(50,50)
        self.b1 = QtWidgets.QPushButton(self)
        self.b1.setText("click me!")
        self.b1.clicked.connect(self.button_clicked)

    def update(self):
        self.label.adjustSize()

def window():
    app = QApplication(sys.argv)
    win = MyWindow()
    win.show()
    sys.exit(app.exec_())

window()





    #name =""
    #doc = docx.Document('masterDoc.docx')
    #p1 = doc.add_paragraph('%s' % name)
    #p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    #doc.save('Out.docx')
    #print(gettex.getText('Out.docx'))

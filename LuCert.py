from docx.enum.text import WD_ALIGN_PARAGRAPH
import gettex
import docx
import openpyxl
import sys,os

from openpyxl import Workbook, load_workbook
os.chdir(os.path.abspath('Lists'))
wb = load_workbook("List.xlsx")
#print("Cek1")
source = wb["Sheet1"]
for cell in source['A']:
    #print("Before chdir masterdoc")
    #print(os.getcwd())
    #print(__file__)
    #print(os.path.join(os.path.dirname(__file__), '..'))
    #print(os.path.dirname(os.path.realpath(__file__)))
    #print(os.path.abspath(os.path.dirname(__file__)))

    os.chdir(os.path.abspath('..'))
    #print("After chdir masterdoc")
    #print(os.getcwd())
    #print(cell.value)
    name =cell.value
    doc = docx.Document('masterDoc.docx')
    p1 = doc.add_paragraph('%s' % name)
    p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    #os.chdir('~/Documents/LuCert/Certs')
    #sys.path.append('/Documents/LuCert/Certs')
    #print(os.path.abspath("Certs"))
    
    if os.path.exists(os.path.abspath("Certs")):
        os.chdir(os.path.abspath("Certs"))
        doc.save('%s.docx' % cell.value)
    else:
        print('failed')
   #doc.save('%s.docx' % cell.value)
   # print(gettex.getText('/Documents/LuCert/Certs/'+ '%s.docx' % cell.value))
 


#name =""
#doc = docx.Document('masterDoc.docx')
#p1 = doc.add_paragraph('%s' % name)
#p1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
#doc.save('Out.docx')
#print(gettex.getText('Out.docx'))


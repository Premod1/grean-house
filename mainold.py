


from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QTextEdit
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
import itertools
from docx.shared import Pt
from docx.enum.text import WD_BREAK
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import os
import platform
from docx.oxml.ns import qn
import pypandoc
import pdfkit
import asyncio
from pyppeteer import launch
from datetime import datetime


class Ui_Home(object):
    def setupUi(self, Home):
        Home.setObjectName("Home")
        Home.resize(739, 438)
        self.centralwidget = QtWidgets.QWidget(Home)
        self.centralwidget.setObjectName("centralwidget")
        self.upload_excel = QtWidgets.QPushButton(self.centralwidget)
        self.upload_excel.setGeometry(QtCore.QRect(180, 20, 361, 51))
        self.upload_excel.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.upload_excel.setMouseTracking(False)
        self.upload_excel.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.upload_excel.setObjectName("upload_excel")

        self.print_or_download = QtWidgets.QPushButton(self.centralwidget)
        self.print_or_download.setGeometry(QtCore.QRect(180, 90, 361, 51))
        self.print_or_download.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.print_or_download.setFocusPolicy(QtCore.Qt.ClickFocus)
        self.print_or_download.setObjectName("print_or_download")

        Home.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(Home)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 739, 22))
        self.menubar.setObjectName("menubar")
        Home.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(Home)
        self.statusbar.setObjectName("statusbar")
        Home.setStatusBar(self.statusbar)

        self.retranslateUi(Home)
        QtCore.QMetaObject.connectSlotsByName(Home)

        self.upload_excel.clicked.connect(self.upload_excel_file)
        self.print_or_download.clicked.connect(self.print_or_download_file)
       

    def retranslateUi(self, Home):
        _translate = QtCore.QCoreApplication.translate
        Home.setWindowTitle(_translate("Home", "MainWindow"))
        self.upload_excel.setText(_translate("Home", "Upload Excel"))
        

    def upload_excel_file(self):
     options = QFileDialog.Options()
     file_path, _ = QFileDialog.getOpenFileName(None, "Select Excel File", "", "Excel Files (*.xlsx *.xls)", options=options)
 
     if file_path:
            try:
                df = pd.read_excel(file_path)
                
                if 'Item' in df.columns and 'Invoice' in df.columns and 'Name' in df.columns and 'Price' in df.columns and 'City' in df.columns:
                    items = df['Item']
                    invoices = df['Invoice']
                    names = df['Name']
                    prices = df['Price']
                    citys = df['City']
                    
                    # Generate HTML table
                    html_content = '<html><body><table border="1">'
                    html_content += '<tr><th>From</th><th>To</th><th>From</th><th>To</th></tr>'
                    
                    paired_data = itertools.zip_longest(items, invoices, names, prices, citys)
                    for (item1, invoice1, name1, price1, city1), (item2, invoice2, name2, price2, city2) in zip(paired_data, paired_data):
                        html_content += f'''
                  <tr>
                      <td style="padding: 10px;">
                          <div style="text-align: center; font-size: 18px;">
                              From:<br>Green House Lk<br>Matale<br>0728883082<br>
                              <hr style="width: 80%; margin: 10px auto;"> <!-- Added consistent margin and width -->
                              <span style="font-size: 28px; font-weight: bold;">COD</span><br>
                              <span style="font-size: 28px; font-weight: bold;">RS.{price1}</span>
                          </div>
                      </td>
                      <td style="padding: 10px; font-size: 18px;">
                          To:<br>{name1}<br>
                          <h4>{city1}</h4>
                          <hr style="width: 80%; margin: 10px auto;"> <!-- Same styling for consistent alignment -->
                          Item: {item1}<br>Invoice No: {invoice1}
                      </td>
                      <td style="padding: 10px;">
                          <div style="text-align: center; font-size: 18px;">
                              From:<br>Green House Lk<br>Matale<br>0728883082<br>
                              <hr style="width: 80%; margin: 10px auto;"> <!-- Same styling for consistent alignment -->
                              <span style="font-size: 28px; font-weight: bold;">COD</span><br>
                              <span style="font-size: 28px; font-weight: bold;">RS.{price2}</span>
                          </div>
                      </td>
                      <td style="padding: 10px; font-size: 18px;">
                          To:<br>{name2}<br>
                          <h4>{city2}</h4>
                          <hr style="width: 80%; margin: 10px auto;"> <!-- Same styling for consistent alignment -->
                          Item: {item2}<br>Invoice No: {invoice2}
                      </td>
                  </tr>


                        '''
                    
                    html_content += '</table></body></html>'

                    # Save HTML content to file
                    with open('final_output.html', 'w') as file:
                        file.write(html_content)
                    
                    QMessageBox.information(None, "Success", "The HTML document was successfully created.")
                else:
                    QMessageBox.warning(None, "Error", "The selected file doesn't contain the required columns.")
            except Exception as e:
                QMessageBox.critical(None, "Error", f"Failed to read the Excel file: {str(e)}")

   
        
    def print_or_download_file(self):
      html_file_path = 'final_output.html'
      download_folder = os.path.join(os.path.expanduser('~'), 'Downloads')
      current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
      pdf_file_name = f'final_output_{current_time}.pdf'
      pdf_file_path = os.path.join(download_folder, pdf_file_name)
      print(html_file_path)
      pdfkit.from_file(html_file_path, pdf_file_path)
      QMessageBox.information(None, "Success", "Pdf Download Successfull.")
      os.remove(html_file_path)
  
      
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Home = QtWidgets.QMainWindow()
    ui = Ui_Home()
    ui.setupUi(Home)
    Home.show()
    sys.exit(app.exec_())
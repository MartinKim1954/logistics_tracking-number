import os
import sys
# PyQt5
from PyQt5.QtWidgets import *
from PyQt5 import uic
# Excel
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font
# CSV
import csv
# Data
import pandas as pd
# infos
import city_info

# MAC
form_class = uic.loadUiType(
    "/Users/martinkim/GITHUB/00_Automated System/02_Logistics-team/logistics_tracking-number")
# Window
# form_class = uic.loadUiType("")


class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.btn_import.clicked.connect(self.import_file)
        self.btn_exit.clicked.connect(self.exit)
        self.btn_interpret.clicked.connect(self.interpret_file)

    def import_file(self):
        self.file_name = QFileDialog.getOpenFileName(self)
        self.tb_import_directory.setText(self.file_name[0])

    def exit(self):
        QApplication.quit()

    def interpret_file(self):
        item_name = list()
        df = pd.read_csv(
            '/Users/martinkim/GITHUB/00_Automated System/02_Logistics-team/logistics_tracking-number/PO_SKU_LIST_20200904150707.csv')

        # progress bar value changes slowly
        self.pb_status.setValue(0)
        self.completed = 0
        while self.completed < 100:
            self.completed += 0.0001
            self.pb_status.setValue(self.completed)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()

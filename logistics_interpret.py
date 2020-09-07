import os
import sys
# PyQt5
from PyQt5.QtWidgets import *
from PyQt5 import uic
# infos
import interpret


# MAC
form_class = uic.loadUiType(
    "tracking-number.ui")[0]
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
        self.import_directory = self.tb_import_directory.toPlainText()

    def exit(self):
        QApplication.quit()

    def interpret_file(self):
        order_count, city_numbers, order_info = interpret.interpret(
            self.import_directory)
        # To show data at a glance
        self.tb_order_info.setText(
            '총 발주 건수: ' + str(order_count) + '건' + '\n총 도시 갯수: ' + str(city_numbers) + '개 도시' + '\n도시별 발주건수: ')
        output = ''
        for i in range(len(order_info)):
            output += (f'{order_info[i][0]}: {order_info[i][1]}건\n')
        self.tb_order_info.append(output)
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

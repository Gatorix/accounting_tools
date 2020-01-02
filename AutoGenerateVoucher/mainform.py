import sys
from PyQt5 import QtWidgets
from gui.Ui_main import Ui_Dialog
from PyQt5.QtCore import QCoreApplication
# import generater.logicalfunc as func


class Main(QtWidgets.QWidget, Ui_Dialog):
    def __init__(self):
        super(Main, self).__init__()
        self.setupUi(self)
        self.conf()
        self.show()

    def conf(self):
        self.pushButton.clicked.connect(self.pushButton_on_click)
        self.pushButton_2.clicked.connect(QCoreApplication.instance().quit)

        self.comboBox.currentIndexChanged.connect(
            lambda: self.get_current_value())
        self.comboBox_2.currentIndexChanged.connect(
            lambda: self.get_current_value())
        self.comboBox_3.currentIndexChanged.connect(
            lambda: self.get_current_value())

    def get_current_value(self):
        voucher_type = self.comboBox.currentText()
        voucher_bookkeeper = self.comboBox_2.currentText()
        business_type = self.comboBox_3.currentText()
        return voucher_type, voucher_bookkeeper, business_type

    def pushButton_on_click(self):
        print('按钮1')


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    main = Main()
    app.exec_()

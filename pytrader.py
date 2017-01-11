import sys
from PyQt5.QtWidgets import QApplication
from PyQt5 import uic


class MyWindow(object):

    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("pytrader.ui")
        self.ui.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())
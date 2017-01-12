import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from PyQt5 import uic
import Kiwoom


class MyWindow(object):

    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("pytrader.ui")
        self.ui.show()

        self.kiwoom = Kiwoom()
        self.kiwoom.commConnect()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

    def timeout(self):
        self.ui.statusbar.showMessage(self.kiwoom.getConnectState())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())
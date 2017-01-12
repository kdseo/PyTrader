import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtCore import QTimer, QTime
from PyQt5 import uic
from Kiwoom import Kiwoom


ui = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, ui):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()

        self.kiwoom = Kiwoom()
        self.kiwoom.commConnect()

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

    def timeout(self):
        currentTime = QTime.currentTime().toString("hh:mm:ss")
        state = ""

        if self.kiwoom.getConnectState() == 1:
            state = "서버 연결중"
        else:
            state = "서버 미연결"

        self.statusbar.showMessage("현재시간: " + currentTime + " | " + state)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())

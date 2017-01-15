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
        self.codeList = self.kiwoom.getCodeList("0")

        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        self.setAccountComboBox()
        self.codeLineEdit.textChanged.connect(self.setCodeName)
        self.orderBtn.clicked.connect(self.sendOrder)

    def timeout(self):
        """ 타임아웃 이벤트가 발생하면 호출되는 메서드 """

        currentTime = QTime.currentTime().toString("hh:mm:ss")
        state = ""

        if self.kiwoom.getConnectState() == 1:
            state = "서버 연결중"
        else:
            state = "서버 미연결"

        self.statusbar.showMessage("현재시간: " + currentTime + " | " + state)

    def setCodeName(self):
        """ 종목코드에 해당하는 한글명을 codeNameLineEdit에 설정한다. """

        code = self.codeLineEdit.text()

        if code in self.codeList:
            codeName = self.kiwoom.getMasterCodeName(code)
            self.codeNameLineEdit.setText(codeName)

    def setAccountComboBox(self):
        """ accountComboBox에 계좌번호를 설정한다. """

        cnt = int(self.kiwoom.getLoginInfo("ACCOUNT_CNT"))
        accountList = self.kiwoom.getLoginInfo("ACCNO").split(';')
        self.accountComboBox.addItems(accountList[0:cnt])

    def sendOrder(self):
        """ 키움서버로 주문정보를 전송한다. """

        orderTypeTable = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hogaTypeTable = {'지정가': "00", '시장가': "03"}

        account = self.accountComboBox.currentText()
        orderType = orderTypeTable[self.orderTypeComboBox.currentText()]
        code = self.codeLineEdit.text()
        hogaType = hogaTypeTable[self.hogaTypeComboBox.currentText()]
        qty = self.qtySpinBox.value()
        price = self.priceSpinBox.value()

        self.kiwoom.sendOrder("sendOrder_req", "0101", account, orderType,
                              code, qty, price, hogaType, "")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())

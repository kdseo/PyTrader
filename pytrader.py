"""
QtDesigner로 만든 UI와 해당 UI의 위젯에서 발생하는 이벤트를 컨트롤하는 클래스

author: 서경동
last edit: 2017. 01. 18
"""


import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from PyQt5.QtCore import QTimer, QTime
from PyQt5 import uic
from Kiwoom import Kiwoom, ParameterTypeError, ReturnCode, KiwoomProcessingError


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

        try:
            returnCode = self.kiwoom.sendOrder("sendOrder_req", "0101", account, orderType, code, qty, price, hogaType, "")
            self.showDialog('Information', "sendOrer() 결과: " + ReturnCode.CAUSE[returnCode])

        except (ParameterTypeError, KiwoomProcessingError) as e:
            self.showDialog('Critical', e)

    def showDialog(self, grade, text):
        gradeTable = {'Information': 1, 'Warning': 2, 'Critical': 3, 'Question': 4}

        dialog = QMessageBox()
        dialog.setIcon(gradeTable[grade])
        dialog.setText(text)
        dialog.setWindowTitle(grade)
        dialog.setStandardButtons(QMessageBox.Ok)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())

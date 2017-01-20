"""
QtDesigner로 만든 UI와 해당 UI의 위젯에서 발생하는 이벤트를 컨트롤하는 클래스

author: 서경동
last edit: 2017. 01. 18
"""


import sys, time
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt, QTimer, QTime
from PyQt5 import uic
from Kiwoom import Kiwoom, ParameterTypeError, ParameterValueError, KiwoomProcessingError


ui = uic.loadUiType("pytrader.ui")[0]

class MyWindow(QMainWindow, ui):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.show()

        self.kiwoom = Kiwoom()
        self.kiwoom.commConnect()
        self.codeList = self.kiwoom.getCodeList("0")

        # 메인 타이머
        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        # 잔고 및 보유종목 조회 타이머
        self.inquiryTimer = QTimer(self)
        self.inquiryTimer.start(1000*10)
        self.inquiryTimer.timeout.connect(self.timeout)

        self.setAccountComboBox()
        self.codeLineEdit.textChanged.connect(self.setCodeName)
        self.orderBtn.clicked.connect(self.sendOrder)
        self.inquiryBtn.clicked.connect(self.inquiryBalance)

        # 자동 주문
        self.automaticOrder()

        # 자동 선정 종목 리스트 테이블 설정
        self.setAutomatedStocks()

    def timeout(self):
        """ 타임아웃 이벤트가 발생하면 호출되는 메서드 """

        sender = self.sender()

        # 메인 타이머
        if id(sender) == id(self.timer):
            currentTime = QTime.currentTime().toString("hh:mm:ss")
            state = ""

            if self.kiwoom.getConnectState() == 1:
                state = "서버 연결중"
            else:
                state = "서버 미연결"

            self.statusbar.showMessage("현재시간: " + currentTime + " | " + state)

        # 실시간 조회 타이머
        else:
            if self.realtimeCheckBox.isChecked():
                self.inquiryBalance()

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
            self.kiwoom.returnCode = self.kiwoom.sendOrder("sendOrder_req", "0101", account, orderType, code, qty, price, hogaType, "")
            self.inquiryBalance()

        except (ParameterTypeError, KiwoomProcessingError) as e:
            self.showDialog('Critical', e)

    def inquiryBalance(self):
        """ 예수금상세현황과 계좌평가잔고내역을 요청후 테이블에 출력한다. """

        self.inquiryTimer.stop()

        try:
            # 예수금상세현황요청
            self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
            self.kiwoom.setInputValue("비밀번호", "0000")
            self.kiwoom.commRqData("예수금상세현황요청", "opw00001", 0, "2000")

            # 계좌평가잔고내역요청 - opw00018 은 한번에 20개의 종목정보를 반환
            self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
            self.kiwoom.setInputValue("비밀번호", "0000")
            self.kiwoom.commRqData("계좌평가잔고내역요청", "opw00018", 0, "2000")

            while self.kiwoom.inquiry == '2':
                time.sleep(0.2)

                self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
                self.kiwoom.setInputValue("비밀번호", "0000")
                self.kiwoom.commRqData("계좌평가잔고내역요청", "opw00018", 2, "2")

        except (ParameterTypeError, ParameterValueError, KiwoomProcessingError) as e:
            self.showDialog('Critical', e)

        # accountEvaluationTable 테이블에 정보 출력
        item = QTableWidgetItem(self.kiwoom.opw00001Data)   # d+2추정예수금
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.accountEvaluationTable.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018Data['accountEvaluation'][i-1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.accountEvaluationTable.setItem(0, i, item)

        self.accountEvaluationTable.resizeRowsToContents()

        # stocksTable 테이블에 정보 출력
        cnt = len(self.kiwoom.opw00018Data['stocks'])
        self.stocksTable.setRowCount(cnt)

        for i in range(cnt):
            row = self.kiwoom.opw00018Data['stocks'][i]

            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.stocksTable.setItem(i, j, item)

        self.stocksTable.resizeRowsToContents()

        # 데이터 초기화
        self.kiwoom.opwDataReset()

        # inquiryTimer 재시작
        self.inquiryTimer.start(1000*10)

    # 경고창
    def showDialog(self, grade, error):
        gradeTable = {'Information': 1, 'Warning': 2, 'Critical': 3, 'Question': 4}

        dialog = QMessageBox()
        dialog.setIcon(gradeTable[grade])
        dialog.setText(error.msg)
        dialog.setWindowTitle(grade)
        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.exec_()

    def setAutomatedStocks(self):
        fileList = ["buy_list.txt", "sell_list.txt"]
        automatedStocks = []

        for file in fileList:
            # utf-8로 작성된 파일을
            # cp949 환경에서 읽기위해서 encoding 지정
            with open(file, 'rt', encoding='utf-8') as f:
                stocksList = f.readlines()
                automatedStocks += stocksList

        # 테이블 행수 설정
        cnt = len(automatedStocks)
        self.automatedStocksTable.setRowCount(cnt)

        # 테이블에 출력
        for i in range(cnt):
            stocks = automatedStocks[i].split(';')

            for j in range(len(stocks)):
                if j == 1:
                    name = self.kiwoom.getMasterCodeName(stocks[j].rstrip())
                    item = QTableWidgetItem(name)
                else:
                    item = QTableWidgetItem(stocks[j].rstrip())

                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                self.automatedStocksTable.setItem(i, j, item)

        self.automatedStocksTable.resizeRowsToContents()

    def automaticOrder(self):
        fileList = ["buy_list.txt", "sell_list.txt"]
        hogaTypeTable = {'지정가': "00", '시장가': "03"}
        account = self.accountComboBox.currentText()
        automatedStocks = []

        # 파일읽기
        for file in fileList:
            # utf-8로 작성된 파일을
            # cp949 환경에서 읽기위해서 encoding 지정
            with open(file, 'rt', encoding='utf-8') as f:
                stocksList = f.readlines()
                automatedStocks += stocksList

        cnt = len(automatedStocks)

        # 주문하기
        buyResult = []
        sellResult = []

        for i in range(cnt):
            stocks = automatedStocks[i].split(';')

            code = stocks[1]
            hoga = stocks[2]
            qty = stocks[3]
            price = stocks[4]

            try:
                if stocks[5].rstrip() == '매수전':
                    self.kiwoom.sendOrder("sendOrder_req", "0101", account, 1, code, qty, price, hogaTypeTable[hoga], "")
                    buyResult += automatedStocks[i].replace("매수전", "매수주문완료")

                # 참고: 해당 종목을 현재도 보유하고 있다고 가정함.
                elif stocks[5].rstrip() == '매도전':
                    self.kiwoom.sendOrder("sendOrder_req", "0101", account, 2, code, qty, price, hogaTypeTable[hoga], "")
                    sellResult += automatedStocks[i].replace("매도전", "매도주문완료")

                self.inquiryBalance()

            except (ParameterTypeError, KiwoomProcessingError) as e:
                self.showDialog('Crtical', e)

        # 결과저장하기
        for file, result in zip(fileList, [buyResult, sellResult]):
            with open(file, 'wt', encoding='utf-8') as f:
                for data in result:
                    f.write(data)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())

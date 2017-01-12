import sys
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        self.loginLoop = QEventLoop()
        self.rqLoop = QEventLoop()
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveTrData.connect(self.receiveTrData())

    def commConnect(self):
        self.dynamicCall("CommConnect()")
        self.loginLoop.exec_()

    def eventConnect(self, errCode):
        if errCode == 0:
            print("connected")
        else:
            print("not connected")

        self.loginLoop.exit()

    def getCodeListByMarket(self, market):
        func = 'GetCodeListByMarket("%s")' % market
        codeList = self.dynamicCall(func)
        return codeList.split(';')

    def getMasterCodeName(self, code):
        func = 'GetMasterCodeName("%s")' % code
        name = self.dynamicCall(func)
        return name

    def getLoginInfo(self, tag):
        func = 'GetLoginInfo("%s")' % tag
        info = self.dynamicCall(func)
        return info

    def setInputValue(self, key, value):
        self.dynamicCall("SetInputValue(QString, QString)", key, value)

    def commRqData(self, requestName, trCode, prevNext, screenNo):
        self.dynamicCall("CommRqData(QString, QString, int, QString)", requestName, trCode, prevNext, screenNo)
        self.rqLoop.exec_()

    def commGetData(self, trCode, realType, fieldName, index, key):
        data = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trCode, realType, fieldName, index, key)
        return data.strip()

    def getRepeatCnt(self, code, recordName):
        ret = self.dynamicCall("GetRepeatCnt(QString, QString)", code, recordName)
        return ret

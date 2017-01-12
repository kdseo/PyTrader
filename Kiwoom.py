import sys
import time
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import QApplication


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        self.loginLoop = None
        self.rqLoop = None
        self.prevNext = 0
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveTrData.connect(self.receiveTrData)

    def receiveTrData(self, screenNo, requestName, trCode, recordName, prevNext, dontUse1, dontUse2, dontUse3, dontUse4):
        """ 데이터 수신시 발생하는 이벤트 """

        self.prevNext = prevNext

        if requestName == "opt10081_req":
            cnt = self.getRepeatCnt(trCode, requestName)

            for i in range(cnt):
                date = self.commGetData(trCode, "", requestName, i, "일자")
                open = self.commGetData(trCode, "", requestName, i, "시가")
                high = self.commGetData(trCode, "", requestName, i, "고가")
                low = self.commGetData(trCode, "", requestName, i, "저가")
                close = self.commGetData(trCode, "", requestName, i, "현재가")
                print(date, ": ", open, ' ', high, ' ', low, ' ', close)

        self.rqLoop.exit()

    def eventConnect(self, errCode):
        """
        통신 연결 상태 변경시 이벤트

        :param errCode: int
        """

        if errCode == 0:
            print("connected")
        else:
            print("not connected")

        self.loginLoop.exit()

    def commConnect(self):
        """ 로그인 윈도우를 실행한다. """

        self.dynamicCall("CommConnect()")
        self.loginLoop = QEventLoop()
        self.loginLoop.exec_()

    def getCodeListByMarket(self, market):
        """
        시장 구분에 따른 종목코드의 목록을 List로 반환한다.

        market에 올 수 있는 값은 아래와 같다.
        0: 장내, 3: ELW, 4: 뮤추얼펀드, 5: 신주인수권, 6: 리츠, 8: ETF, 9: 하이일드펀드, 10: 코스닥, 30: 제3시장

        :param market: string
        :return: List
        """

        func = 'GetCodeListByMarket("%s")' % market
        codeList = self.dynamicCall(func)
        return codeList.split(';')

    def getMasterCodeName(self, code):
        """
        종목코드의 한글명을 반환한다.

        :param code: string
        :return: string
        """

        func = 'GetMasterCodeName("%s")' % code
        name = self.dynamicCall(func)
        return name

    def getLoginInfo(self, tag):
        """
        사용자의 tag에 해당하는 정보를 반환한다.

        tag에 올 수 있는 값은 아래와 같다.
        ACCOUNT_CNT: 전체 계좌의 개수를 반환한다.
        ACCNO: 전체 계좌 목록을 반환한다. 계좌별 구분은 ;(세미콜론) 이다.
        USER_ID: 사용자 ID를 반환한다.
        USER_NAME: 사용자명을 반환한다.

        :param tag: string
        :return: string
        """

        func = 'GetLoginInfo("%s")' % tag
        info = self.dynamicCall(func)
        return info

    def setInputValue(self, key, value):
        """
        TR 전송에 필요한 값을 설정한다.

        :param key: string
        :param value: string
        """

        self.dynamicCall("SetInputValue(QString, QString)", key, value)

    def commRqData(self, requestName, trCode, prevNext, screenNo):
        """
        키움서버에 TR 요청을 한다.

        :param requestName: string - TR을 구분하기 위해 개발자가 정의
        :param trCode: string
        :param prevNext: int - 0(조회), 2(연속)
        :param screenNo: string - 화면번호(4자리)
        :return: int
        """

        self.dynamicCall("CommRqData(QString, QString, int, QString)", requestName, trCode, prevNext, screenNo)
        self.rqLoop = QEventLoop()
        self.rqLoop.exec_()

    def commGetData(self, trCode, realType, requestName, index, key):
        """
        요청한 TR의 반환 값을 가져온다.

        :param trCode: string
        :param realType: string - TR 요청시 ""(빈문자)로 처리
        :param requestName: string
        :param index: int
        :param key: string
        :return: string
        """

        data = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                trCode, realType, requestName, index, key)
        return data.strip()

    def getRepeatCnt(self, trCode, requestName):
        """
        requestName으로 요청한 TR의 반환 값의 index 수를 반환 합니다.

        :param trCode: string
        :param requestName: string
        :return: int
        """

        count = self.dynamicCall("GetRepeatCnt(QString, QString)", trCode, requestName)
        return count

    def getConnectState(self):
        """
        현재 접속상태를 반환합니다.

        반환되는 접속상태는 아래와 같습니다.
        0: 미연결, 1: 연결

        :return: int
        """

        state = self.dynamicCall("GetConnectState()")
        return state


if __name__ == "__main__":
    app = QApplication(sys.argv)

    kiwoom = Kiwoom()
    kiwoom.commConnect()

    kiwoom.setInputValue("종목코드", "039490")
    kiwoom.setInputValue("기준일자", "20160624")
    kiwoom.setInputValue("수정주가구분", 1)

    kiwoom.commRqData("opt10081_req", "opt10081", 0, "0101")

    while kiwoom.prevNext == '2':
        time.sleep(0.2)

        kiwoom.setInputValue("종목코드", "039490")
        kiwoom.setInputValue("기준일자", "20160624")
        kiwoom.setInputValue("수정주가구분", 1)

        kiwoom.commRqData("opt10081_req", "opt10081", 2, "0101")

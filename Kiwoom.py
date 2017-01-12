import sys
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()

        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")
        self.loginLoop = None
        self.rqLoop = None
        self.prevNext = 0
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveTrData.connect(self.receiveTrData())

    def receiveTrData(self, screenNo, requestName, trCode, recordName, prevNext):
        self.prevNext = prevNext
        raise NotImplementedError()

    def commConnect(self):
        """ 로그인 윈도우를 실행한다. """

        self.dynamicCall("CommConnect()")
        self.loginLoop = QEventLoop()
        self.loginLoop.exec_()

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
        사용자의 tag 정보를 반환한다.

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
        self.rqLoop.exec_()

    def commGetData(self, trCode, realType, recordName, index, key):
        """
        요청한 TR의 반환 값을 가져온다.

        :param trCode: string
        :param realType: string - TR 요청시 ""(빈문자)로 처리
        :param recordName: string
        :param index: int
        :param key: string
        :return: string
        """

        data = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                trCode, realType, recordName, index, key)
        return data.strip()

    def getRepeatCnt(self, trCode, recordName):
        """
        요청한 TR의 반환 값중 recordName의 index 수를 반환한다.

        :param trCode: string
        :param recordName: string
        :return: int
        """

        count = self.dynamicCall("GetRepeatCnt(QString, QString)", trCode, recordName)
        return count

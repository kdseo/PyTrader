"""
Kiwoom 클래스는 OCX를 통해 API 함수를 호출할 수 있도록 구현되어 있습니다.
OCX 사용을 위해 QAxWidget 클래스를 상속받아서 구현하였으며,
주식(현물) 거래에 필요한 메서드들만 구현하였습니다.

author: 서경동
last edit: 2017. 01. 14
"""


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
        self.OnReceiveChejanData.connect(self.receiveChejanData)

    # 이벤트 정의
    def receiveChejanData(self, gubun, itemCnt, fidList):
        """ 주문 체결시 발생하는 이벤트 """

        print("gubun: ", gubun)
        print("주문번호: ", self.getChejanData(9203))
        print("종목명: ", self.getChejanData(302))
        print("주문수량: ", self.getChejanData(900))
        print("주문가격: ", self.getChejanData(901))

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

    # 메서드 정의
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

        cmd = 'GetCodeListByMarket("%s")' % market
        codeList = self.dynamicCall(cmd)
        return codeList.split(';')

    def getMasterCodeName(self, code):
        """
        종목코드의 한글명을 반환한다.

        :param code: string
        :return: string
        """

        cmd = 'GetMasterCodeName("%s")' % code
        name = self.dynamicCall(cmd)
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

        cmd = 'GetLoginInfo("%s")' % tag
        info = self.dynamicCall(cmd)
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

    def sendOrder(self, requestName, screenNo, accountNo, orderType, code, qty, price, hogaGb, originOrderNo):
        """
        주식 주문을 키움서버로 전송한다.

        :param requestName: string - 요청을 구분하기 위해서 개발자가 붙인 요청명
        :param screenNo: string - 화면번호(4자리)
        :param accountNo: string - 계좌번호(10자리)
        :param orderType: int - 주문유형(1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정)
        :param code: string - 종목코드
        :param qty: int - 주문수량
        :param price: int - 주문단가
        :param hogaGb: string - 거래구분(00: 지정가, 03: 시장가, 05: 조건부지정가, 06: 최유리지정가, 그외에는 api 문서참조)
        :param originOrderNo: string - 원 주문번호
        :return: int - api 문서의 에러코드표 참조
        """

        errCode = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                   [requestName, screenNo, accountNo, orderType, code, qty, price, hogaGb, originOrderNo])
        return errCode

    def getChejanData(self, fid):
        """
        체결잔고 데이터를 반환한다.

        주문 체결관련 FID
        9202: 주문번호, 302: 종목명, 900: 주문수량, 901: 주문가격, 902: 미체결수량, 904: 원주문번호,
        905: 주문구분, 908: 주문/체결시간, 909: 체결번호, 910: 체결가, 911: 체결량, 10: 현재가, 체결가, 실시간 종가
        그 외의 FID는 api 문서 참조

        :param fid: int
        :return: string
        """

        cmd = 'GetChejanData("%s")' % fid
        data = self.dynamicCall(cmd)
        return data


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

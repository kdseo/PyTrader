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

        # instance var
        self.loginLoop = None
        self.rqLoop = None
        self.inquiry = 0

        # signal & slot
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveTrData.connect(self.receiveTrData)
        self.OnReceiveChejanData.connect(self.receiveChejanData)
        self.OnReceiveRealData.connect(self.receiveRealData)
        self.OnReceiveMsg.connect(self.receiveMsg)

    # 이벤트 정의
    def receiveRealData(self, code, realType, realData):
        print("receiveRealData 실행")

    def receiveMsg(self, screenNo, requestName, trCode, msg):
        print("receiveMsg 실행")

    def receiveChejanData(self, gubun, itemCnt, fidList):
        """
        주문 접수/확인 수신시 이벤트

        :param gubun: string - 체결구분('0': 주문체결통보, '1': 잔고통보, '3': 특이신호)
        :param itemCnt: int
        :param fidList: string - 데이터 구분은 ;(세미콜론) 이다.
        """

        print("receiveChejanData 실행")

        print("gubun: ", gubun)
        print("주문번호: ", self.getChejanData(9203))
        print("종목명: ", self.getChejanData(302))
        print("주문수량: ", self.getChejanData(900))
        print("주문가격: ", self.getChejanData(901))

    def receiveTrData(self, screenNo, requestName, trCode, recordName, inquiry,
                      deprecated1, deprecated2, deprecated3, deprecated4):
        """
        TR 수신시 이벤트

        requestName과 trCode는 commRqData()메소드의 매개변수와 매핑되는 값 이다.

        :param screenNo: string - 화면번호(4자리)
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :param trCode: string
        :param recordName: string
        :param inquiry: string - 조회('0': 남은 데이터 없음, '2': 남은 데이터 있음)
        :return:
        """

        print("receiveTrData 실행")

        self.inquiry = inquiry

        if requestName == "opt10081_req":
            cnt = self.getRepeatCnt(trCode, requestName)

            for i in range(cnt):
                date = self.commGetData(trCode, "", requestName, i, "일자")
                open = self.commGetData(trCode, "", requestName, i, "시가")
                high = self.commGetData(trCode, "", requestName, i, "고가")
                low = self.commGetData(trCode, "", requestName, i, "저가")
                close = self.commGetData(trCode, "", requestName, i, "현재가")
                print(date, ": ", open, ' ', high, ' ', low, ' ', close)

        try:
            # commRqData()에서 발생시킨 루프를 종료시킨다.
            self.rqLoop.exit()
        except AttributeError:
            pass

    def eventConnect(self, errCode):
        """
        통신 연결 상태 변경시 이벤트

        errCode가 0이면 로그인 성공
        그 외에는 에러코드표 참조

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
        '0': 장내, '3': ELW, '4': 뮤추얼펀드, '5': 신주인수권, '6': 리츠, '8': ETF, '9': 하이일드펀드, '10': 코스닥, '30': 제3시장

        :param market: string
        :return: List
        """

        if not isinstance(market, str):
            raise ParameterTypeError()

        cmd = 'GetCodeListByMarket("%s")' % market
        codeList = self.dynamicCall(cmd)
        return codeList.split(';')

    def getCodeList(self, *market):
        """
        여러 시장의 종목코드를 List 형태로 반환하는 헬퍼 메서드.

        :param market: Tuple - 여러 개의 문자열을 매개변수로 받아 Tuple로 처리한다.
        :return: List
        """

        codeList = []

        for m in market:
            tmpList = self.getCodeListByMarket(m)
            codeList += tmpList

        return codeList

    def getMasterCodeName(self, code):
        """
        종목코드의 한글명을 반환한다.

        :param code: string - 종목코드
        :return: string - 종목코드의 한글명
        """

        if not isinstance(code, str):
            raise ParameterTypeError()

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

        if not isinstance(tag, str):
            raise ParameterTypeError()

        cmd = 'GetLoginInfo("%s")' % tag
        info = self.dynamicCall(cmd)
        return info

    def setInputValue(self, key, value):
        """
        TR 전송에 필요한 값을 설정한다.

        :param key: string
        :param value: string
        """

        if not (isinstance(key, str) and isinstance(value, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetInputValue(QString, QString)", key, value)

    def commRqData(self, requestName, trCode, inquiry, screenNo):
        """
        키움서버에 TR 요청을 한다.

        :param requestName: string - TR 요청명(사용자 정의)
        :param trCode: string
        :param inquiry: int - 0(조회), 2(연속)
        :param screenNo: string - 화면번호(4자리)
        :return: int
        """

        if not (isinstance(requestName, str)
                and isinstance(trCode, str)
                and isinstance(inquiry, int)
                and isinstance(screenNo, str)):

            raise ParameterTypeError()

        self.dynamicCall("CommRqData(QString, QString, int, QString)", requestName, trCode, inquiry, screenNo)
        self.rqLoop = QEventLoop()
        self.rqLoop.exec_()

    def commGetData(self, trCode, realType, requestName, index, key):
        """
        요청한 TR의 반환 값을 가져온다.

        :param trCode: string
        :param realType: string - TR 요청시 ""(빈문자)로 처리
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :param index: int
        :param key: string
        :return: string
        """

        if not (isinstance(trCode, str)
                and isinstance(realType, str)
                and isinstance(requestName, str)
                and isinstance(index, int)
                and isinstance(key, str)):

            raise ParameterTypeError()

        data = self.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                trCode, realType, requestName, index, key)
        return data.strip()

    def getRepeatCnt(self, trCode, requestName):
        """
        requestName으로 요청한 TR의 반환 값의 index 수를 반환 합니다.

        :param trCode: string
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :return: int
        """

        if not (isinstance(trCode, str)
                and isinstance(requestName, str)):

            raise ParameterTypeError()

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

    def sendOrder(self, requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo):
        """
        주식 주문을 키움서버로 전송한다.

        sendOrder() 메소드 실행시,
        OnReceiveMsg, OnReceiveTrData, OnReceiveChejanData 이벤트가 발생한다.
        이 중, 주문에 대한 결과 데이터를 얻기 위해서는 OnReceiveChejanData 이벤트를 통해서 처리한다.

        :param requestName: string - 주문 요청명(사용자 정의)
        :param screenNo: string - 화면번호(4자리)
        :param accountNo: string - 계좌번호(10자리)
        :param orderType: int - 주문유형(1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정)
        :param code: string - 종목코드
        :param qty: int - 주문수량
        :param price: int - 주문단가
        :param hogaType: string - 거래구분(00: 지정가, 03: 시장가, 05: 조건부지정가, 06: 최유리지정가, 그외에는 api 문서참조)
        :param originOrderNo: string - 원 주문번호
        :return: int - api 문서의 에러코드표 참조
        """

        if not (isinstance(requestName, str)
                and isinstance(screenNo, str)
                and isinstance(accountNo, str)
                and isinstance(orderType, int)
                and isinstance(code, str)
                and isinstance(qty, int)
                and isinstance(price, int)
                and isinstance(hogaType, str)
                and isinstance(originOrderNo, str)):

            raise ParameterTypeError()

        errCode = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                   [requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo])

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

        if not isinstance(fid, int):
            raise ParameterTypeError()

        cmd = 'GetChejanData("%s")' % fid
        data = self.dynamicCall(cmd)
        return data


class ParameterTypeError(Exception):
    """ 파라미터 타입이 일치하지 않을 경우 발생하는 예외 """

    def __init__(self, msg="파라미터 타입이 일치하지 않습니다."):
        self.msg = msg

    def __str__(self):
        return self.msg


class ReturnCode(object):

    OP_ERR_NONE = 0 # 정상처리
    OP_ERR_FAIL = -10   # 실패
    OP_ERR_LOGIN = -100 # 사용자정보교환실패
    OP_ERR_CONNECT = -101   # 서버접속실패
    OP_ERR_VERSION = -102   # 버전처리실패
    OP_ERR_FIREWALL = -103  # 개인방화벽실패
    OP_ERR_MEMORY = -104    # 메모리보호실패
    OP_ERR_INPUT = -105 # 함수입력값오류
    OP_ERR_SOCKET_CLOSED = -106 # 통신연결종류
    OP_ERR_SISE_OVERFLOW = -200 # 시세조회과부하
    OP_ERR_RQ_STRUCT_FAIL = -201    # 전문작성초기화실패
    OP_ERR_RQ_STRING_FAIL = -202    # 전문작성입력값오류
    OP_ERR_NO_DATA = -203   # 데이터없음
    OP_ERR_OVER_MAX_DATA = -204 # 조회가능한종목수초과
    OP_ERR_DATA_RCV_FAIL = -205 # 데이터수신실패
    OP_ERR_OVER_MAX_FID = -206  # 죄회가능한FID수초과
    OP_ERR_REAL_CANCEL = -207   # 실시간해제오류
    OP_ERR_ORD_WRONG_INPUT = -300   # 입력값오류
    OP_ERR_ORD_WRONG_ACCTNO = -301  # 계좌비밀번호없음
    OP_ERR_OTHER_ACC_USE = -302 # 타인계좌사용오류
    OP_ERR_MIS_2BILL_EXC = -303 # 주문가격이20억원을초과
    OP_ERR_MIS_5BILL_EXC = -304 # 주문가격이50억원을초과
    OP_ERR_MIS_1PER_EXC = -305  # 주문수량이총발행주수의1%초과오류
    OP_ERR_MIS_3PER_EXC = -306  # 주문수량이총발행주수의3%초과오류
    OP_ERR_SEND_FAIL = -307 # 주문전송실패
    OP_ERR_ORD_OVERFLOW = -308  # 주문전송과부하
    OP_ERR_MIS_300CNT_EXC = -309    # 주문수량300계약초과
    OP_ERR_MIS_500CNT_EXC = -310    # 주문수량500계약초과
    OP_ERR_ORD_WRONG_ACCTINFO = -340    # 계좌정보없음
    OP_ERR_ORD_SYMCODE_EMPTY = -500 # 종목코드없음


if __name__ == "__main__":
    app = QApplication(sys.argv)

    kiwoom = Kiwoom()
    kiwoom.commConnect()

    kiwoom.setInputValue("종목코드", "039490")
    kiwoom.setInputValue("기준일자", "20160624")
    kiwoom.setInputValue("수정주가구분", "1")

    kiwoom.commRqData("opt10081_req", "opt10081", 0, "0101")

    while kiwoom.inquiry == '2':
        time.sleep(0.2)

        kiwoom.setInputValue("종목코드", "039490")
        kiwoom.setInputValue("기준일자", "20160624")
        kiwoom.setInputValue("수정주가구분", "1")

        kiwoom.commRqData("opt10081_req", "opt10081", 2, "0101")

    sys.exit(app.exec_())

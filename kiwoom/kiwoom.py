from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()
        print('kiwoom class initiated...')

        # Collection of variables to run event_loop
        self.login_event_loop = QEventLoop()

        # Variables of account
        self.account_num = None
        self.deposit = 0                                # 예수금
        self.use_money = 0                              # 실제 투자에 사용될 금액
        self.use_money_rate = 0.5                       # 에수금에서 실제 사용할 금액의 비율
        self.output_deposit = 0                         # 출력가능금액

        # Activate initial setting functions
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")               # 로그인 요청 시그널
        self.login_event_loop.exec_()

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)  # 트랜젝션 요청관련

    def login_slot(self, err_code):
        print(errors(err_code)[1])
        self.login_event_loop.exit()                    # 로그인 처리 완료시 루프 종료

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(Qstring)", "ACCNO")
        account_num = account_list.split(';')

        self.account_num = account_num

        print('Account Number: %s' % account_num)

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("SetInputValue(QString, QString, int, QString)",
                         "예수금상세현황요청", "opw00001", sPrevNext, self.screen_my_info)

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            use_money = float(self.deposit) * self.use_money_rate
            self.use_money = int(use_money)
            self.use_money = self.use_money / 4                             # 한 종목에 돈을 다쓰지 않게 4종목에 나눠서

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.output_deposit = int(output_deposit)

            print("예수금: $s" % self.output_deposit)

            self.stop_screen_cancel(self.screen_my_info)
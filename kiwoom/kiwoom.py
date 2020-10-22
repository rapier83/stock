from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()
        print('kiwoom class initiated...')

        # Collection of variables to run event_loop
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()

        # Variables of account
        self.account_num = None
        self.deposit = 0                                # 예수금
        self.use_money = 0                              # 실제 투자에 사용될 금액
        self.use_money_rate = 0.5                       # 에수금에서 실제 사용할 금액의 비율
        self.output_deposit = 0                         # 출력가능금액

        # Requested Screen Number
        self.screen_my_info = "2000"

        # Activate initial setting functions
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()                         # 계좌번호
        self.detail_account_info()                      # 예수금
        self.detail_account_mystock()                   # 계좌평가잔고내역

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
        account_num = account_list.split(';')[0]

        self.account_num = account_num

        print('Account Number: %s' % account_num)

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            use_money = float(self.deposit) * self.use_money_rate
            self.use_money = int(use_money)
            self.use_money = self.use_money / 4                             # 한 종목에 돈을 다쓰지 않게 4종목에 나눠서

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.output_deposit = int(output_deposit)

            print(f'예수금: {self.output_deposit:,} 원')

            # self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money =           self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            self.total_buy_money = int(total_buy_money)
            total_profit_loss_money =   self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            # total_profit_loss_rate =    self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익율(%)")
            # self.total_profit_loss_rate = float(total_profit_loss_rate)

            print(f'계좌평가잔고내역요청 싱글데이터: {self.total_buy_money:,} - {self.total_profit_loss_money:,}')

            self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

    def stop_screen_my_info(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QSting)", sScrNo)



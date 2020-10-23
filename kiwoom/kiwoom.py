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

        # Variables of account setting
        self.account_num = None                         # 계좌번호
        self.deposit = 0                                # 예수금
        self.use_money = 0                              # 실제 투자에 사용될 금액
        self.use_money_rate = 0.5                       # 에수금에서 실제 사용할 금액의 비율
        self.output_deposit = 0                         # 출력가능금액
        self.account_stock_dict = {}                    # 보유주식 딕셔너리

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
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")   # 추정조회는 1 일반조회는 2
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opw00001_req", "opw00001", sPrevNext, self.screen_my_info)    # 예수금상세조회

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opw00018", "opw00018", sPrevNext, self.screen_my_info)        # 계좌평가잔고내역

        self.detail_account_info_event_loop = QEventLoop()
        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "opw00001_req":
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

        elif sRQName == "opw00018":
            total_buy_money =           self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                         sTrCode, sRQName, 0, "총매입금액")
            self.total_buy_money = int(total_buy_money)
            total_profit_loss_money =   self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                         sTrCode, sRQName, 0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            total_profit_loss_rate =    self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                         sTrCode, sRQName, 0, "총수익률(%)")
            self.total_profit_loss_rate = float(total_profit_loss_rate)

            print(f'계좌평가잔고내역요청 싱글데이터:'
                  f' {self.total_buy_money:,} - {self.total_profit_loss_money:,} - {self.total_profit_loss_rate}')

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code =              self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                code_nm =           self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "종목명")
                holding_quantity =  self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "보유수량")
                buy_price =         self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                        sTrCode, sRQName, i, "매입가")             #매입 평균 가격
                learn_rate =        self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "수익률(%)")
                current_price =     self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "현재가")
                total_buy_price =   self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "매입금액")
                sell_amount =       self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                     sTrCode, sRQName, i, "매매가능수량")

                # print(f'종목번호 {code} - {code_nm} - {holding_quantity} 주 - {buy_price} 원에 매입 - 수익률 {learn_rate} - 현재 {current_price}원')

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}

                code_nm = code_nm.strip()
                holding_quantity = int(holding_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_buy_price = int(total_buy_price.strip())
                sell_amount = int(sell_amount.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": holding_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_buy_price})
                self.account_stock_dict[code].update({"매매가능수량": sell_amount})

                print(self.account_stock_dict[code])

            print(f'sPreNext: {sPrevNext}')
            print(f'The number of shares in your account: {rows}')

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

    def stop_screen_my_info(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QSting)", sScrNo)

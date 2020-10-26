from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.errorCode import *


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()
        print('kiwoom class initiated...')

        # Collection of variables to run event_loop
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()
        self.calc_event_loop = QEventLoop()

        # Variables of account setting
        self.account_num = None  # 계좌번호
        self.deposit = 0  # 예수금
        self.use_money = 0  # 실제 투자에 사용될 금액
        self.use_money_rate = 0.5  # 에수금에서 실제 사용할 금액의 비율
        self.output_deposit = 0  # 출력가능금액
        self.account_stock_dict = {}  # 보유주식 딕셔너리
        self.not_account_stock_dict = {}  # 미체결

        self.calc_data = []  # 종목분석용

        # Requested Screen Number
        self.screen_my_info = "2000"  # 조회용 스크린 번호
        self.screen_calc_stock = "4000"  # 계산용 스크린 번호

        # Activate initial setting functions
        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()  # 계좌번호
        self.detail_account_info()  # 예수금
        self.detail_account_mystock()  # 계좌평가잔고내역
        QTimer.singleShot(5000, self.not_concluded_account)  # 5초 뒤 미체결

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)  # 트랜젝션 요청관련

    def login_slot(self, err_code):
        print(errors(err_code)[1])
        self.login_event_loop.exit()  # 로그인 처리 완료시 루프 종료

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(Qstring)", "ACCNO")
        account_num = account_list.split(';')[0]

        self.account_num = account_num

        print('Account Number: %s' % account_num)

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")  # 추정조회는 1 일반조회는 2
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opw00001_req", "opw00001", sPrevNext, self.screen_my_info)  # 예수금상세조회

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opw00018_req", "opw00018", sPrevNext, self.screen_my_info)  # 계좌평가잔고내역

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext="0"):
        print("Requesting Not concluded Item list...")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opt10075_req", "opt10075", sPrevNext, self.screen_my_info)  # 실시간미체결요청

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        global price_top_moving
        if sRQName == "opw00001_req":
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            use_money = float(self.deposit) * self.use_money_rate
            self.use_money = int(use_money)
            self.use_money = self.use_money / 4  # 한 종목에 돈을 다쓰지 않게 4종목에 나눠서

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,
                                              "출금가능금액")
            self.output_deposit = int(output_deposit)

            print(f'예수금: {self.output_deposit:,} 원')

            self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

        elif sRQName == "opw00018_req":
            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                               sTrCode, sRQName, 0, "총매입금액")
            self.total_buy_money = int(total_buy_money)
            total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                       sTrCode, sRQName, 0, "총평가손익금액")
            self.total_profit_loss_money = int(total_profit_loss_money)
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                                      sTrCode, sRQName, 0, "총수익률(%)")
            self.total_profit_loss_rate = float(total_profit_loss_rate)

            print(f'계좌평가잔고내역요청 싱글데이터:'
                  f' {self.total_buy_money:,} - {self.total_profit_loss_money:,} - {self.total_profit_loss_rate}')

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                        sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                holding_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                    "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                             "매입가")  # 매입 평균 가격
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "현재가")
                total_buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                   "매입금액")
                sell_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "매매가능수량")

                # print(f'종목번호 {code} - {code_nm} - {holding_quantity} 주 - '
                #       f'{buy_price} 원에 매입 - 수익률 {learn_rate} - 현재 {current_price}원')
                print(f'{code}'
                      f' - {code_nm}'
                      f' - {holding_quantity} 주'
                      f' - 매입가 {buy_price}'
                      f' - 수익률 {learn_rate}'
                      f' - 현재가 {current_price}'
                      f' - 매입금액 {total_buy_price}'
                      f' - 매매가능수량 {sell_amount}')

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

                # print(self.account_stock_dict[code])

            print(f'sPrevNext: {sPrevNext}')
            print(f'The number of shares in your account: {rows}')

            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName == "opt10075_req":
            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목코드")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                "주문상태")
                not_concluded = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                               "주문가격")
                order_type = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "주문구분")
                amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                confirmed_amount = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                    "체결량")

                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                not_concluded = int(not_concluded.strip())
                order_price = int(order_price.strip())
                order_type = order_type.strip().lstrip('+').lstrip('-')
                amount = int(amount.strip())
                confirmed_amount = int(confirmed_amount)

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                self.not_account_stock_dict[order_no].update({'종목코드': code})
                self.not_account_stock_dict[order_no].update({'종목명': code_nm})
                self.not_account_stock_dict[order_no].update({'주문번호': order_no})
                self.not_account_stock_dict[order_no].update({'주문상태': order_status})
                self.not_account_stock_dict[order_no].update({'주문수량': not_concluded})
                self.not_account_stock_dict[order_no].update({'주문가격': order_price})
                self.not_account_stock_dict[order_no].update({'주문구분': order_type})
                self.not_account_stock_dict[order_no].update({'미체결수량': amount})
                self.not_account_stock_dict[order_no].update({'체결량': confirmed_amount})

                print(f'미체결 종목: {self.not_account_stock_dict[order_no]}')

        elif sRQName == 'opt10081_req':
            code = self.dynamicCall('GetCommData(QString, QString, int, QString)', sTrCode, sRQName, 0, '종목코드')
            code = code.strip()
            # data = self.dynamicCall('GetCommDataEx(QString, QString, itn, QString)', sTrCode, sRQName)
            # [[‘’, ‘현재가’, ‘거래량’, ‘거래대금’, ‘날짜’, ‘시가’, ‘고가’, ‘저가’. ‘’], [‘’, ‘현재가’, ’거래량’,
            # ‘거래대금’, ‘날짜’, ‘시가’, ‘고가’, ‘저가’, ‘’]. […]]
            cnt = self.dynamicCall('GetRepeatCnt(QString, QString)', sTrCode, sRQName)
            print(f'Remaining Days {cnt} Day(s)')

            for i in range(cnt):
                data = []

                current_price = self.dynamicCall('GetCommData(QString, QString, int, QString',
                                                  sTrCode, sRQName, i, '현재가')
                value = self.dynamicCall('GetCommData(QString, QString, int, QString', sTrCode, sRQName, i, '거래량')
                trading_value = self.dynamicCall('GetCommData(QString, QString, int, QString)',
                                                  sTrCode, sRQName, i, '거래대금')
                date = self.dynamicCall('GetCommData(QString, QString, int, QString)',
                                                  sTrCode, sRQName, i, '일자')
                start_price = self.dynamicCall('GetCommData(QString, QString, int, QString', sTrCode, sRQName, i, '시가')
                high_price = self.dynamicCall('GetCommData(QString, QString, int, QString', sTrCode, sRQName, i, '고가')
                low_price = self.dynamicCall('GetCommData(QString, QString, int, QString', sTrCode, sRQName, i, '저가')

                data.append('')
                data.append(current_price)
                data.append(value.strip())
                data.append(trading_value.strip())
                data.append(date.strip())
                data.append(start_price.strip())
                data.append(high_price.strip())
                data.append(low_price.strip())
                data.append('')

                self.calc_data.append(data.copy())

            if sPrevNext == '2':
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                print(f'Total days {len(self.calc_data)}')

                pass_success = False

                if self.calc_data is None or len(self.calc_data) < 120:
                    pass_success = False
                else:
                    total_price = 0
                    for value in self.calc_data[:120]:
                        total_price += int(value[1])
                    moving_avg_price = total_price / 120

                    # 오늘자 주가가 120일 이평선에 걸쳐있는지 확인
                    bottom_stock_price = False
                    check_price = None
                    if int(self.calc_data[0][7]) <= moving_avg_price <= int(self.calc_data[0][6]):
                        print('Confirmed today\'s price is over 120 days-avg.-line')
                        bottom_stock_price = True
                        check_price = int(self.calc_data[0][6])

                    # 걸쳐있지 않으면 그런 구간이 있었는지 확인
                    prev_price = None
                    if bottom_stock_price is True:
                        moving_avg_price_prev = 0
                        price_top_moving = False
                        idx = 1

                        while True:
                            if len(self.calc_data[idx:]) < 120:
                                print('There\'s no 120 days amount of data.')
                                break

                            total_price = 0
                            for value in self.calc_data[idx:120+idx]:
                                total_price += int(value[1])
                            moving_avg_price_prev = total_price / 120   # 120일 평균을 계산하고...

                            if moving_avg_price_prev <= int(self.calc_data[idx][6]) and idx <= 20:
                                print('If for 20days price was equal with 120-avg.-line or above then it\'s false')
                                price_top_moving = False
                                break   # 20일 이내 시가가 이평선 위에 있으므로 루프 탈출

                            elif int(self.calc_data[idx][7]) > moving_avg_price_prev and idx > 20:
                                print('Confirmed bound above 120-avg.-line ')
                                price_top_moving = True
                                prev_price = int(self.calc_data[idx][7])
                                break   # 20일을 초과 고가가 이평선 위에 있으므로 루프 탈출

                            idx += 1

                        if price_top_moving is True:    # 고가가 변경된 신호가 있으면
                            if moving_avg_price > moving_avg_price_prev and check_price > prev_price:
                                print('Confirmed that found 120-avg.-line price is lower than today\'s avg.-line.')
                                print('Confirmed that the bottom of found candle bar is lower than today\'s top price')
                                pass_success = True

                if pass_success is True:
                    print('Conditionally Pass')

                    code_nm = self.dynamicCall('GetMasterCodeName(QString)', code)

                    f = open('files/condition_sotck.txt', 'a', encoding=utf8)
                    f.write(f'{code}\t{code_nm}\t{str(self.calc_data[0][1])}\n')
                    f.close()

                elif pass_success is False:
                    print('Fail (Conditional)')

                self.calc_data.clear()
                self.detail_account_info_event_loop.exit()

    def stop_screen_cancel(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QSting)", sScrNo)

    def get_code_list_by_market(self, market_code):
        code_list = self.dynamicCall('GetCodeListByMarket(QString)', market_code)
        code_list = code_list.split(';')[:-1]
        return code_list

    def calc_fnc(self):
        code_list = self.get_code_list_by_matket("10")

        print(f'The number of KOSDAQ shares: {len(code_list)}')

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calc_stock)
            print(f'{idx + 1} / {len(code_list)} KOSDAQ Stock Code : {code} is updating......')

    def day_kiwoom_db(self, code=None, date=None, sPrevNext='0'):
        QTest.qWait(4000)

        self.dynamicCall('SetInputValue(QString, QString)', '종목코드', code)
        self.dynamicCall('SetInputValue(QString, QString)', '수정주가구분', '1')

        if date is not None:
            self.dynamicCall('SetInputValue(QString, QString)', '기준일자', date)

        self.dynamicCall('CommRqData(QString, QString, int, QString',
                         'opt10081_req', 'opt10081', sPrevNext, self.screen_calc_stock)

        self.calc_event_loop.exec_()

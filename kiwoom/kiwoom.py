import os
import sys
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
from config.errorCode import *
from config.kiwoomType import *
from config.log_class import *
from config.slack import *


class Kiwoom(QAxWidget):

    realType: RealType

    def __init__(self):
        super().__init__()

        self.realType = RealType()  # FID 번호 할당
        self.logging = Logging()
        self.slack = Slack()

        # print('kiwoom class initiated...')
        self.logging.logger.debug('Kiwoom class initiated...')

        # Collection of variables to run event_loop
        self.login_event_loop = QEventLoop()
        self.detail_account_info_event_loop = QEventLoop()
        self.calc_event_loop = QEventLoop()

        self.all_stock_dict = {}  # 전체 보유 종목 딕셔너리

        # Variables of account setting
        self.account_num = None  # 계좌번호
        self.deposit = 0  # 예수금
        self.use_money = 0  # 실제 투자에 사용될 금액
        self.use_money_rate = 0.5  # 에수금에서 실제 사용할 금액의 비율
        self.output_deposit = 0  # 출력가능금액
        self.total_profit_loss_money = 0    # 총 평가손익
        self.total_profit_loss_rate = 0.0    # 총 평가수익률

        # 종목정보 불러오기
        self.account_stock_dict = {}  # 보유주식 딕셔너리
        self.pending_dict = {}  # 미체결
        self.portfolio_stock_dict = {}  # 포트폴리오
        self.balance_dict = {}  # 잔고

        self.calc_data = []  # 종목분석용

        # Requested Screen Number
        self.screen_start_stop_real = '1000'  # 장 시작/종료 스크린 번호
        self.screen_my_info = '2000'  # 조회용 스크린 번호
        self.screen_calc_stock = '4000'  # 계산용 스크린 번호
        self.screen_real_stock = '5000'
        self.screen_trading_stock = '6000'

        # Activate initial setting functions
        self.get_ocx_instance()
        self.event_slots()
        self.real_event_slots()
        self.signal_login_commConnect()
        self.get_account_info()  # 계좌번호
        self.detail_account_info()  # 예수금
        self.detail_account_mystock()  # 계좌평가잔고내역
        QTimer.singleShot(5000, self.pending_account)  # 5초 뒤 미체결

        QTest.qWait(10000)
        self.read_code()
        self.screen_number_setting()

        QTest.qWait(5000)

        # 실시간 수신 관련 함수
        self.dynamicCall(
            'SetRealReg(QString, QString, QString, QString)',
            self.screen_start_stop_real, '', self.realType.REALTYPE['장시작시간']['장운영구분'], '0')

        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall('SetRealReg(QString, QString, QString, QString)', screen_num, code, fids, '1')

        self.slack.notification(
            pretext='주식 자동화 프로그램 동작했습니다.',
            title='주식 자동화 프로그램 동작',
            fallback='주식자동화프로그램동작',
            text='주식 자동화 프로그램이 동작했습니다.'
        )

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)  # 트랜젝션 요청관련
        self.OnReceiveMsg.connect(self.msg_slot)

    def real_event_slots(self):
        self.OnReceiveTrData.connect(self.realdata_slot)  # 실시간 이벤트 연결
        self.OnReceiveChejanData.connect(self.chejan_slot)  # 종목 주식체결 관련한 이벤트

    def login_slot(self, err_code):
        print(errors(err_code)[1])
        self.login_event_loop.exit()  # 로그인 처리 완료시 루프 종료

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        account_num = account_list.split(';')[0]

        self.account_num = account_num

        print(f'Account Number: {account_num}' )

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

    def pending_account(self, sPrevNext="0"):
        print("Requesting Pending Items list...")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "opt10075_req", "opt10075", sPrevNext, self.screen_my_info)  # 실시간미체결요청

        self.detail_account_info_event_loop.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

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

                if order_no in self.pending_dict:
                    pass
                else:
                    self.pending_dict[order_no] = {}

                self.pending_dict[order_no].update({'종목코드': code})
                self.pending_dict[order_no].update({'종목명': code_nm})
                self.pending_dict[order_no].update({'주문번호': order_no})
                self.pending_dict[order_no].update({'주문상태': order_status})
                self.pending_dict[order_no].update({'주문수량': not_concluded})
                self.pending_dict[order_no].update({'주문가격': order_price})
                self.pending_dict[order_no].update({'주문구분': order_type})
                self.pending_dict[order_no].update({'미체결수량': amount})
                self.pending_dict[order_no].update({'체결량': confirmed_amount})

                print(f'미체결 종목: {self.pending_dict[order_no]}')

            self.detail_account_info_event_loop.exit()

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

                # 120일 만큼의 데이터가 있는지 확인
                if self.calc_data is None or len(self.calc_data) < 120:
                    pass_success = False

                else:

                    # 120일 이평선 최근가격
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
                            for value in self.calc_data[idx:120 + idx]:
                                total_price += int(value[1])
                            moving_avg_price_prev = total_price / 120  # 120일 평균을 계산하고...

                            if moving_avg_price_prev <= int(self.calc_data[idx][6]) and idx <= 20:
                                print('If for 20days price was equal with 120-avg.-line or above then it\'s false')
                                price_top_moving = False
                                break  # 20일 이내 시가가 이평선 위에 있으므로 루프 탈출

                            elif int(self.calc_data[idx][7]) > moving_avg_price_prev and idx > 20:
                                print('Confirmed bound above 120-avg.-line ')
                                price_top_moving = True
                                prev_price = int(self.calc_data[idx][7])
                                break  # 20일을 초과 고가가 이평선 위에 있으므로 루프 탈출

                            idx += 1

                        if price_top_moving is True:  # 고가가 변경된 신호가 있으면
                            if moving_avg_price > moving_avg_price_prev and check_price > prev_price:
                                print('Confirmed that found 120-avg.-line price is lower than today\'s avg.-line.')
                                print('Confirmed that the bottom of found candle bar is lower than today\'s top price')
                                pass_success = True

                if pass_success is True:
                    print('Conditionally Pass')

                    code_nm = self.dynamicCall('GetMasterCodeName(QString)', code)

                    f = open('files/condition_stock.txt', 'a', encoding='utf8')
                    f.write(f'{code}\t{code_nm}\t{str(self.calc_data[0][1])}\n')
                    f.close()

                elif pass_success is False:
                    print('Conditionally Fail')

                self.calc_data.clear()
                self.detail_account_info_event_loop.exit()

    def realdata_slot(self, sCode, sRealType, sRealData):
        if sRealType == '장시작시간':
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            # 0:장시작전 2: 장종료20분전 3:장시작 4,8:장종료30분전 9:장마감
            value = self.dynamicCall('GetCommRealData(QString, int)', sCode, fid)

            if value == '0':
                print(f'장 시작 전')
            elif value == '3':
                print(f'장 시작')
            elif value == '2':
                print(f'장 종료, 동시호가로 변경')
            elif value == '4':
                print('3시 30분 장 종료')

                for code in self.portfolio_stock_dict.keys():
                    self.dynamicCall(
                        "SetRealRemove(QString, QString)", self.portfolio_stock_dict[code]['스크린번호'], code)

                QTest.qWait(5000)

                self.file_delete()
                self.calc_fnc()

                sys.exit()

        elif sRealType == '주식체결':
            a = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['체결시간'])
            # HHMMSS
            b = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['현재가'])
            # +(-)NNNN
            b = abs(int(b))

            c = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['전일대비'])
            c = abs(int(c))

            d = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['등락율'])
            d = float(d)

            e = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['(최우선)매도호가'])
            e = abs(int(e))

            f = self.dynamicCall('GetCommRealData(QString, int)', sCode, self.realType.REALTYPE[sRealType]['(최우선)매수호가'])
            f = abs(int(f))

            g = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['거래량'])  # 출력 : +240124  매수일때, -2034 매도일 때
            g = abs(int(g))

            h = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['누적거래량'])  # 출력 : 240124
            h = abs(int(h))

            i = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['고가'])  # 출력 : +(-)2530
            i = abs(int(i))

            j = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['시가'])  # 출력 : +(-)2530
            j = abs(int(j))

            k = self.dynamicCall("GetCommRealData(QString, int)", sCode,
                                 self.realType.REALTYPE[sRealType]['저가'])  # 출력 : +(-)2530
            k = abs(int(k))

            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({sCode: {}})

            self.portfolio_stock_dict[sCode].update({"체결시간": a})
            self.portfolio_stock_dict[sCode].update({"현재가": b})
            self.portfolio_stock_dict[sCode].update({"전일대비": c})
            self.portfolio_stock_dict[sCode].update({"등락율": d})
            self.portfolio_stock_dict[sCode].update({"(최우선)매도호가": e})
            self.portfolio_stock_dict[sCode].update({"(최우선)매수호가": f})
            self.portfolio_stock_dict[sCode].update({"거래량": g})
            self.portfolio_stock_dict[sCode].update({"누적거래량": h})
            self.portfolio_stock_dict[sCode].update({"고가": i})
            self.portfolio_stock_dict[sCode].update({"시가": j})
            self.portfolio_stock_dict[sCode].update({"저가": k})

            if sCode in self.account_stock_dict.keys() and sCode not in self.balance_dict.keys():
                asd = self.account_stock_dict[sCode]
                trading_rate = (b - asd['매입가']) / asd['매입가'] * 100

                if asd['매매가능수량'] > 0 and (trading_rate > 5 or trading_rate < -5):
                    order_success = self.dynamicCall(
                        'SendOrder(''QString, QString, QString, int, QString, int, int,'' QString, QString)',
                        ['신규매도', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 2, sCode,
                         asd['매매가능수량'], 0 , self.realType.SENDTYPE['거래구분']['시장가'], '']
                    )

                    if order_success == 0:
                        print('Selling order transfer succeed.')
                    else:
                        print('Selling order transfer failed')

            elif sCode in self.balance_dict.keys():
                bal = self.balance_dict[sCode]
                trading_rate = (b - bal['매입단가']) / bal['매입단가'] * 100

                if bal['주문가능수량'] > 0 and (trading_rate > 5 or trading_rate < -5):
                    order_success = self.dynamicCall(
                        'SendOrder(''QString, QString, QString, int, QString, int, int,'' QString, QString)',
                        ['신규매도', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 2, sCode,
                         bal['주문가능수량'], 0, self.realType.SENDTYPE['거래구분']['시장가'], '']
                    )

                    if order_success == 0:
                        print(f'Selling order transfer succeed.')
                    else:
                        print(f'Selling order transfer failed.')

            elif d > 2.0 and sCode not in self.balance_dict:
                print(f'Buying Condition passed {sCode}')

                result = (self.use_money * 0.1) / e
                quantity = int(result)

                order_success = self.dynamicCall(
                    'SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)',
                    ['신규매수', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 1, sCode,
                     quantity, e, self.realType.SENDTYPE['거래구분']['지정자'], '']
                )

                if order_success == 0:
                    print('Buying Order transfer succeed')
                else:
                    print('Buying Order transfer failed')

            pending_list = list(self.pending_dict)
            for order_num in pending_list:
                code = self.pending_dict[order_num]['종목코드']
                order_price = self.pending_dict[order_num]['주문가격']
                pending_quantity = self.pending_dict[order_num]['미체결수량']
                order_type = self.pending_dict[order_num]['주문구분']

                if order_type == '매수' and pending_quantity > 0 and e > order_price:
                    # 매수주문인지 수량이 0보다 큰지 최우선매도호가보다 높은지
                    order_success = self.dynamicCall('SendOrder('
                                                     'QString, QString, QString, int, QString, int, int, QString, '
                                                     'QString)',
                                                     ['매수취소',
                                                      self.portfolio_stock_dict[sCode]['주문용스크린번호'],
                                                      self.account_num,
                                                      3,
                                                      code,
                                                      0,
                                                      0,
                                                      self.realType.SENDTYPE['거래구분']['지정가'],
                                                      order_num])

                    if order_success == 0:
                        print('Order Cancellation transfer succeed')
                    else:
                        print('Order Cancellation transfer failed')
                elif pending_quantity == 0:
                    del self.pending_dict[order_num]

    def chejan_slot(self, sGubun, nItemCnt, sFidList):
        
        if int(sGubun) == 0:    # 주문체결
            account_num = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목명'])
            stock_name = stock_name.strip()

            origin_order_number = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['원주문번호'])
            order_number = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문번호'])
            order_status = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문상태'])
            order_qty = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문수량'])
            order_qty = int(order_qty)
            order_price = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문가격'])
            pending_check_out_qty = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['미체결수량'])
            order_type = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문구분'])
            check_out_time = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['주문/체결시간'])
            check_out_price = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['체결가'])
            if check_out_price == '':
                check_out_price = 0
            else:
                check_out_price = int(check_out_price)
            check_out_qty = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['체결량'])
            if check_out_qty == '':
                check_out_qty = 0
            else:
                check_out_qty = int(check_out_qty)
            current_price = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['주문체결']['현재가'])
            current_price = abs(int(current_price))
            first_sell_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))
            first_buy_price = self.dynamicCall(
                'GetChejanData(int)', self.realType.REALTYPE['주문체결']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            if order_number not in self.pending_dict.keys():
                self.pending_dict.update({order_number: {}})

            self.pending_dict[order_number].update({"종목코드": sCode})
            self.pending_dict[order_number].update({"주문번호": order_number})
            self.pending_dict[order_number].update({"종목명": stock_name})
            self.pending_dict[order_number].update({"주문상태": order_status})
            self.pending_dict[order_number].update({"주문수량": order_qty})
            self.pending_dict[order_number].update({"주문가격": order_price})
            self.pending_dict[order_number].update({"미체결수량": pending_check_out_qty})
            self.pending_dict[order_number].update({"원주문번호": origin_order_number})
            self.pending_dict[order_number].update({"주문구분": order_type})
            self.pending_dict[order_number].update({"주문/체결시간": check_out_time})
            self.pending_dict[order_number].update({"체결가": check_out_price})
            self.pending_dict[order_number].update({"체결량": check_out_qty})
            self.pending_dict[order_number].update({"현재가": current_price})
            self.pending_dict[order_number].update({"(최우선)매도호가": first_sell_price})
            self.pending_dict[order_number].update({"(최우선)매수호가": first_buy_price})

        elif int(sGubun) == 1:  # 잔고
            account_num = self.dynamicCall('GetChejanData(int)', self.realType.REALTYPE['잔고']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목명'])
            stock_name = stock_name.strip()

            current_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['현재가'])
            current_price = abs(int(current_price))

            stock_qty = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['보유수량'])
            stock_qty = int(stock_qty)

            like_qty = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['주문가능수량'])
            like_qty = int(like_qty)

            buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매입단가'])
            buy_price = abs(int(buy_price))

            total_buy_price = self.dynamicCall(
                "GetChejanData(int)", self.realType.REALTYPE['잔고']['총매입가'])  # 계좌에 있는 종목의 총매입가
            total_buy_price = int(total_buy_price)

            check_out_type = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매도매수구분'])
            check_out_type = self.realType.REALTYPE['매도수구분'][check_out_type]

            first_sell_price = self.dynamicCall(
                "GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            if sCode not in self.balance_dict.keys():
                self.balance_dict.update({sCode: {}})

            self.balance_dict[sCode].update({"현재가": current_price})
            self.balance_dict[sCode].update({"종목코드": sCode})
            self.balance_dict[sCode].update({"종목명": stock_name})
            self.balance_dict[sCode].update({"보유수량": stock_qty})
            self.balance_dict[sCode].update({"주문가능수량": like_qty})
            self.balance_dict[sCode].update({"매입단가": buy_price})
            self.balance_dict[sCode].update({"총매입가": total_buy_price})
            self.balance_dict[sCode].update({"매도매수구분": check_out_type})
            self.balance_dict[sCode].update({"(최우선)매도호가": first_sell_price})
            self.balance_dict[sCode].update({"(최우선)매수호가": first_buy_price})

    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        print(f'스크린: {sScrNo},  요청이름: {sRQName},  tr코드: {sTrCode} --- {msg}')

    def file_delete(self):
        if os.path.isfile('files/condition_stock.txt'):
            os.remove('file/condition_stock.txt')

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

    def read_code(self):
        if os.path.exists('files/condition_stock.txt'):
            f = open('files/condition_stock.txt', 'r', encoding='utf8')

            lines = f.readline()
            for line in lines:
                ls = line.split('\t')

                stock_code = ls[0]
                stock_name = ls[1]
                stock_price = int(ls[2].split('\n')[0])
                stock_price = abs(stock_price)

                self.portfolio_stock_dict.update({stock_code: {'종목명': stock_name, '현재가': stock_price}})
            f.close()

    def merge_dict(self):
        self.all_stock_dict.update({'account': self.account_stock_dict})
        self.all_stock_dict.update({'pending': self.pending_dict})
        self.all_stock_dict.update({'portfolio': self.portfolio_stock_dict})

    def screen_number_setting(self):
        screen_overwrite = []

        # 계좌평가잔고(account)내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 미체결(pending)에 있는 종목들
        for order_number in self.pending_dict.keys():
            code = self.pending_dict[order_number]['종목코드']

            if code not in screen_overwrite:
                screen_overwrite.append(code)

        # 포트폴리오에 있는 종목들
        for code in self.portfolio_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)

        cnt = 0
        for code in screen_overwrite:
            temp_screen = int(self.screen_real_stock)
            trading_screen = int(self.screen_trading_stock)

            if (cnt % 50) == 0:
                temp_screen += 1
                trading_screen += 1
                self.screen_real_stock = str(temp_screen)
                self.screen_trading_stock = str(trading_screen)

            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({'스크린번호': str(self.screen_real_stock)})
                self.portfolio_stock_dict[code].update({'주문용스크린번호': str(self.screen_trading_stock)})

            elif code not in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict.update({code: {'스크린번호': str(self.screen_real_stock),
                                                         '주문용스크린번호': str(self.screen_trading_stock)}
                                                  })
            cnt += 1

        print(self.portfolio_stock_dict)

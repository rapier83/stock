from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import *


class Kiwoom(QAxWidget):

    def __init__(self):
        super().__init__()
        print('kiwoom class initiated...')

        # Collection of variables to run event_loop
        self.login_event_loop = QEventLoop()

        # Activate initial setting functions
        #self.kw = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        self.get_ocx_instance()
        self.event_slots()
        self.signal_login_commConnect()
        self.get_account_info()

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")               # 로그인 요청 시그널
        self.login_event_loop.exec_()

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)

    def login_slot(self, err_code):
        print(errors(err_code)[1])
        self.login_event_loop.exit()                    # 로그인 처리 완료시 루프 종료

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(Qstring)", "ACCNO")
        account_num = account_list.split(';')

        self.account_num = account_num

        print('Account Number: %s' % account_num)
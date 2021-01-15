from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop()  # 로그인 담당 이벤트 루프

        self.create_kiwoom_instance()
        self.login()

    # COM 오브젝트 생성.
    def create_kiwoom_instance(self):
        # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def login(self):
        self.OnEventConnect.connect(self.login_slot)  # 이벤트와 슬롯을 메모리에 먼저 생성.
        self.dynamicCall("CommConnect()")  # 시그널 함수 호출.
        self.login_event_loop.exec_()

    def login_slot(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print("로그인 실패 - 에러 내용 :", errors(err_code)[1])
        self.login_event_loop.exit()

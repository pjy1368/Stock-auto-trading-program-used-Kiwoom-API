import sys
import os
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop()  # 로그인 담당 이벤트 루프

        # 계좌 관련 변수
        self.account_number = None

        # 초기 작업
        self.create_kiwoom_instance()
        self.login()
        input()
        self.get_account_info()

        self.menu()

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
            print("로그인에 성공하였습니다.")
        else:
            if err_code == -106:  # 사용자가 강제로 키움api 프로그램을 종료하였을 경우
                os.system('cls')
                print(errors(err_code)[1])
                sys.exit(0)
            os.system('cls')
            print("로그인에 실패하였습니다.")
            print("에러 내용 :", errors(err_code)[1])
            sys.exit(0)
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        account_number = account_list.split(';')[0]
        self.account_number = account_number

    def menu(self):
        sel = ""
        while True:
            os.system('cls')
            print("1. 현재 로그인 상태 확인")
            print("2. 개인 정보 조회")
            print("Q. 프로그램 종료")
            sel = input("=> ")

            if sel == "Q" or sel == "q":
                sys.exit(0)

            if sel == "1":
                self.print_login_connect_state()
            elif sel == "2":
                self.print_my_info()

    def print_login_connect_state(self):
        isLogin = self.dynamicCall("GetConnectState()")
        if isLogin == 1:
            print("\n현재 계정은 로그인 상태입니다.\n")
        else:
            print("\n현재 계정은 로그아웃 상태입니다.\n")
        input()

    def print_my_info(self):
        user_name = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        user_id = self.dynamicCall("GetLoginInfo(QString)", "USER_ID")
        account_count = self.dynamicCall(
            "GetLoginInfo(QString)", "ACCOUNT_CNT")

        print(f"\n이름 : {user_name}")
        print(f"ID : {user_id}")
        print(f"보유 계좌 수 : {account_count}")
        print(f"1번째 계좌번호 : {self.account_number}\n")
        input()

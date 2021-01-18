import sys
import os
from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errCode import *


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.login_event_loop = QEventLoop()  # 로그인 담당 이벤트 루프
        self.get_detail_account_loop = QEventLoop()  # 상세 계좌 정보 얻어오기 담당 이벤트 루프

        # 계좌 관련 변수
        self.account_number = None
        self.deposit = 0
        self.withdraw_deposit = 0
        self.order_deposit = 0

        # 화면 번호
        self.screen_my_account = "1000"

        # 초기 작업
        self.create_kiwoom_instance()
        self.event_collection()  # 이벤트와 슬롯을 메모리에 먼저 생성.
        self.login()
        input()
        self.get_account_info()  # 계좌 번호만 얻어오기
        self.get_detail_account_info()  # 상세 계좌 정보 얻어오기

        self.menu()

    # COM 오브젝트 생성.
    def create_kiwoom_instance(self):
        # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

    def event_collection(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.tr_slot)  # 트랜잭션 요청 관련 이벤트

    def login(self):
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
            print("2. 사용자 정보 조회")
            print("3. 예수금 조회")
            print("Q. 프로그램 종료")
            sel = input("=> ")

            if sel == "Q" or sel == "q":
                sys.exit(0)

            if sel == "1":
                self.print_login_connect_state()
            elif sel == "2":
                self.print_my_info()
            elif sel == "3":
                self.print_detail_account_info()

    def print_login_connect_state(self):
        os.system('cls')
        isLogin = self.dynamicCall("GetConnectState()")
        if isLogin == 1:
            print("\n현재 계정은 로그인 상태입니다.\n")
        else:
            print("\n현재 계정은 로그아웃 상태입니다.\n")
        input()

    def print_my_info(self):
        os.system('cls')
        user_name = self.dynamicCall("GetLoginInfo(QString)", "USER_NAME")
        user_id = self.dynamicCall("GetLoginInfo(QString)", "USER_ID")
        account_count = self.dynamicCall(
            "GetLoginInfo(QString)", "ACCOUNT_CNT")

        print(f"\n이름 : {user_name}")
        print(f"ID : {user_id}")
        print(f"보유 계좌 수 : {account_count}")
        print(f"1번째 계좌번호 : {self.account_number}\n")
        input()

    def print_detail_account_info(self):
        os.system('cls')
        print(f"예수금 : {self.deposit}")
        print(f"출금 가능 금액 : {self.withdraw_deposit}")
        print(f"주문 가능 금액 : {self.order_deposit}")
        input()

    def get_detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)",
                         "계좌번호", self.account_number)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")
        self.dynamicCall("CommRqData(QString, QString, int, QString)",
                         "예수금상세현황요청", "opw00001", sPrevNext, self.screen_my_account)

        self.get_detail_account_loop.exec_()

    def tr_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        if sRQName == "예수금상세현황요청":
            deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            self.deposit = int(deposit)

            withdraw_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.withdraw_deposit = int(withdraw_deposit)

            order_deposit = self.dynamicCall(
                "GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.order_deposit = int(order_deposit)

            self.get_detail_account_loop.exit()

    def cancel_screen_number(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)

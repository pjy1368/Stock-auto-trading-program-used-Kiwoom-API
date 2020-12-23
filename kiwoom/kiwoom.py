from PyQt5.QAxContainer import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("키움 클래스입니다.")
        
        self.create_kiwoom_instance()

    # COM 오브젝트 생성.
    def create_kiwoom_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") # 레지스트리에 저장된 키움 openAPI 모듈 불러오기
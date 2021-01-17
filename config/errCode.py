def errors(err_code):
    err_dic = {0: ('OP_ERR_NONE', '정상처리'),
               -100: ('OP_ERR_LOGIN', '사용자정보교환실패'),
               -101: ('OP_ERR_CONNECT', '서버접속실패'),
               -102: ('OP_ERR_VERSION', '버전처리실패'),
               -106: ('OP_ERR_SOCKET_CLOSED', '통신연결종료')}
    result = err_dic[err_code]
    return result

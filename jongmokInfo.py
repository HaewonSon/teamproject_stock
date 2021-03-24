# 종목정보 구하는 예제 : 거래소, 코스닥

import win32com.client

# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    print("PLUS가 정상적으로 연결되지 않음. ")
    exit()

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥

print("거래소 종목코드", len(codeList))
for i, code in enumerate(codeList):

    # GetStockSectionKind : code에 해당하는 부구분코드를 반환한다.
    # code : 주식코드, 반환값 : 부구분코드
    secondCode = objCpCodeMgr.GetStockSectionKind(code)

    # GetStockStdPrice : code에 해당하는 권리락등으로 인한 기준가를 반환한다.
    # code : 주식코드 , 반환값 : 전일시가(LONG)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)

    # CodeToName : code에 해당하는 주식/선물/옵션종목명을 반환한다.
    # code : 주식/선물/옵션코드 , 반환값 : 주식/선물/옵션종목명
    name = objCpCodeMgr.CodeToName(code)

    print(i, code, secondCode, stdPrice, name)

print("코스닥 종목코드", len(codeList2))
for i, code in enumerate(codeList2):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    print(i, code, secondCode, stdPrice, name)

print("거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2))

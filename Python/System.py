import GlobalData
import EnvData
import Util
import pyautogui
import pyperclip
import subprocess
import openpyxl
import re
import ProductRegistration
import win32api
import webbrowser
from enum import Enum

class Enum_COLUMN(Enum):
    A = ("",) # 상품 번호 칸
    B = ("",) # 상품 url 칸
    C = ("",) # 상품 구매 url 칸
    E = ("",) # 브랜드 칸
    F = ("",) # 색 RGB(16진수) 리스트 칸
    G = ("",) # 색명(사이즈 리스트) 칸
    H = ("",) # 업데이트 시간 칸
    I = ("",) # 체크 시간 칸
    J = ("",) # 체크 상태 칸
    K = ("",) # 이전 색RGB(16진수) 리스트 칸
    L = ("",) # 이전 색명(사아즈 리스트) 칸
    O = ("",) 
    P = ("",) # 마지막 실행 시켰던 라인
    Q = ("",) # 마지막 실행 시켰던 라인의 시간 입력
    T = ("",) # 상품 이름
    U = ("",) # 상품 원본 가격
    
def SaveWorksheet(wb):
    # 원본 시트 선택
    ws_original = wb.active

    ws_copy = wb.copy_worksheet(ws_original)
    ws_copy.title = "복제 시트"  # 복제된 시트의 이름 설정

    # 원본 엑셀 파일 저장
    wb.save(EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.xlsx")

    # 복제된 엑셀 파일 저장
    wb_copy = openpyxl.Workbook()
    wb_copy.save(EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트_복제.xlsx")


def GetElementsData():
    Util.KeyboardKeyPress("f12")
    Util.SleepTime(2)
    if Util.WhileFoundImage(r"크롬\Elements에 html"):
        Util.MoveAtWhileFoundImage(r"크롬\Elements에 html", 5, 5)
        Util.SleepTime(0.5)
        Util.NowMouseClickRight()
        Util.SleepTime(3)
        currentPos = pyautogui.position()
        Util.MoveAtWhileFoundImage(r"크롬\Elements에 html의 copy", 5, 5, 10, 1, currentPos.x, currentPos.y)
        Util.SleepTime(1)
        Util.MoveAtWhileFoundImage(r"크롬\Elements에 html의 copy에 copy element", 5, 5, 10, 1,currentPos.x,currentPos.y)
        Util.SleepTime(1)
        Util.NowMouseClick()
        Util.SleepTime(3)
        outElementsData = pyperclip.paste()
        Util.SleepTime(0.5)
        return outElementsData
    return ""


# 등록 된 상품 최신화
def UpdateStoreWithColorInformation(inputRow=-1):
    Util.TelegramSend("등록 된 상품 최신화 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.xlsx"
    wb = openpyxl.load_workbook(xlFile)
    ws = wb.active
    lastRow = ws.max_row

    if inputRow != -1:
        row = inputRow
    else:
        row = round(ws[f"{Enum_COLUMN.A.name}{1}"].value)

    krwUsd = Util.KRWUSD()
    krwEur = Util.KRWEUR()

    while True:
        row += 1

        if row > lastRow:
            break

        Util.Debug("row(" + row + ") / lastRow(" + lastRow + ")")
        if row % 10 == 0:
            Util.TelegramSend(f"__ row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}")
            
        if "품절 상태로 변경 완료" in ws[f"{Enum_COLUMN.J.name}{row}"].value:
            ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
            ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
            SaveWorksheet(wb)
            continue

        url = ws[f"{Enum_COLUMN.C.name}{row}"].value

        if "www.ugg.com" in url:
            Util.TelegramSend(f"www.ugg.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}")
            isUpdateProduct = UpdateProductInfo_UGG(wb, ws, url, row, krwUsd)
            if isUpdateProduct:
                ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
                ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
                SaveWorksheet(wb)
            else:
                row -= 1

            continue

        if "www.mytheresa.com" in url:
            Util.TelegramSend(f"www.mytheresa.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}")
            isUpdateProduct = UpdateProductInfoMoney_Mytheresa(wb, ws, url, row, krwEur)
            if isUpdateProduct:
                ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
                ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
                SaveWorksheet(wb)
            else:
                row -= 1

            continue

    # True를 전달하여 저장 여부 설정
    wb.save(xlFile)  # 저장
    wb.close()  # 파일 닫기

    Util.TelegramSend("등록 된 상품 최신화 -- 끝")


def UpdateStoreWithColorInformationMoney_Mytheresa():
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.xlsx"
    wb = openpyxl.load_workbook(xlFile)
    ws = wb.active
    lastRow = ws.max_row

    row = round(ws[f"{Enum_COLUMN.P.name}{"1"}"].value)

    krwEur = Util.KRWEUR()

    while True:
        row += 1

        if row > lastRow:
            break

        Util.Debug("row(" + row + ") / lastRow(" + lastRow + ")")
        if row % 10 == 0:
            Util.TelegramSend(f"__ row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}")
            
        # 웹 브라우저 열기 및 상품 url로 이동
        url = ws[f"{Enum_COLUMN.C.name}{row}"].value
        if "www.mytheresa.com" not in url:
            ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
            ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
            SaveWorksheet(wb)
            continue

        if "품절 상태로 변경 완료" in ws[f"{Enum_COLUMN.J.name}{row}"].value:
            ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
            ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
            SaveWorksheet(wb)
            continue

        Util.TelegramSend(f"row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}")
        
        isUpdateProduct = UpdateProductInfoMoney_Mytheresa(wb, ws, url, row, krwEur)
        if isUpdateProduct:
            ws[f"{Enum_COLUMN.P.name}{"1"}"].value = row
            ws[f"{Enum_COLUMN.Q.name}{"1"}"].value = Util.GetFormattedCurrentDateTime()
            SaveWorksheet(wb)
        else:
            row -= 1

    # True를 전달하여 저장 여부 설정
    wb.save(xlFile)  # 저장
    wb.close()  # 파일 닫기


def UpdateProductInfo_UGG(wb, ws, url, row, krwUsd):
    data = GetUggData(url, krwUsd)

    # UGG에 사이즈 정보로 정보 취합
    useMoney = data.useMoney

    # 이중 배열
    arraySizesAndImgUrls = data.arraySizesAndImgUrls

    # 기존 것과 같은지 비교(같으면 스마트 스토어에 하지 않기 위함)
    before_SaveColorList = ws[f"{Enum_COLUMN.F.name}{row}"].value
    Util.Debug("before_SaveColorList : " + before_SaveColorList)

    # 기존 색 이름 과 사아즈를 변수로 저장
    before_SaveColorNameDoubleArray = ws[f"{Enum_COLUMN.G.name}{row}"].value
    Util.Debug("before_SaveColorNameDoubleArray : " + before_SaveColorNameDoubleArray)

    # 색이름 리스트 값
    colorNames = []
    for item in arraySizesAndImgUrls:
        colorNames.append(item[1])
    str_saveColorList = Util.JoinArrayToString(colorNames)
    Util.Debug("str_saveColorList : " + str_saveColorList)

    # 색 이름 과 사아즈 리스트 값(이중 배열)
    str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
    Util.Debug("str_saveColorNameDoubleArray : " + str_saveColorNameDoubleArray)

    # 색이 없은 경우 자체가 연결 되지 않거나 물건 자체가 없어졌을 경우
    if str_saveColorNameDoubleArray == "" or useMoney == 0:
        # 스마트 스토어 수정 화면까지 이동
        managedata = ManageAndModifyProducts(ws, row)
        if managedata.isNoNetwork == True:
            return False

        if managedata.isNoProduct == True:
            xl_J_(wb, ws, row, "스토어에 상품이 없습니다.")
            return True

        # 품절
        SoldOut(wb, ws, row)
    else:
        if (before_SaveColorNameDoubleArray == str_saveColorNameDoubleArray and ws[f"{Enum_COLUMN.U.name}{row}"].value == useMoney):
            # 이전과 정보가 변함이 없을 경우(이전과 동일하다고 적고 다음으로 넘어감)
            xl_J_(wb, ws, row, "이전과 동일합니다.")
        else:
            # 이전과 달라졌음
            xl_J_(wb, ws, row, "이전과 동일하지 않아서 변경 하려고 합니다.")

            # 스마트 스토어 수정 화면까지 이동
            managedata = ManageAndModifyProducts(ws, row)
            if managedata.isNoNetwork == True:
                return False

            if managedata.isNoProduct == True:
                xl_J_(wb, ws, row, "스토어에 상품이 없습니다.")
                return True

            # 가격 변동이 있으면 변경
            if ws[f"{Enum_COLUMN.U.name}{row}"].value != useMoney:
                # 판매가 입력
                UpdateAndReturnSalePrice(data.korMony)

            if before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray:
                # 관세 부가 여부 체크
                customsDuty = useMoney >= 200

                # 옵션 엑셀 세팅
                Util.SetExcelOption(arraySizesAndImgUrls, customsDuty)

                xl_J_(wb,ws,row,"이전과 동일하지 않아서 변경 하려고 합니다.(옵션 엑셀 세팅 완료)",)
                
                # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
                UpdateOptionsFromExcel(customsDuty)

                # 색은 그대로 인 상태에서 사이즈 숫자만 바꿔서 상세 페이지 갱신 하지 않도록 처리
                if before_SaveColorList != str_saveColorList:
                    # HTML 으로 등록
                    SetHTML(arraySizesAndImgUrls)

            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
            # Util.SleepTime(5)
            # Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", 5, 5)
            # Util.SleepTime(1)

            if (before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray and ws[f"{Enum_COLUMN.U.name}{row}"].value != useMoney):
                ws[f"{Enum_COLUMN.U.name}{row}"].value = useMoney

                # 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
                if True:
                    # 색 이름 리스트 표시
                    ws[f"{Enum_COLUMN.F.name}{row}"].value = str_saveColorList
                    Util.Debug("str_saveColorList : " + str_saveColorList)
                    
                    # 색 이름과 사아즈 리스트 표시
                    ws[f"{Enum_COLUMN.G.name}{row}"].value = str_saveColorNameDoubleArray
                    Util.Debug("str_saveColorNameDoubleArray : " + str_saveColorNameDoubleArray)
                    
                    xl_J_(wb,ws,row,"변경 완료(이전과 동일하지 않아)(이전 값 등록 전)",True)
                    
                    # 이전 색 이름 리스트 표시
                    ws[f"{Enum_COLUMN.K.name}{row}"].value = before_SaveColorList
                    Util.Debug("before_SaveColorList : " + before_SaveColorList)

                    # 이전 색 이름과 사아즈 리스트 표시
                    ws[f"{Enum_COLUMN.L.name}{row}"].value = before_SaveColorNameDoubleArray
                    Util.Debug(f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}")
                    
                    xl_J_(wb, ws, row, "변경 완료(이전과 동일하지 않아)(가격과 사이즈)")
            else:
                # 가격 변동이 있으면 변경
                if ws[f"{Enum_COLUMN.U.name}{row}"].value != useMoney:
                    ws[f"{Enum_COLUMN.U.name}{row}"].value = useMoney

                    xl_J_(wb, ws, row, "변경 완료(가격만 변동)")

                if before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray:
                    # 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
                    if True:
                        # 색 이름 리스트 표시
                        ws[f"{Enum_COLUMN.F.name}{row}"].value = str_saveColorList
                        Util.Debug(f"str_saveColorList : {str_saveColorList}")
                        
                        # 색 이름과 사아즈 리스트 표시
                        ws[f"{Enum_COLUMN.G.name}{row}"].value = str_saveColorNameDoubleArray
                        Util.Debug(f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}")
                        
                        xl_J_(wb,ws,row,"변경 완료(이전과 동일하지 않아)(이전 값 등록 전)",True)
                        
                        # 이전 색 이름 리스트 표시
                        ws[f"{Enum_COLUMN.K.name}{row}"].value = before_SaveColorList
                        Util.Debug("before_SaveColorList : " + before_SaveColorList)

                        # 이전 색 이름과 사아즈 리스트 표시
                        ws[f"{Enum_COLUMN.L.name}{row}"].value = before_SaveColorNameDoubleArray
                        Util.Debug(f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}")

                        xl_J_(wb, ws, row, "변경 완료(이전과 동일하지 않아)")

    return True


def UpdateProductInfoMoney_Mytheresa(wb, ws, url, row, krwEur):
    data = GetMytheresaData(url, krwEur)

    # 스마트 스토어 수정 화면까지 이동
    managedata = ManageAndModifyProducts(ws, row)
    if managedata.isNoNetwork == True:
        return False

    if managedata.isNoProduct == True:
        xl_J_(wb, ws, row, "스토어에 상품이 없습니다.")
        return True

    if data.isSoldOut:
        # 품절
        SoldOut(wb, ws, row)
    else:
        if data.sizesLength == 0:

            useMoney = data.useMoney

            # 판매가 입력
            UpdateAndReturnSalePrice(data.korMony)

            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)

            if data.korMony != 0:
                xl_J_(wb, ws, row, "변경 완료(가격만 변동)")
            else:
                xl_J_(wb, ws, row, "가격이 0이 나왔습니다.")
        else:

            useMoney = data.useMoney
            arraySizesAndImgUrls = data.arraySizesAndImgUrls

            # 관세 부가 여부 체크
            customsDuty = useMoney >= 150

            # 옵션 엑셀 세팅
            Util.SetExcelOption(arraySizesAndImgUrls, customsDuty)

            # 1. 가격 세팅
            # 2. 엑셀로 옵셥 세팅
            if True:
                # 판매가 입력
                UpdateAndReturnSalePrice(data.korMony)

                # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
                UpdateOptionsFromExcel(customsDuty)

                Util.SleepTime(1)
                Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
                # Util.SleepTime(5)
                # Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", 5, 5)
                # Util.SleepTime(1)

            if data.korMony != 0:
                xl_J_(wb, ws, row, "변경 완료(가격과 사이즈 변동)")
            else:
                xl_J_(wb, ws, row, "가격이 0이 나왔습니다.")

    return True


def xl_J_(wb, ws, row, value, updateTime=False):
    if updateTime:
        # 갱신 시간 표시
        ws[f"{Enum_COLUMN.H.name}{row}"].value = Util.GetFormattedCurrentDateTime()
    # 체크 시간 표시
    ws[f"{Enum_COLUMN.I.name}{row}"].value = Util.GetFormattedCurrentDateTime()
    # 체크 상태 표시
    ws[f"{Enum_COLUMN.J.name}{row}"].value = value

    SaveWorksheet(wb)


# UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
def GetNewProductURLs_UGG(name, url, filterUrls):
    Util.TelegramSend("GetNewProductURLs_UGG() " + name + " -- 시작")
    webbrowser.open(url)
    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
    # WinWait, ugg
    Util.SleepTime(10)

    # 웹 제일 끝까지 스코롤 한다.
    while True:
        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        # -10000 틱 스크롤 다운
        Util.MouseWheelScroll(-10000)
        Util.SleepTime(1.5)
        Util.KeyboardKeyPress("up")
        Util.SleepTime(1)
        Util.KeyboardKeyPress("down")
        Util.KeyboardKeyPress("down")
        Util.SleepTime(1)
        # 화면 가로 및 세로 해상도 얻기
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        if (Util.ClickAtWhileFoundImage(r"크롬\오른쪽 스트롤바가 제일 아래인 이미지",0,0,1,1,screen_width - 200,screen_height - 200) or 
            Util.ClickAtWhileFoundImage(r"크롬\오른쪽 스트롤바가 제일 아래인 이미지_v2",0,0,1,1,screen_width - 200,screen_height - 200) or 
            Util.ClickAtWhileFoundImage(r"크롬\오른쪽 스트롤바가 제일 아래인 이미지_v3",0,0,1,1,screen_width - 200,screen_height - 200)):
            # 상품 더 보기가 있는지 체크
            if Util.ClickAtWhileFoundImage(r"UGG\상품 리스트\상품 더 보기 버튼", 0, 0, 3):
                Util.SleepTime(5)
            else:  # 상품이 더이상 없음
                break

    htmlElementsData = GetElementsData()
    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    productUrls = []
    # <a href=" 과 " class="js-pdp-link image-link pdp-link"> 중간에 있는 값
    productUrlLines = Util.GetRegExMatche1List(
        htmlElementsData, r'<a href="(.*?)" class="js-pdp-link image-link pdp-link">'
    )
    for productUrlLine in productUrlLines:
        splitList = productUrlLine.split(".html")
        if len(splitList) > 0:
            productUrls.append("https://www.ugg.com" + splitList[1] + ".html")
        else:
            productUrls.append(productUrlLine)
    uniqueArr = []
    for productUrl in productUrls:
        for filterUrl in filterUrls:
            if productUrl == filterUrl:
                uniqueArr.append(productUrl)
                break

    for uniqueValue in uniqueArr:
        ArrayRemove(productUrls, uniqueValue)

    Util.TelegramSend("GetNewProductURLs_UGG() " + name + " -- 끝")

    return [name, productUrls]


def ArrayRemove(arr, value):
    for index, element in enumerate(arr):
        if element == value:
            arr.pop(index)
            break


# 신규 등록 할 UGG 목록을 엑셀에 정리
def SetXlsxUGGNewProductURLs():
    Util.TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.xlsx"
    wb = openpyxl.load_workbook(xlFile)
    ws = wb.active
    lastRow = ws.max_row

    Util.Debug("start xlsx ugg url")
    # C 열의 데이터를 배열에 저장
    filterUrls = []
    for row_index in range(1, lastRow + 1):
        url = ws.cell(row=row_index, column=3).value
        if url is not None and "www.ugg.com" in url:
            filterUrls.append(url)
    Util.TelegramSend("end xlsx ugg url Length : " + str(len(filterUrls)))

    wb.save(xlFile)  # 저장
    wb.close()  # 파일 닫기

    # 메뉴 창이 한번은 열려야지 세부 메뉴 창이 정상으로 열림
    webbrowser.open("https://www.ugg.com/women-footwear")
    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
    # WinWait, ugg
    Util.SleepTime(10)

    # UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
    uggProductUrls = []
    uggProductUrls.append(GetNewProductURLs_UGG("패션잡화 여성신발 부츠 미들부츠","https://www.ugg.com/women-footwear/?prefn1=type&prefv1=boots%7Cclassic-boots%7Ccold-weather-boots",filterUrls,))  # 부츠(미들부츠)
    uggProductUrls.append(GetNewProductURLs_UGG("패션잡화 여성신발 샌들 뮬","https://www.ugg.com/women-footwear/?prefn1=type&prefv1=dress-shoes%7Csandals",filterUrls,))  # 샌들(뮬)
    uggProductUrls.append(GetNewProductURLs_UGG("패션잡화 여성신발 슬리퍼","https://www.ugg.com/women-footwear/?prefn1=type&prefv1=clogs%7Cslippers",filterUrls))  # 슬리퍼
    uggProductUrls.append(GetNewProductURLs_UGG("패션잡화 여성신발 운동화 러닝화","https://www.ugg.com/women-footwear/?prefn1=type&prefv1=sneakers",filterUrls,))  # 운동화 

    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    xlFile = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.xlsx"
    wb = openpyxl.load_workbook(xlFile)
    ws = wb.active
    rowCount = ws.max_row

    # 모든 행을 삭제합니다.
    for _ in range(rowCount):
        ws(1).Delete()  # 각 반복에서 첫 번째 행을 삭제합니다.

    allCount = 0
    for item in uggProductUrls:
        for item2 in item[2]:
            allCount += 1
            ws("A" + allCount).value = "UGG"
            ws("B" + allCount).value = item[1]  # 메뉴
            ws("C" + allCount).value = item2  # url

    wb.save(xlFile)  # 저장
    wb.close()  # 파일 닫기

    Util.TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 끝")


# HTML 으로 등록
def SetHTML(arraySizesAndImgUrls, isAdd=False):
    # 상세설명 찾아서 그 아래로 ONE 원형 검색이 존재 하는데 체크(이때는 아래로 조금씩 내리기)
    Util.WheelAndMoveAtWhileFoundImage(r"스마트 스토어\상품 수정\상세 설명")
    findIndex = 0
    while True:
        if Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\녹색 상세설명", 0, 0, 1):
            findIndex = 1
            break
        else:
            if Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\녹색 상세설명_v2", 0, 0, 1):
                findIndex = 2
                break
            else:
                Util.MouseWheelScroll(-500)
                Util.SleepTime(1)

    if Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\HTML 작성", 0, 0, 1):
        Util.NowMouseClick()
        if isAdd == False:
            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\확인", 5, 5)
    Util.SleepTime(1)
    if findIndex == 1:
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\녹색 상세설명", 100, -150, 1)
    elif findIndex == 2:
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\녹색 상세설명_v2", 100, -150, 1)
    Util.SleepTime(1)
    Util.KeyboardKeyHotkey("ctrl", "a")
    Util.SleepTime(1)
    Util.KeyboardKeyPress("delete")
    Util.SleepTime(1)
    # html 내용 작성
    if True:
        htmlData = '<div style="text-align: center;">'
        htmlData += '<img src="https://nacharhan.github.io/photo/2.png"/>'
        htmlData += "`r`n"
        for index in range(arraySizesAndImgUrls.MaxIndex()):
            colorName = arraySizesAndImgUrls[index][1]
            imgUrls = arraySizesAndImgUrls[index][3]

            htmlData += '<div style="text-align: center;">'
            htmlData += (
                '<div><span style="font-size: 30px;">' + colorName + "</span></div>"
            )
            htmlData += "`r`n"
            for index2 in range(imgUrls.MaxIndex()):
                htmlData += '<div style="text-align: center;">'
                htmlData += '<img src="' + imgUrls[index2] + '"/>'
                htmlData += "`r`n"

        htmlData += '<div style="text-align: center;">'
        htmlData += '<img src="https://nacharhan.github.io/photo/11.png"/>'
        htmlData += "`r`n"
    pyperclip.copy(htmlData)
    Util.SleepTime(1)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(1)


# 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
def UpdateOptionsFromExcel(customsDuty):
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션", 0, 0, -500)
    Util.SleepTime(1)
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\엑셀 일괄등록", 0, 0)
    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\확인", 0, 0, 2)
    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\엑셀 일괄등록하기", 0, 0)
    Util.SleepTime(4)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\열기\즐겨찾기에 엑셀 폴더", 40, 5)
    Util.SleepTime(1)
    if customsDuty == 1:
        Util.DoubleClickAtWhileFoundImage(r"스마트 스토어\열기\옵션 세팅된 엑셀", 5, 5)
    else:
        Util.DoubleClickAtWhileFoundImage(r"스마트 스토어\열기\옵션 세팅된 엑셀2", 5, 5)
    Util.SleepTime(1)


def AddOneProduct_Ugg(wbAddBefore, wbAdd, addOneProductSuccess, krwUsd):
    wsAddBefore = wbAddBefore.Sheets(1)
    wsAdd = wbAdd.Sheets(1)

    url = wsAddBefore[f"{Enum_COLUMN.C.name}{1}"].value

    data = GetUggData(url, krwUsd)

    # UGG에 사이즈 정보로 정보 취합
    useMoney = data.useMoney

    # 이중 배열
    arraySizesAndImgUrls = data.arraySizesAndImgUrls

    # 상품 이름
    title = data.title

    Util.FolderToDelete(EnvData.g_DefaultPath() + r"\DownloadImage")

    if len(arraySizesAndImgUrls) == 0:
        Util.TelegramSend("len(arraySizesAndImgUrls) == 0 url : " + url)
        # 등록해야 될 것에서 삭제
        wsAddBefore.Rows(1).Delete()
        wbAddBefore.Save()

        values = {}
        values.addCount = False
        values.addOneProductSuccess = True
        return values

    if len(arraySizesAndImgUrls) >= 1:
        imgUrls = arraySizesAndImgUrls[1][3]
        for i in range(len(imgUrls)):
            Util.DownloadImageUrl(imgUrls[i], i)

    Util.SleepTime(1)
    webbrowser.open("https://sell.smartstore.naver.com/#/products/create")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "tab")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(2)

    if Util.ClickAtWhileFoundImage(r"스마트 스토어\로그인하기", 5, 5, 1):
        Util.SleepTime(2)
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\로그인", 5, 5, 1):
        Util.SleepTime(1)
        webbrowser.open("https://sell.smartstore.naver.com/#/products/create")
        Util.SleepTime(0.5)
        Util.KeyboardKeyHotkey("ctrl", "tab")
        Util.SleepTime(0.5)
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(2)

    Util.SleepTime(2)
    Util.KeyboardKeyPress("esc")
    Util.SleepTime(1)

    # if not addOneProductSuccess:
    # 	Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\이전 내용 불러오기 확인", -80, 5, 5)

    # 상품 등록시 기본 세팅 들
    ProductRegistration.ProductRegistrationDefaultSettings()

    # 카테고리명 입력
    Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\카테고리명 선택", 0, 50)
    Util.SleepTime(0.5)
    Util.NowMouseClick()
    pyperclip.copy(wsAddBefore[f"{Enum_COLUMN.B.name}{1}"].value)
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)
    Util.KeyboardKeyPress("enter")
    Util.SleepTime(2)

    # 상품명 입력
    Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\상품명", 50, 85)
    Util.SleepTime(1)
    Util.NowMouseClick()
    Util.SleepTime(0.5)
    pyperclip.copy(title)
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)

    # 판매가 입력
    UpdateAndReturnSalePrice(data.korMony)

    # 옵션 세팅
    if True:
        # 관세 부가 여부 체크
        customsDuty = useMoney >= 200

        # 옵션 엑셀 세팅
        Util.SetExcelOption(arraySizesAndImgUrls, customsDuty)

        # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
        UpdateOptionsFromExcel(customsDuty)

    # 이미지 등록(대표, 추가)
    if ProductRegistration.IamgeRegistration_v2() == False:
        Util.TelegramSend("대표 이미지 등록 못함(이미지 너무 큼) url : " + url)
        # 등록해야 될 것에서 삭제
        wsAddBefore.Rows(1).Delete()
        wbAddBefore.Save()

        values = {}
        values.addCount = False
        values.addOneProductSuccess = True
        return values

    # HTML 으로 등록
    SetHTML(arraySizesAndImgUrls, True)

    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
    Util.SleepTime(5)
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", -80, 5):
        Util.SleepTime(3)

        # 동록한 엑셀에 기록(앞에 추가)
        wsAdd.Rows(2).Insert()

        # 상품 url
        Util.GoToTheAddressWindow()
        Util.SleepTime(0.5)
        addurl = Util.CopyToClipboardAndGet()
        Util.Debug("addurl : " + addurl)
        wsAdd[f"{Enum_COLUMN.B.name}{2}"].value = addurl
        # 크롬 탭 닫기
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(0.5)
        # 상품 번호
        addUrlSplitArray = addurl.split("/")
        if len(addUrlSplitArray) > 0:
            wsAdd[f"{Enum_COLUMN.A.name}{2}"].value = addUrlSplitArray[len(addUrlSplitArray)]

        wsAdd[f"{Enum_COLUMN.C.name}{2}"].value = url
        # 상품명 기재
        wsAdd[f"{Enum_COLUMN.T.name}{2}"].value = title
        # 가격
        wsAdd[f"{Enum_COLUMN.U.name}{2}"].value = useMoney
        # 브랜드
        brand = wsAddBefore(f"{Enum_COLUMN.A.name}{2}").value
        wsAdd[f"{Enum_COLUMN.E.name}{2}"].value = brand

        # 색이름 리스트 값
        colorNames = []
        for arraySizesAndImgUrl in arraySizesAndImgUrls:
            colorNames.append(arraySizesAndImgUrl[1])
        str_saveColorList = Util.JoinArrayToString(colorNames)
        Util.Debug("str_saveColorList : " + str_saveColorList)

        wsAdd[f"{Enum_COLUMN.F.name}{2}"].value = str_saveColorList

        # 색 이름 과 사아즈 리스트 값(이중 배열)
        str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
        Util.Debug("str_saveColorNameDoubleArray : " + str_saveColorNameDoubleArray)

        wsAdd[f"{Enum_COLUMN.G.name}{2}"].value = str_saveColorNameDoubleArray

        xl_J_(wbAdd, wsAdd, 2, "신규 등록", True)

        wbAdd.Save()

        # 등록해야 될 것에서 삭제
        wsAddBefore.Rows(1).Delete()
        wbAddBefore.Save()

        values = {}
        values.addCount = True
        values.addOneProductSuccess = True
        return values
    else:
        Util.TelegramSend("++++++++++++++ 이름이 입력 안됬음 왜지??")
        # 이름이 입력 안됬음  왜지??
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\취소", 5, 5)
        Util.SleepTime(1)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품취소 유실 확인", 5, 5)
        Util.SleepTime(1)

        values = {}
        values.addCount = False
        values.addOneProductSuccess = False
        return values


# 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
def AddDataFromExcel_Ugg():
    Util.TelegramSend(
        "추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 시작"
    )

    xlFileAddBefore = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.xlsx"
    wbAddBefore = openpyxl.load_workbook(xlFileAddBefore)
    wsAddBefore = wbAddBefore.active
    rowCountAddBefore = wsAddBefore.max_row

    xlFileAdd = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.xlsx"
    wbAdd = openpyxl.load_workbook(xlFileAdd)
    wsAdd = wbAdd.active
    rowCountAdd = wsAdd.max_row

    krwUsd = Util.KRWUSD()

    # 모두 반복
    addCount = 0
    count = 1
    for _ in range(rowCountAddBefore):
        addOneProductSuccess = True
        while True:
            Util.TelegramSend(count + "/" + rowCountAddBefore)
            data = AddOneProduct_Ugg(wbAddBefore, wbAdd, addOneProductSuccess, krwUsd)
            addOneProductSuccess = data.addOneProductSuccess
            if addOneProductSuccess:
                if data.addCount:
                    addCount += 1
                count += 1
                break

    wbAdd.Quit()
    wbAddBefore.Quit()
    Util.TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 끝")

    return addCount


# url에 필요 정보 가져오기
def GetUggData(url, exchangeRate, onlyUseMoney=False):
    # UGG에 사이즈 정보로 정보 취합
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    # 상품 이름
    title = ""

    if True:
        # 1145990은 url의 끝에 .html 전에 있는 값
        urlSplitArray = url.split("/")
        if len(urlSplitArray) > 0:
            productNumber = urlSplitArray[-1].replace(".html", "")
        else:
            productNumber = url

        webbrowser.open(url)
        # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
        # WinWait, ugg
        Util.SleepTime(10)
        htmlElementsData = GetElementsData()
        # Ctrl + W를 눌러 현재 Chrome 탭 닫기
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(1)

        # 상품 이름
        match = re.search(
            r'<div\s+class\s*=\s*"sticky-toolbar__content">\s*<span>([^<]*)</span>',
            htmlElementsData,
        )
        if match:
            title = match.group(1)
            Util.Debug("title : " + title)

        # <div class="sticky-toolbar__content"
        match = re.search(
            r'<div\s+class\s*=\s*"sticky-toolbar__content"', htmlElementsData
        )
        if match:
            startPos = match.start()
        # aria-labelledby="size" 까지에서 찾기
        match = re.search(r'aria-labelledby\s*=\s*"size"', htmlElementsData[startPos:])
        if match:
            endPos = match.end() + startPos
        if startPos and endPos:
            contentValue = htmlElementsData[startPos:endPos]
            useMoneys = Util.GetRegExMatche1List(contentValue, r"\$(.+)")
            if len(useMoneys) > 0:
                useMoney = useMoneys[len(useMoneys)] + 0
                Util.Debug("useMoney : " + useMoney)

            korMony = Util.GetKorMony(useMoney, exchangeRate)

            if onlyUseMoney == False:
                urlEndColorNames = Util.GetRegExMatche1List(contentValue,r'<span data-attr-value="([^"]*)" class="color-value swatch swatch-circle"')
                colorNames = Util.GetRegExMatche1List(contentValue, r'data-attr-color-swatch="[^"]*"\s*title="([^"]*)"')

                if len(urlEndColorNames) == len(colorNames):
                    for index in range(len(urlEndColorNames)):
                        # .html?dwvar_1145990_color=BCDR 제일 뒤에 색 정보 적어서 url 열 수 있음
                        # 색 위치로 클릭하는 것보다 url 열는 것이 더 낫다고 생각됨
                        colorUrl = f"{url}?dwvar_{productNumber}_color={urlEndColorNames[index]}"

                        Util.Debug(f"urlEndColorNames[{index}] : {urlEndColorNames[index]}")

                        Util.Debug("colorNames[" + index + "] : " + colorNames[index])

                        webbrowser.open(colorUrl)
                        # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
                        # WinWait, ugg
                        Util.SleepTime(10)
                        colorUrlHtmlElementsData = GetElementsData()
                        # Ctrl + W를 눌러 현재 Chrome 탭 닫기
                        Util.KeyboardKeyHotkey("ctrl", "w")
                        Util.SleepTime(1)

                        # 이미지 url 알아오는 것
                        if True:
                            imgBigUrls = []
                            imgUrls = Util.GetRegExMatche1List(colorUrlHtmlElementsData,r'<img[^>]+data-srcset="([^"]+)"[^>]+>')
                            for imgUrl in imgUrls:
                                splitArray = imgUrl.split(" , ")
                                if len(splitArray) > 0:
                                    # 제일 끝에 것으로 하는 이유는 이미지 제일 큰 Url 이기 때문에
                                    if re.search(r"https:(.+)\.(png|jpg)", splitArray[-1]):
                                        imgBigUrls.append(re.search(r"https:(.+)\.(png|jpg)", splitArray[-1]))

                        # 구매 가능한 사이즈 구하기
                        if True:
                            sizes = []
                            # <div class="sticky-toolbar__content"> 는 뒤에 줄에 이는 곳 부터 찾기
                            match = re.search(r'<div\s+class\s*=\s*"sticky-toolbar__content"',colorUrlHtmlElementsData)
                            if match:
                                # 패턴을 찾았을 때의 시작 위치
                                startPos = match.start()
                            else:
                                # 패턴을 찾지 못했을 때의 처리
                                startPos = 0
                            # 정규식으로 특정 줄에 options-select과 https://www과 _1145990_ 과 data-attr-value 이 포함되는 줄이 있는지 체크
                            # sizeLines 리스트 초기화
                            sizeLines = Util.GetRegExMatcheList(
                                colorUrlHtmlElementsData,
                                r'options-select\s*.*value="(https:\/\/www\.ugg\.com\/on\/demandware\.store\/Sites-UGG-US-Site\/en_US\/Product-Variation\?dwvar_'
                                + re.escape(productNumber)
                                + '_color=[^"]+)".*data-attr-value="([^"]+)"',
                                startPos,
                            )
                            # sizeLines 리스트에 대한 루프
                            for line in sizeLines:
                                # options-select와 value 사이에 값이 없는 것만 찾기
                                if re.search(r'options-select\s*""\s*value="([^"]*)"', line):
                                    # 데이터 추출
                                    match = re.search(r'data-attr-value="([^"]+)"', line)
                                    if match:
                                        extractedValue = match.group(1)
                                        defaultExtractedValue = extractedValue
                                        extractedValueSplitArray = extractedValue.split("/")
                                        if len(extractedValueSplitArray) == 2:
                                            extractedValue = extractedValueSplitArray[-1]

                                        # 숫자로 판단되면 앞의 0 제거
                                        if re.match(r"^[0-9]+(\.[0-9]+)?$", extractedValue):
                                            if extractedValue.startswith("0") and (len(extractedValue) < 2 or extractedValue[1] != "."):
                                                extractedValue = extractedValue[1:]

                                            sizeData = f"US_{extractedValue}({Util.GetUggKorSize(extractedValue)})"

                                            Util.Debug("size : " + sizeData)
                                            sizes.append(sizeData)
                                        else:
                                            sizes.append(defaultExtractedValue)

                        if len(sizes) > 0 and len(imgBigUrls) > 0:
                            # 문자열의 길이가 25보다 큰지 확인
                            if len(colorNames[index]) > 25:
                                # 25자까지만 잘라내기
                                colorName = colorNames[index][:25]
                            else:
                                colorName = colorNames[index]
                            arraySizesAndImgUrls.append([colorName, sizes, imgBigUrls])

    values = {}
    values.useMoney = useMoney
    values.korMony = korMony
    values.arraySizesAndImgUrls = arraySizesAndImgUrls
    values.title = title
    return values


def GetMytheresaData(url, exchangeRate):
    # 사이즈 정보로 정보 취합
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    # 상품 이름
    title = ""

    isSoldOut = False

    if True:
        webbrowser.open(url)
        # "mytheresa"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
        # WinWait, mytheresa
        Util.SleepTime(10)
        htmlElementsData = GetElementsData()
        # Ctrl + W를 눌러 현재 Chrome 탭 닫기
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(1)

        # 정규 표현식과 매치되는지 확인
        match = re.search(r'PriceSpecification": {\s+"price": (\d+)', htmlElementsData)
        if match:
            useMoney = match.group(1)

        korMony = Util.GetKorMony(useMoney, exchangeRate)

        # 상품 이름
        match = re.search(r'"priceCurrency":\s*"([A-Z]{3})"', htmlElementsData)
        if match:
            priceCurrency = match.group(1)

        match = re.search(r">Sold Out<", htmlElementsData)
        if match:
            isSoldOut = True

        match = re.search(r"error404__title", htmlElementsData)
        if match:
            isSoldOut = True

        sizeLines = Util.GetRegExMatche1List(
            htmlElementsData, r'<span class="sizeitem__label">(.*?)</span>'
        )

        sizes = []
        for sizeLine in sizeLines:
            sizeData = sizeLine + "(" + Util.GetMytheresaKorSize(sizeLine) + ")"
            Util.Debug("size : " + sizeData)
            sizes.append(sizeData)

        colorName = "One Color"
        imgBigUrls = []
        arraySizesAndImgUrls.append([colorName, sizes, imgBigUrls])

    values = {}
    values.useMoney = useMoney
    values.korMony = korMony
    values.arraySizesAndImgUrls = arraySizesAndImgUrls
    values.title = title
    values.sizesLength = len(sizes)
    values.isSoldOut = isSoldOut
    return values


# 스마트 스토어 수정 화면까지 이동
def ManageAndModifyProducts(ws, row):
    values = {}
    values.isNoProduct = False
    values.isNoNetwork = False

    Util.SleepTime(1)
    webbrowser.open("https://sell.smartstore.naver.com/#/products/origin-list")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "tab")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(2)

    if Util.ClickAtWhileFoundImage(r"스마트 스토어\로그인하기", 5, 5, 1):
        Util.SleepTime(2)
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\로그인", 5, 5, 1):
        Util.SleepTime(1)
        webbrowser.open("https://sell.smartstore.naver.com/#/products/origin-list")
        Util.SleepTime(0.5)
        Util.KeyboardKeyHotkey("ctrl", "tab")
        Util.SleepTime(0.5)
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(2)

    Util.SleepTime(2)
    Util.KeyboardKeyPress("esc")
    Util.SleepTime(1)

    # 상품 조회해서 상품 수정 화면으로 이동
    if True:
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 조회\상품번호", 150, 10)
        Util.SleepTime(0.5)
        pyperclip.copy(round(ws(f"{Enum_COLUMN.A.name}{row}").value))
        Util.SleepTime(0.5)
        # 상품번호 붙여넣기
        Util.KeyboardKeyHotkey("ctrl", "v")
        Util.SleepTime(0.5)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 조회\검색", 0, 0)
        Util.SleepTime(1)
        if False == Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 조회\수정", 0, 0, 5):
            values.isNoProduct = True
            return values
        Util.SleepTime(2)

    # "스마트 스토어\네트워크 불안정 느낌표"
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\네트워크 불안정", 0, 0, 1):
        values.isNoNetwork = True
        return values

    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\KC인증", 0, 0, 2):
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\KC인증 닫기", 10, 10, 2)

    return values


# 품절
def SoldOut(wb, ws, row):
    # 품절
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션", 0, 0, -500)
    Util.SleepTime(1)
    Util.WheelAndMoveAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션에서 선택형", 0, 0, -500)
    Util.SleepTime(1)
    optionEnd = pyautogui.position()
    Util.SleepTime(0.5)
    Util.MoveAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션", 0, 0, 2, 1, optionEnd.x - 50, optionEnd.y - 150)
    Util.SleepTime(1)
    optionStart = pyautogui.position()
    Util.SleepTime(0.5)
    Util.MouseMove(optionEnd.x + 1000, optionEnd.y + 100)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션에서 선택형에서 설정함 상태",0,0,2,1,optionStart.x,optionStart.y,optionEnd.x + 1000,optionEnd.y + 100)
    Util.WheelAndMoveAtWhileFoundImage(r"스마트 스토어\상품 수정\재고수량에 개", 0, 0, 500)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\재고수량에 개", -80, 0)
    Util.SleepTime(1)
    Util.KeyboardKeyHotkey("ctrl", "a")
    Util.SleepTime(0.5)
    Util.KeyboardKeyPress("0")
    Util.SleepTime(1.5)
    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
    # Util.SleepTime(5)
    # Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", 5, 5)
    # Util.SleepTime(1)

    xl_J_(wb, ws, row, "품절 상태로 변경 완료", True)

    Util.TelegramSend("품절 상태로 변경 완료 row(" + row + ") ")


def UpdateAndReturnSalePrice(korMony):
    # 판매가 입력
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\판매가", 250, 85)
    Util.SleepTime(0.5)
    Util.NowMouseClick()
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "a")
    Util.SleepTime(0.5)
    Util.KeyboardKeyPress("delete")
    Util.SleepTime(0.5)
    pyperclip.copy(korMony)
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)


# 품절 완료 된 것 엑셀에서 제거
def RemoveCompletedSoldOutItems(wb):
    ws = wb.Sheets(1)

    startCount = 2
    while True:
        Util.Debug("xlWorksheet.UsedRange.Rows.Count : " + ws.UsedRange.Rows.Count)
        # 중복 여부를 체크합니다.
        isFind = False
        for i in range(ws.UsedRange.Rows.Coun):
            if i >= startCount:
                if "품절 상태로 변경 완료" in ws[f"{Enum_COLUMN.J.name}{i}"].value:
                    ws.Rows(i).Delete()
                    wb.Save()
                    isFind = True
                    startCount = i
                    Util.Debug(i)
                    break

        if isFind == False:
            break

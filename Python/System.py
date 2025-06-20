﻿import GlobalData
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
import System
import shutil
import psutil
import pandas as pd
import numpy as np
from lxml import etree
from collections import OrderedDict
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json



from collections import namedtuple


class Data_Ugg:
    def __init__(self):
        self.useMoney: float = 0
        self.korMony: float = 0
        self.arraySizesAndImgUrls = []
        self.title = ""
        self.isCheckRobot: bool = False
        self.details = ""  # 상세 정보
        self.htmlEndAdd = ""

class Data_BananarePublic:
    def __init__(self):
        self.useMoney: float = 0
        self.korMony: float = 0
        self.arraySizesAndImgUrls = []
        self.title = ""
        self.isSoldOut: bool = False
        self.details = ""  # 상세 정보
        self.fabricAndCare = ""  # 패브릭&케어
        self.htmlEndAdd = ""
        
class Data_Zara:
    def __init__(self):
        self.useMoney: float = 0
        self.korMony: float = 0
        self.arraySizesAndImgUrls = []
        self.title = ""
        self.isSoldOut: bool = False
        self.details = ""  # 상세 정보
        self.fabricAndCare = ""  # 패브릭&케어
        self.htmlEndAdd = ""

class Data_Mytheresa:
    def __init__(self):
        self.useMoney: float = 0
        self.korMony: float = 0
        self.arraySizesAndImgUrls = []
        self.title = ""
        self.isSoldOut: bool = False
        self.details = ""  # 상세 정보
        self.fabricAndCare = ""  # 패브릭&케어
        self.htmlEndAdd = ""

class AddOneProduct_Data_Common:
    def __init__(self):
        self.addCount: bool = False
        self.addOneProductSuccess: bool = False
        self.dfAddBefore = None
        self.dfAdd = None

class AddOneProduct_Data_Ugg:
    def __init__(self):
        self.addCount: bool = False
        self.addOneProductSuccess: bool = False
        self.dfAddBefore = None
        self.dfAdd = None

class AddOneProduct_Data_BananarePublic:
    def __init__(self):
        self.addCount: bool = False
        self.addOneProductSuccess: bool = False
        self.dfAddBefore = None
        self.dfAdd = None
        
class AddOneProduct_Data_Zara:
    def __init__(self):
        self.addCount: bool = False
        self.addOneProductSuccess: bool = False
        self.dfAddBefore = None
        self.dfAdd = None


class ManageAndModifyProductsData:
    def __init__(self):
        self.isNoProduct: bool = False
        self.isNoNetwork: bool = False


class NewProductURLs_Ugg:
    def __init__(self):
        self.name = ""
        self.productUrls = []


class NewProducts_BananarePublic:
    def __init__(self):
        self.name = ""
        self.titleAndPids = []


class NewProducts_Zara:
    def __init__(self):
        self.name = ""
        self.titleAndPids = []
        
class NewProducts_Common:
    def __init__(self):
        self.name = ""
        self.titleAndPids = []

class COLUMN_Add(Enum):
    A = "A"  # 브랜드 칸 = E
    B = "B"  # 카테고리 칸 = V
    C = "C"  # 상품 구매 url 칸

class COLUMN(Enum):
    A = "A"  # 상품 번호 칸
    B = "B"  # 상품 url 칸
    C = "C"  # 상품 구매 url 칸
    E = "E"  # 브랜드 칸
    F = "F"  # 색 RGB(16진수) 리스트 칸
    G = "G"  # 색명(사이즈 리스트) 칸
    H = "H"  # 업데이트 시간 칸
    I = "I"  # 체크 시간 칸
    J = "J"  # 체크 상태 칸
    K = "K"  # 이전 색RGB(16진수) 리스트 칸
    L = "L"  # 이전 색명(사아즈 리스트) 칸
    O = "O"
    P = "P"  # 마지막 실행 시켰던 라인
    Q = "Q"  # 마지막 실행 시켰던 라인의 시간 입력
    T = "T"  # 상품 이름
    U = "U"  # 상품 원본 가격
    V = "V"  # 상품 카테고리

firstName_BananarePublic = "[BananarePublic]"
firstName_Zara = "[Zara] 자라"

# 현재 열려 있는 엑셀 프로세스 닫기
def CloseExcelProcesses():
    for process in psutil.process_iter():
        if process.name() == "EXCEL.EXE":
            process.kill()


def SaveWorksheet(df):
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    # xlFile_copy = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트_복제.CSV"

    Util.Debug("save 원본 엑셀 파일 저장 시작", False)
    # 원본 엑셀 파일 저장
    Util.CsvSave(df, xlFile)
    Util.Debug("save 원본 엑셀 파일 저장 끝", False)

    # shutil.copy(xlFile, xlFile_copy)


def GetElementsData() -> str:
    Util.KeyboardKeyPress("f12")
    Util.SleepTime(2)
    if Util.WhileFoundImage(r"크롬\Elements에 lang"):
        Util.MoveAtWhileFoundImage(r"크롬\Elements에 lang", 5, 5)
        Util.SleepTime(0.5)
        Util.NowMouseClickRight()
        Util.SleepTime(3)
        # currentPos = pyautogui.position()
        Util.MoveAtWhileFoundImage(
            r"크롬\Elements에 html의 copy", 5, 5, 10, 1 #, currentPos.x, currentPos.y
        )
        Util.SleepTime(1)
        Util.MoveAtWhileFoundImage(
            r"크롬\Elements에 html의 copy에 copy element",
            5,
            5,
            10,
            1,
            # currentPos.x,
            # currentPos.y,
        )
        Util.SleepTime(1)
        Util.NowMouseClick()
        Util.SleepTime(3)
        outElementsData = pyperclip.paste()
        Util.SleepTime(0.5)
        return outElementsData
    return ""

def GetElementsData_Zara_v2(url: str, colorName = None) -> str:
    html = ""
    
    # 웹 드라이버 초기화
    driver = webdriver.Chrome()

    # 웹 페이지 열기
    driver.get(url)
    
    Util.SleepTime(5)
    
    try:
        # 페이지를 전체 화면으로 표시
        driver.maximize_window()
        
        # <button id="onetrust-accept-btn-handler">Accept All Cookies</button>
        # "Accept All Cookies" 버튼을 찾습니다.
        accept_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(text(), 'Accept All Cookies')]")
            )
        )

        # 버튼 클릭
        accept_button.click()

        # <span>Yes, stay on United States</span>
        yes_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Yes, stay on United States')]")
            )
        )

        # 버튼 클릭
        yes_button.click()

        if colorName != None:
            # <button class="product-detail-color-selector__color-button" data-qa-action="select-color"><div class="product-detail-color-selector__color-area" style="background-color:#88491D"> <span class="screen-reader-text">Brown</span></div></button>
            # "Brown" 텍스트를 포함하는 버튼을 클릭합니다. // <span> 요소 내에 있으므로
            button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, f"//button[contains(., '{colorName}')]")
                )
            )

            # 버튼 클릭
            button.click()

        html = driver.page_source

    except Exception as e:
        Util.DiscordSend("버튼을 찾을 수 없습니다:", e);

    finally:
        # 웹 브라우저 종료
        driver.quit()

    return html

def GetElementsData_v2(url: str) -> str:
    html = ""
    
    # 웹 드라이버 초기화
    driver = webdriver.Chrome()

    # 웹 페이지 열기
    driver.get(url)
    Util.SleepTime(5)

    try:
        html = driver.page_source

    except Exception as e:
        print("버튼을 찾을 수 없습니다:", e)

    finally:
        # 웹 브라우저 종료
        driver.quit()

    return html

def GetElementsData_v3(url: str, sleepTime) -> str:
    html = ""
    
    # 웹 드라이버 초기화
    driver = webdriver.Chrome()

    # 웹 페이지 열기
    driver.get(url)
    Util.SleepTime(sleepTime)

    try:
        html = driver.page_source

    except Exception as e:
        print("버튼을 찾을 수 없습니다:", e)

    finally:
        # 웹 브라우저 종료
        driver.quit()

    return html



# 등록 된 상품 최신화
def UpdateStoreWithColorInformation(inputRow=-1):
    Util.TelegramSend("등록 된 상품 최신화 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    df = df.astype(str)
    lastRow = df.shape[0]

    if inputRow != -1:
        row = inputRow
    else:
        row = round(float(df.at[1, COLUMN.P.name]))

    krwUsd = Util.KRWUSD()
    krwEur = Util.KRWEUR()

    while True:
        row += 1

        if row >= lastRow:
            break

        Util.Debug(f"row({row}) / lastRow({lastRow})")
        if row % 10 == 0:
            Util.TelegramSend(
                f"__ row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}"
            )

        # if (
        #     df.at[row, COLUMN.J.name]
        #     and "품절 상태로 변경 완료" in df.at[row, COLUMN.J.name]
        # ):
        #     df.at[1, COLUMN.P.name] = row
        #     df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
        #     System.SaveWorksheet(df)
        #     continue

        url = df.at[row, COLUMN.C.name]

        if "www.ugg.com" in url:
            Util.TelegramSend(
                f"www.ugg.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()} -- url:{url}"
            )
            isUpdateProduct = UpdateProductInfo_UGG(df, url, row, krwUsd)
            if isUpdateProduct:
                df.at[1, COLUMN.P.name] = row
                df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
                System.SaveWorksheet(df)
                Util.TelegramSend(f"url : {df.at[row, COLUMN.B.name]}")
            else:
                row -= 1
        elif "www.mytheresa.com" in url:
            Util.TelegramSend(
                f"www.mytheresa.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()} -- url:{url}"
            )
            UpdateAndNotifyProduct(df, row, GetData_Mytheresa(url, krwEur))
        elif "bananarepublic.gap.com" in url:
            Util.TelegramSend(
                f"bananarepublic.gap.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()} -- url:{url}"
            )
            UpdateAndNotifyProduct(df, row, GetData_BananarePublic(url, krwUsd, df.at[row, COLUMN.V.name]))
        elif "zara.com" in url:
            Util.TelegramSend(
                f"zara.com row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()} -- url:{url}"
            )
            UpdateAndNotifyProduct(df, row, GetData_Zara(url, krwUsd))
        else:
            df.at[1, COLUMN.P.name] = row
            df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
            System.SaveWorksheet(df)

    Util.CsvSave(df, xlFile)

    Util.TelegramSend("등록 된 상품 최신화 -- 끝")

def UpdateAndNotifyProduct(df, row, data):
    isUpdateProduct = UpdateProductInfoMoney_Common(df, row, data)
    if isUpdateProduct:
        df.at[1, COLUMN.P.name] = row
        df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
        System.SaveWorksheet(df)
        Util.TelegramSend(f"url : {df.at[row, COLUMN.B.name]}")
    else:
        row -= 1


# def UpdateStoreWithColorInformationMoney_Mytheresa():
#     xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
#     df = pd.read_csv(xlFile, encoding="cp949")
#     lastRow = df.shape[0]

#     row = round(float(df.at[1, COLUMN.P.name]))

#     krwEur = Util.KRWEUR()

#     while True:
#         row += 1

#         if row > lastRow:
#             break

#         Util.Debug(f"row({row}) / lastRow({lastRow})")
#         if row % 10 == 0:
#             Util.TelegramSend(
#                 f"__ row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}"
#             )

#         # 웹 브라우저 열기 및 상품 url로 이동
#         url = df.at[row, COLUMN.C.name]
#         if "www.mytheresa.com" not in url:
#             df.at[1, COLUMN.P.name] = row
#             df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
#             System.SaveWorksheet(df)
#             continue

#         if (
#             df.at[row, COLUMN.J.name]
#             and "품절 상태로 변경 완료" in df.at[row, COLUMN.J.name]
#         ):
#             df.at[1, COLUMN.P.name] = row
#             df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
#             System.SaveWorksheet(df)
#             continue

#         Util.TelegramSend(
#             f"row({row}) / lastRow({lastRow}) {Util.GetFormattedCurrentDateTime()}"
#         )

#         isUpdateProduct = UpdateProductInfoMoney_Mytheresa(df, url, row, krwEur)
#         if isUpdateProduct:
#             df.at[1, COLUMN.P.name] = row
#             df.at[1, COLUMN.Q.name] = Util.GetFormattedCurrentDateTime()
#             System.SaveWorksheet(df)
#         else:
#             row -= 1

#     Util.CsvSave(df, xlFile)


def UpdateProductInfo_UGG(df, url, row, krwUsd):
    data: Data_Ugg = GetData_Ugg(url, krwUsd)

    if data.isCheckRobot:
        System.xl_J_(df, row, "로봇인지 체크해 걸려서 나중에 다시 시도 하세요")
        Util.TelegramSend("로봇인지 체크해 걸려서 나중에 다시 시도 하세요")
        Util.DiscordSend("로봇인지 체크해 걸려서 나중에 다시 시도 하세요")
        return True

    if int(data.korMony) == 25000:
        System.xl_J_(df, row, "korMony이 25000 나와서 나중에 다시 시도 하세요")
        Util.TelegramSend("korMony이 25000 나와서 나중에 다시 시도 하세요")
        return True

    # UGG에 사이즈 정보로 정보 취합
    useMoney = data.useMoney

    # 이중 배열
    arraySizesAndImgUrls = data.arraySizesAndImgUrls

    # 기존 것과 같은지 비교(같으면 스마트 스토어에 하지 않기 위함)
    before_SaveColorList = df.at[row, COLUMN.F.name]
    Util.Debug(f"before_SaveColorList : {before_SaveColorList}")

    # 기존 색 이름 과 사아즈를 변수로 저장
    before_SaveColorNameDoubleArray = df.at[row, COLUMN.G.name]
    Util.Debug(f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}")

    # 색이름 리스트 값
    colorNames = []
    for item in arraySizesAndImgUrls:
        colorNames.append(item[Util.Array_ColroName])
    str_saveColorList = Util.JoinArrayToString(colorNames)
    Util.Debug(f"str_saveColorList : {str_saveColorList}")

    # 색 이름 과 사아즈 리스트 값(이중 배열)
    str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
    Util.Debug(f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}")

    # 색이 없은 경우 자체가 연결 되지 않거나 물건 자체가 없어졌을 경우
    if str_saveColorNameDoubleArray == "" or useMoney == 0:
        # 스마트 스토어 수정 화면까지 이동
        managedata = ManageAndModifyProducts(df, row)
        if managedata.isNoNetwork == True:
            return False

        if managedata.isNoProduct == True:
            System.xl_J_(df, row, "스토어에 상품이 없습니다.")
            return True

        # 품절
        SoldOut(df, row)
    else:
        if (
            before_SaveColorNameDoubleArray == str_saveColorNameDoubleArray
            and df.at[row, COLUMN.U.name] == useMoney
        ):
            # 이전과 정보가 변함이 없을 경우(이전과 동일하다고 적고 다음으로 넘어감)
            System.xl_J_(df, row, "이전과 동일합니다.")
        else:
            # 이전과 달라졌음
            System.xl_J_(df, row, "이전과 동일하지 않아서 변경 하려고 합니다.")

            # 스마트 스토어 수정 화면까지 이동
            managedata = ManageAndModifyProducts(df, row)
            if managedata.isNoNetwork == True:
                return False

            if managedata.isNoProduct == True:
                System.xl_J_(df, row, "스토어에 상품이 없습니다.")
                return True

            # 가격 변동이 있으면 변경
            if df.at[row, COLUMN.U.name] != useMoney:
                # 판매가 입력
                UpdateAndReturnSalePrice(data.korMony)

            if before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray:
                # 관세 부가 여부 체크
                is_customsDuty = useMoney >= 200

                # 옵션 엑셀 세팅
                Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

                System.xl_J_(
                    df,
                    row,
                    "이전과 동일하지 않아서 변경 하려고 합니다.(옵션 엑셀 세팅 완료)",
                )

                # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
                UpdateOptionsFromExcel(is_customsDuty)

                # 색은 그대로 인 상태에서 사이즈 숫자만 바꿔서 상세 페이지 갱신 하지 않도록 처리
                if before_SaveColorList != str_saveColorList:
                    # HTML 으로 등록
                    SetHTML(arraySizesAndImgUrls, data.details, "")

            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
            # Util.SleepTime(5)
            # Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", 5, 5)
            # Util.SleepTime(1)

            if (
                before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray
                and df.at[row, COLUMN.U.name] != useMoney
            ):
                df.at[row, COLUMN.U.name] = useMoney

                # 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
                if True:
                    # 색 이름 리스트 표시
                    df.at[row, COLUMN.F.name] = str_saveColorList
                    Util.Debug(f"str_saveColorList : {str_saveColorList}")

                    # 색 이름과 사아즈 리스트 표시
                    df.at[row, COLUMN.G.name] = str_saveColorNameDoubleArray
                    Util.Debug(
                        f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}"
                    )

                    System.xl_J_(
                        df,
                        row,
                        "변경 완료(이전과 동일하지 않아)(이전 값 등록 전)",
                        True,
                    )

                    # 이전 색 이름 리스트 표시
                    df.at[row, COLUMN.K.name] = before_SaveColorList
                    Util.Debug(f"before_SaveColorList : {before_SaveColorList}")

                    # 이전 색 이름과 사아즈 리스트 표시
                    df.at[row, COLUMN.L.name] = before_SaveColorNameDoubleArray
                    Util.Debug(
                        f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}"
                    )

                    System.xl_J_(
                        df, row, "변경 완료(이전과 동일하지 않아)(가격과 사이즈)", True
                    )
            else:
                # 가격 변동이 있으면 변경
                if df.at[row, COLUMN.U.name] != useMoney:
                    df.at[row, COLUMN.U.name] = useMoney

                    System.xl_J_(df, row, "변경 완료(가격만 변동)", True)

                if before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray:
                    # 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
                    if True:
                        # 색 이름 리스트 표시
                        df.at[row, COLUMN.F.name] = str_saveColorList
                        Util.Debug(f"str_saveColorList : {str_saveColorList}")

                        # 색 이름과 사아즈 리스트 표시
                        df.at[row, COLUMN.G.name] = str_saveColorNameDoubleArray
                        Util.Debug(
                            f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}"
                        )

                        System.xl_J_(
                            df,
                            row,
                            "변경 완료(이전과 동일하지 않아)(이전 값 등록 전)",
                            True,
                        )

                        # 이전 색 이름 리스트 표시
                        df.at[row, COLUMN.K.name] = before_SaveColorList
                        Util.Debug(f"before_SaveColorList : {before_SaveColorList}")

                        # 이전 색 이름과 사아즈 리스트 표시
                        df.at[row, COLUMN.L.name] = before_SaveColorNameDoubleArray
                        Util.Debug(
                            f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}"
                        )

                        System.xl_J_(df, row, "변경 완료(이전과 동일하지 않아)", True)

    return True


def UpdateProductInfoMoney_Mytheresa(df, url, row, krwEur):
    data = GetData_Mytheresa(url, krwEur)

    if int(data.korMony) == 25000:
        System.xl_J_(df, row, "korMony이 25000 나와서 나중에 다시 시도 하세요")
        Util.TelegramSend("korMony이 25000 나와서 나중에 다시 시도 하세요")
        return True

    # 스마트 스토어 수정 화면까지 이동
    managedata = ManageAndModifyProducts(df, row)
    if managedata.isNoNetwork == True:
        return False

    if managedata.isNoProduct == True:
        System.xl_J_(df, row, "스토어에 상품이 없습니다.")
        return True

    if data.isSoldOut:
        # 품절
        SoldOut(df, row)
    else:
        if data.sizesLength == 0:

            useMoney = data.useMoney

            # 판매가 입력
            UpdateAndReturnSalePrice(data.korMony)

            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)

            if int(data.korMony) != 0:
                System.xl_J_(df, row, "변경 완료(가격만 변동)", True)
            else:
                System.xl_J_(df, row, "가격이 0이 나왔습니다.")
        else:

            useMoney = data.useMoney
            arraySizesAndImgUrls = data.arraySizesAndImgUrls

            # 관세 부가 여부 체크
            is_customsDuty = useMoney >= 150

            # 옵션 엑셀 세팅
            Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

            # 1. 가격 세팅
            # 2. 엑셀로 옵셥 세팅
            if True:
                # 판매가 입력
                UpdateAndReturnSalePrice(data.korMony)

                # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
                UpdateOptionsFromExcel(is_customsDuty)

                Util.SleepTime(1)
                Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
                # Util.SleepTime(5)
                # Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", 5, 5)
                # Util.SleepTime(1)

            if data.korMony != 0:
                System.xl_J_(df, row, "변경 완료(가격과 사이즈 변동)", True)
            else:
                System.xl_J_(df, row, "가격이 0이 나왔습니다.")

    return True


# def UpdateProductInfoMoney_BananarePublic(df, url, row, krwUsd):
#     data = GetData_BananarePublic(url, krwUsd)

#     if int(data.korMony) == 25000:
#         System.xl_J_(df, row, "korMony이 25000 나와서 나중에 다시 시도 하세요")
#         Util.TelegramSend("korMony이 25000 나와서 나중에 다시 시도 하세요")
#         return True

#     # 스마트 스토어 수정 화면까지 이동
#     managedata = ManageAndModifyProducts(df, row)
#     if managedata.isNoNetwork == True:
#         return False

#     if managedata.isNoProduct == True:
#         System.xl_J_(df, row, "스토어에 상품이 없습니다.")
#         return True

#     if data.isSoldOut:
#         # 품절
#         SoldOut(df, row)
#     else:
#         useMoney = data.useMoney
#         arraySizesAndImgUrls = data.arraySizesAndImgUrls

#         # 관세 부가 여부 체크
#         is_customsDuty = useMoney >= 150

#         # 옵션 엑셀 세팅
#         Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

#         # 1. 가격 세팅
#         # 2. 엑셀로 옵셥 세팅
#         if True:
#             # 판매가 입력
#             UpdateAndReturnSalePrice(data.korMony)

#             # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
#             UpdateOptionsFromExcel(is_customsDuty)

#             Util.SleepTime(1)
#             Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)

#         if data.korMony != 0:
#             System.xl_J_(df, row, "변경 완료(가격과 사이즈 변동)", True)
#         else:
#             System.xl_J_(df, row, "가격이 0이 나왔습니다.")

#     return True

def UpdateProductInfoMoney_Common(df, row, data):
    if int(data.korMony) == 25000:
        System.xl_J_(df, row, "korMony이 25000 나와서 나중에 다시 시도 하세요")
        Util.TelegramSend("korMony이 25000 나와서 나중에 다시 시도 하세요")
        return True

    # 스마트 스토어 수정 화면까지 이동
    managedata = ManageAndModifyProducts(df, row)
    if managedata.isNoNetwork == True:
        return False

    if managedata.isNoProduct == True:
        System.xl_J_(df, row, "스토어에 상품이 없습니다.")
        return True

    if data.isSoldOut or len(data.arraySizesAndImgUrls) == 0:
        # 품절
        SoldOut(df, row)
    else:
        useMoney = data.useMoney
        arraySizesAndImgUrls = data.arraySizesAndImgUrls

        # 관세 부가 여부 체크
        is_customsDuty = useMoney >= 150

        # 옵션 엑셀 세팅
        Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

        # 1. 가격 세팅
        # 2. 엑셀로 옵셥 세팅
        if True:
            # 판매가 입력
            UpdateAndReturnSalePrice(data.korMony)

            # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
            UpdateOptionsFromExcel(is_customsDuty)
            
            # 기존 것과 같은지 비교(같으면 스마트 스토어에 하지 않기 위함)
            before_SaveColorList = df.at[row, COLUMN.F.name]
            Util.Debug(f"before_SaveColorList : {before_SaveColorList}")

            # 기존 색 이름 과 사아즈를 변수로 저장
            before_SaveColorNameDoubleArray = df.at[row, COLUMN.G.name]
            Util.Debug(f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}")

            # 색이름 리스트 값
            colorNames = []
            for item in arraySizesAndImgUrls:
                colorNames.append(item[Util.Array_ColroName])
            str_saveColorList = Util.JoinArrayToString(colorNames)
            Util.Debug(f"str_saveColorList : {str_saveColorList}")

            # 색 이름 과 사아즈 리스트 값(이중 배열)
            str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
            Util.Debug(f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}")

            # # 색은 그대로 인 상태에서 사이즈 숫자만 바꿔서 상세 페이지 갱신 하지 않도록 처리
            # if before_SaveColorList != str_saveColorList:
            #     # HTML 으로 등록
            #     SetHTML(arraySizesAndImgUrls, data.details, data.htmlEndAdd)
                
            SetHTML(arraySizesAndImgUrls, data.details, data.htmlEndAdd)

            # 색 이름 리스트 표시
            df.at[row, COLUMN.F.name] = str_saveColorList
            Util.Debug(f"str_saveColorList : {str_saveColorList}")

            # 색 이름과 사아즈 리스트 표시
            df.at[row, COLUMN.G.name] = str_saveColorNameDoubleArray
            Util.Debug(
                f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}"
            )

            # 이전 색 이름 리스트 표시
            df.at[row, COLUMN.K.name] = before_SaveColorList
            Util.Debug(f"before_SaveColorList : {before_SaveColorList}")

            # 이전 색 이름과 사아즈 리스트 표시
            df.at[row, COLUMN.L.name] = before_SaveColorNameDoubleArray
            Util.Debug(
                f"before_SaveColorNameDoubleArray : {before_SaveColorNameDoubleArray}"
            )

            Util.SleepTime(1)
            Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)

        if data.korMony != 0:
            System.xl_J_(df, row, "변경 완료(가격과 사이즈 변동)", True)
        else:
            System.xl_J_(df, row, "가격이 0이 나왔습니다.")

    return True


def xl_J_(df, row, value, updateTime=False):
    if updateTime:
        # 갱신 시간 표시
        df.at[row, COLUMN.H.name] = Util.GetFormattedCurrentDateTime()
    # 체크 시간 표시
    df.at[row, COLUMN.I.name] = Util.GetFormattedCurrentDateTime()
    # 체크 상태 표시
    df.at[row, COLUMN.J.name] = value

    System.SaveWorksheet(df)


# UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
def GetNewProductURLs_UGG(name, url, filterUrls) -> NewProductURLs_Ugg:
    Util.TelegramSend(f"GetNewProductURLs_UGG() {name} -- 시작")
    webbrowser.open(url)
    Util.SleepTime(1)
    Util.KeyboardKeyPress("esc")
    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
    # WinWait, ugg
    Util.SleepTime(10)

    # 웹 제일 끝까지 스코롤 한다.
    while True:
        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        # -100000 틱 스크롤 다운
        Util.MouseWheelScroll(-100000)
        Util.SleepTime(1.5)
        Util.KeyboardKeyPress("up")
        Util.SleepTime(1)
        Util.KeyboardKeyPress("down")
        Util.KeyboardKeyPress("down")
        Util.SleepTime(1)
        # 화면 가로 및 세로 해상도 얻기
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        if EndBar(screen_width, screen_height):
            # 상품 더 보기가 있는지 체크
            if Util.ClickAtWhileFoundImage(
                r"UGG\상품 리스트\상품 더 보기 버튼", 0, 0, 3
            ):
                Util.SleepTime(5)
            else:  # 상품이 더이상 없음
                break

    htmlElementsData: str = System.GetElementsData()
    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    productUrls = []
    # <a href=" 과 " class="js-pdp-link image-link pdp-link"> 중간에 있는 값
    productUrlLines = Util.GetRegExMatcheGroup1List(
        htmlElementsData, r'<a href="(.*?)" class="js-pdp-link image-link pdp-link">'
    )
    for productUrlLine in productUrlLines:
        splitList = productUrlLine.split(".html")
        if len(splitList) > 0:
            productUrls.append(f"https://www.ugg.com{splitList[0]}.html")
        else:
            productUrls.append(productUrlLine)
    uniqueArr = []
    for productUrl in productUrls:
        for filterUrl in filterUrls:
            if str(productUrl) == str(filterUrl):
                uniqueArr.append(productUrl)
                break

    for uniqueValue in uniqueArr:
        ArrayRemove(productUrls, uniqueValue)

    # 중복 제거
    productUrls = list(set(productUrls))

    Util.TelegramSend(f"GetNewProductURLs_UGG() {name} -- 끝")
    returnValue = NewProductURLs_Ugg()
    returnValue.name = name
    returnValue.productUrls = productUrls
    return returnValue


# BananarePublic 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
def GetNewProducts_BananarePublic(name, url, filterTitles) -> NewProducts_BananarePublic:
    Util.TelegramSend(f"GetNewProductURLs_BananarePublic() {name} -- 시작")
    webbrowser.open(url)
    Util.SleepTime(1)
    Util.KeyboardKeyPress("esc")
    Util.SleepTime(10)

    # 웹 제일 끝까지 스코롤 한다.
    while True:
        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        # -100000 틱 스크롤 다운
        Util.MouseWheelScroll(-100000)
        Util.SleepTime(1.5)
        Util.KeyboardKeyPress("up")
        Util.SleepTime(1)
        Util.KeyboardKeyPress("down")
        Util.KeyboardKeyPress("down")
        Util.SleepTime(1)
        # 화면 가로 및 세로 해상도 얻기
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        if EndBar(screen_width, screen_height):
            # 상품 더 보기가 있는지 체크
            if False == Util.ClickAtWhileFoundImage(
                r"바나나 리퍼블릭\스크롤 끝 인식", 0, 0, 3
            ):
                Util.SleepTime(5)
            else:  # 상품이 더이상 없음
                break

    htmlElementsData: str = System.GetElementsData()
    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    titleAndPids = []
    productTitleAndPids = Util.GetRegExMatcheGroup1And2List(
        htmlElementsData, r'0"><img alt="(.*?)".*?" id="product(.*?)"'
    )
    titles = [item[0] for item in productTitleAndPids]
    # 중첩된 리스트의 중복 제거
    for titleAndPid in productTitleAndPids:
        title = titleAndPid[0]
        title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")
        count = titles.count(title)
        if (
            count != 0
        ):
            titles = [value for value in titles if value != title]
            pid = titleAndPid[1]
            if (
                filterTitles.count(f"{firstName_BananarePublic} {Util.TranslateToKorean(title)}") == 0
            ):  # 중복 제거
                if '"' not in pid:
                    titleAndPids.append(titleAndPid)

    Util.TelegramSend(f"len(titleAndPids) : {len(titleAndPids)}")

    Util.TelegramSend(f"GetNewProductURLs_BananarePublic() {name} -- 끝")
    returnValue = NewProducts_BananarePublic()
    returnValue.name = name
    returnValue.titleAndPids = titleAndPids
    return returnValue


# Zara 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
def GetNewProducts_Zara(name, url, filterTitles) -> NewProducts_Zara:
    Util.TelegramSend(f"GetNewProductURLs_Zara() {name} -- 시작")
    webbrowser.open(url)
    Util.SleepTime(1)
    Util.KeyboardKeyPress("esc")
    Util.SleepTime(10)

    # 웹 제일 끝까지 스코롤 한다.
    while True:
        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        # -100000 틱 스크롤 다운
        Util.MouseWheelScroll(-100000)
        Util.SleepTime(1.5)
        Util.KeyboardKeyPress("up")
        Util.SleepTime(1)
        Util.KeyboardKeyPress("down")
        Util.KeyboardKeyPress("down")
        Util.SleepTime(1)
        # 화면 가로 및 세로 해상도 얻기
        screen_width = win32api.GetSystemMetrics(0)
        screen_height = win32api.GetSystemMetrics(1)
        if EndBar(screen_width, screen_height):
            # 상품 더 보기가 있는지 체크
            if False == Util.ClickAtWhileFoundImage(
                r"자라\스크롤 끝 인식", 0, 0, 3
            ):
                Util.SleepTime(5)
            else:  # 상품이 더이상 없음
                break

    htmlElementsData: str = System.GetElementsData()
    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    titleAndPids = []
    productTitleAndPids = Util.GetRegExMatcheGroup1And2List(
        htmlElementsData,
        r'="product-click" draggable="false" href="https://www\.zara\.com/us/en/(.*?)-p(\d+)\.html" tabindex=',
    )
    titles = [item[0] for item in productTitleAndPids]
    # 중첩된 리스트의 중복 제거
    unique_productTitleAndPids = [list(set(sublist)) for sublist in productTitleAndPids]
    for titleAndPid in unique_productTitleAndPids:
        title = titleAndPid[0]
        title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")
        count = titles.count(title)
        if (
            count != 0
        ):
            titles = [value for value in titles if value != title]
            pid = titleAndPid[1]
            if (
                filterTitles.count(f"{firstName_Zara} {Util.TranslateToKorean(title.replace("-", " "))} {pid[1:5]}/{pid[6:]}")
                == 0
            ):  # 중복 제거
                if '"' not in pid:
                    titleAndPids.append(titleAndPid)

    Util.TelegramSend(f"len(titleAndPids) : {len(titleAndPids)}")

    Util.TelegramSend(f"GetNewProductURLs_Zara() {name} -- 끝")
    returnValue = NewProducts_Zara()
    returnValue.name = name
    returnValue.titleAndPids = titleAndPids
    return returnValue


def EndBar(screen_width, screen_height):
    return (
        Util.ClickAtWhileFoundImage(
            r"크롬\오른쪽 스트롤바가 제일 아래인 이미지",
            0,
            0,
            1,
            1,
            screen_width - 200,
            screen_height - 200,
        )
        or Util.ClickAtWhileFoundImage(
            r"크롬\오른쪽 스트롤바가 제일 아래인 이미지_v2",
            0,
            0,
            1,
            1,
            screen_width - 200,
            screen_height - 200,
        )
        or Util.ClickAtWhileFoundImage(
            r"크롬\오른쪽 스트롤바가 제일 아래인 이미지_v3",
            0,
            0,
            1,
            1,
            screen_width - 200,
            screen_height - 200,
        )
    )


def ArrayRemove(arr, value):
    for index, element in enumerate(arr):
        if element == value:
            arr.pop(index)
            break


# 신규 등록 할 UGG 목록을 엑셀에 정리
def SetCsvNewProductURLs_Ugg():
    Util.TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    lastRow = df.shape[0]

    Util.Debug("start Csv ugg url")
    # C 열의 데이터를 배열에 저장
    filterUrls = []
    for row_index in range(0, lastRow):
        url = str(df.at[row_index, "C"])
        if url is not None and "www.ugg.com" in url:
            filterUrls.append(url)
    Util.TelegramSend(f"end Csv ugg url Length : {str(len(filterUrls))}")

    # 메뉴 창이 한번은 열려야지 세부 메뉴 창이 정상으로 열림
    webbrowser.open("https://www.ugg.com/women-footwear")
    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
    # WinWait, ugg
    Util.SleepTime(10)

    # UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
    uggProductUrls: list[NewProductURLs_Ugg] = []
    with open('NewProductData.txt', 'r', encoding='utf-8') as file:
        jsonDataList = json.load(file)

    for jsonData in jsonDataList:
        if jsonData["brand"] == "UGG":
            uggProductUrls.append(
                GetNewProductURLs_UGG(
                    jsonData["category"],
                    jsonData["url"],
                    filterUrls,
                )
            )

    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    xlFile = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    try:
        df = pd.read_csv(xlFile, encoding="cp949")
    except pd.errors.EmptyDataError:
        # 빈 파일이므로 빈 데이터프레임 생성
        df = pd.DataFrame()

    # 모든 행을 삭제합니다.
    df.drop(df.index, inplace=True)

    allCount = 0
    for item in uggProductUrls:
        for productUrl in item.productUrls:
            allCount += 1
            # 각 셀에 값을 설정합니다.
            df.loc[allCount, COLUMN_Add.A.name] = "UGG"
            df.loc[allCount, COLUMN_Add.B.name] = item.name  # 메뉴
            df.loc[allCount, COLUMN_Add.C.name] = productUrl  # url

    Util.CsvSave(df, xlFile)

    Util.TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 끝")


# 신규 등록 할 BananarePublic 목록을 엑셀에 정리
def SetCsvNewProductURLs_BananarePublic():
    Util.TelegramSend("신규 등록 할 BananarePublic 목록을 엑셀에 정리 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    lastRow = df.shape[0]

    Util.Debug("start Csv BananarePublic url")
    # C 열의 데이터를 배열에 저장
    filterTitles = []
    for row_index in range(0, lastRow):
        url = str(df.at[row_index, COLUMN.C.name])
        title = str(df.at[row_index, COLUMN.T.name])
        if url is not None and "https://bananarepublic.gap.com/" in url:
            filterTitles.append(title)
    Util.TelegramSend(f"end Csv BananarePublic url Length : {str(len(filterTitles))}")
    
    newProducts: list[NewProducts_BananarePublic] = []
    with open('NewProductData.txt', 'r', encoding='utf-8') as file:
        jsonDataList = json.load(file)

    for jsonData in jsonDataList:
        if jsonData["brand"] == "BananarePublic":
            newProducts.append(
                GetNewProducts_BananarePublic(
                    jsonData["category"],
                    jsonData["url"],
                    filterTitles,
                )
            )
    
    # 중복 제거
    titles: list[str] = []
    unique_newProducts: list[NewProducts_BananarePublic] = []
    for newProduct in newProducts:
        titleAndPids = []
        for titleAndPid in newProduct.titleAndPids:
            title = titleAndPid[0]
            title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")
            if titles.count(title) == 0:
                titleAndPids.append(titleAndPid)
                titles.append(title)

        Util.TelegramSend(f"len(titleAndPids) : {len(titleAndPids)}")
        newProduct.titleAndPids = titleAndPids
        unique_newProducts.append(newProduct)

    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    xlFile = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    try:
        df = pd.read_csv(xlFile, encoding="cp949")
    except pd.errors.EmptyDataError:
        # 빈 파일이므로 빈 데이터프레임 생성
        df = pd.DataFrame()

    # 모든 행을 삭제합니다.
    df.drop(df.index, inplace=True)

    allCount = 0
    for item in unique_newProducts:
        for titleAndPid in item.titleAndPids:
            allCount += 1
            # 각 셀에 값을 설정합니다.
            df.loc[allCount, COLUMN_Add.A.name] = "BananarePublic"
            df.loc[allCount, COLUMN_Add.B.name] = item.name  # 메뉴
            df.loc[allCount, COLUMN_Add.C.name] = (
                f"https://bananarepublic.gap.com/browse/product.do?pid={titleAndPid[1]}"  # url
            )

    Util.CsvSave(df, xlFile)

    Util.TelegramSend("신규 등록 할 BananarePublic 목록을 엑셀에 정리 -- 끝")

def SetCsvNewProductURLs_Zara_v2():
    System.SetCsvNewProductURLs_Common(
        "Zara",
        "https://www.zara.com/",
        "https://www.zara.com/us/en/--p",
        ".html",
        System.GetNewProducts_Zara,
        "Zara"
    )

def SetCsvNewProductURLs_BananarePublic_v2():
    System.SetCsvNewProductURLs_Common(
        "BananarePublic",
        "https://bananarepublic.gap.com/",
        "https://bananarepublic.gap.com/browse/product.do?pid=",
        "",
        System.GetNewProducts_BananarePublic,
        "BananarePublic",
    )
    
# 신규 등록 할 Zara 목록을 엑셀에 정리
def SetCsvNewProductURLs_Common(logName, findFirstUrl, addStartUrl, addEndUrl, GetNewProducts, brand):
    Util.TelegramSend(f"신규 등록 할 {logName} 목록을 엑셀에 정리 -- 시작")
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    lastRow = df.shape[0]

    Util.Debug(f"Start Csv {logName} url")
    # C 열의 데이터를 배열에 저장
    filterTitles = []
    for row_index in range(0, lastRow):
        url = str(df.at[row_index, COLUMN.C.name])
        title = str(df.at[row_index, COLUMN.T.name])
        if url is not None and findFirstUrl in url:
            filterTitles.append(title)
    Util.TelegramSend(f"end Csv {logName} url Length : {str(len(filterTitles))}")

    # UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
    newProducts: list[NewProducts_Common] = []
    
    with open('NewProductData.txt', 'r', encoding='utf-8') as file:
        jsonDataList = json.load(file)

    for jsonData in jsonDataList:
        if jsonData["brand"] == brand:
            newProducts.append(
                GetNewProducts(
                    jsonData["category"],
                    jsonData["url"],
                    filterTitles,
                )
            )

    # 중복 제거
    titles: list[str] = []
    unique_newProducts: list[NewProducts_Common] = []
    for newProduct in newProducts:
        titleAndPids = []
        for titleAndPid in newProduct.titleAndPids:
            title = titleAndPid[0]
            title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")
            if titles.count(title) == 0:
                titleAndPids.append(titleAndPid)
                titles.append(title)

        Util.TelegramSend(f"len(titleAndPids) : {len(titleAndPids)}")
        newProduct.titleAndPids = titleAndPids
        unique_newProducts.append(newProduct)

    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    xlFile = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    try:
        df = pd.read_csv(xlFile, encoding="cp949")
    except pd.errors.EmptyDataError:
        # 빈 파일이므로 빈 데이터프레임 생성
        df = pd.DataFrame()

    # 모든 행을 삭제합니다.
    df.drop(df.index, inplace=True)

    allCount = 0
    for item in unique_newProducts:
        for titleAndPid in item.titleAndPids:
            allCount += 1
            # 각 셀에 값을 설정합니다.
            df.loc[allCount, COLUMN_Add.A.name] = logName
            df.loc[allCount, COLUMN_Add.B.name] = item.name  # 메뉴
            df.loc[allCount, COLUMN_Add.C.name] = (
                f"{addStartUrl}{titleAndPid[1]}{addEndUrl}"  # url
            )

    Util.CsvSave(df, xlFile)

    Util.TelegramSend(f"신규 등록 할 {logName} 목록을 엑셀에 정리 -- 끝")


# HTML 으로 등록
def SetHTML(arraySizesAndImgUrls, details, htmlEndAdd, isAdd=False):
    # 상세설명 찾아서 그 아래로 ONE 원형 검색이 존재 하는데 체크(이때는 아래로 조금씩 내리기)
    Util.WheelAndMoveAtWhileFoundImage(r"스마트 스토어\상품 수정\상세 설명")
    findIndex = 0
    while True:
        if Util.MoveAtWhileFoundImage(
            r"스마트 스토어\상품 수정\녹색 상세설명", 0, 0, 1
        ):
            findIndex = 1
            break
        else:
            if Util.MoveAtWhileFoundImage(
                r"스마트 스토어\상품 수정\녹색 상세설명_v2", 0, 0, 1
            ):
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
        Util.ClickAtWhileFoundImage(
            r"스마트 스토어\상품 수정\녹색 상세설명", 100, -150, 1
        )
    elif findIndex == 2:
        Util.ClickAtWhileFoundImage(
            r"스마트 스토어\상품 수정\녹색 상세설명_v2", 100, -150, 1
        )
    Util.SleepTime(1)
    Util.KeyboardKeyHotkey("ctrl", "a")
    Util.SleepTime(1)
    Util.KeyboardKeyPress("delete")
    Util.SleepTime(1)
    # html 내용 작성
    if True:
        htmlData = '<div style="text-align: center;">'
        htmlData += '<img src="https://nacharhan.github.io/photo/2.png"/>'
        if details != "":
            details = details.replace(": ", ". <br>")
            details = details.replace(". ", ". <br>")
            htmlData += "</div>"
            htmlData += '<div style="text-align: center;">'
            htmlData += "<p>상세 정보</p>"
            htmlData += "</div>"
            htmlData += "</div>"
            htmlData += '<div style="text-align: center;">'
            htmlData += f"<p>{details}</p>"
            htmlData += "</div>"

        for item in arraySizesAndImgUrls:
            colorName = item[Util.Array_ColroName]
            imgUrls = item[Util.Array_UrlList]

            htmlData += '<div style="text-align: center;">'
            htmlData += (
                '<div><span style="font-size: 30px;">' + colorName + "</span></div>"
            )
            for imgUrl in imgUrls:
                htmlData += '<div style="text-align: center;">'
                htmlData += '<img src="' + imgUrl + '"/>'

        htmlData += '<div style="text-align: center;">'
        htmlData += '<img src="https://nacharhan.github.io/photo/11.png"/>'
        htmlData += htmlEndAdd
    pyperclip.copy(htmlData)
    Util.SleepTime(1)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(1)


# 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
def UpdateOptionsFromExcel(is_customsDuty):
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션", 0, 0, -500)
    Util.SleepTime(0.5)
    Util.MouseWheelScroll(-100)
    Util.SleepTime(0.5)
    current_mouse_x, current_mouse_y = pyautogui.position()
    Util.MouseMove(current_mouse_x + 220, current_mouse_y + 100 - 100)
    Util.SleepTime(0.5)
    Util.NowMouseClick()
    Util.SleepTime(0.5)
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\엑셀 일괄등록", 0, 0)
    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\확인", 0, 0, 2)
    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\엑셀 일괄등록하기", 0, 0)
    Util.SleepTime(1.5)
    for _ in range(5):
        pyautogui.hotkey("tab")
        Util.SleepTime(0.2)
    pyautogui.hotkey("enter")
    Util.SleepTime(0.3)
    pyperclip.copy(EnvData.g_DefaultPath() + r"\엑셀")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)
    pyautogui.hotkey("enter")
    Util.SleepTime(1)
    if is_customsDuty == True:
        Util.DoubleClickAtWhileFoundImage(r"스마트 스토어\열기\옵션 세팅된 엑셀", 5, 5)
    else:
        Util.DoubleClickAtWhileFoundImage(
            r"스마트 스토어\열기\옵션 세팅된 엑셀2", 5, 5, 0.8
        )
    Util.SleepTime(1)


def AddOneProduct_Ugg(
    dfAddBefore, dfAdd, xlFileAddBefore, xlFileAdd, addOneProductSuccess, krwUsd
) -> AddOneProduct_Data_Ugg:

    url = dfAddBefore.at[dfAddBefore.index[0], COLUMN.C.name]

    data: Data_Ugg = GetData_Ugg(url, krwUsd)

    # UGG에 사이즈 정보로 정보 취합
    useMoney = data.useMoney

    # 이중 배열
    arraySizesAndImgUrls = data.arraySizesAndImgUrls

    # 상품 이름
    title = f"[Ugg] {data.title}"
    filtered_rows = dfAdd[dfAdd[System.COLUMN.T.name] == title]
    if len(filtered_rows) >= 1:
        if len(arraySizesAndImgUrls) >= 1:
            title += "(" + data.arraySizesAndImgUrls[0][Util.Array_ColroName] + ")"
        count = 2
        while True:
            filtered_rows = dfAdd[dfAdd[System.COLUMN.T.name] == title]
            if len(filtered_rows) >= 1:
                # 1개 이상의 행이 일치합니다.
                title += f"v{count}"
                count += 1
            else:
                # 일치하는 행이 없습니다.
                break

    Util.FolderToDelete(EnvData.g_DefaultPath() + r"\DownloadImage")

    if len(arraySizesAndImgUrls) == 0:
        Util.TelegramSend(f"len(arraySizesAndImgUrls) == 0 url : {url}")
        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        returnValue = AddOneProduct_Data_Ugg()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

    imgCount: int = 0
    if len(arraySizesAndImgUrls) >= 1:
        imgUrls = arraySizesAndImgUrls[0][Util.Array_UrlList]
        for i in range(len(imgUrls)):
            if i < 10: # 스마스 스토어 사진 추가 이미지 9까지만 밖에 안되서
                if True == Util.DownloadImageUrl(imgUrls[i], i, 750, 1000):
                    imgCount += 1

    if imgCount == 0:
        Util.TelegramSend(f"imgCount == 0 url : {url}")
        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        returnValue = AddOneProduct_Data_Ugg()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

    Util.SleepTime(1)
    webbrowser.open("https://sell.smartstore.naver.com/#/products/create")
    # 전에 있던 탭 창 삭제
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
        # 전에 있던 탭 창 삭제
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
    ProductRegistration.ProductCategory(
        dfAddBefore.at[dfAddBefore.index[0], COLUMN.B.name]
    )

    # 상품명 입력
    ProductRegistration.ProductTitle(title)

    # 판매가 입력
    UpdateAndReturnSalePrice(data.korMony)

    # 옵션 세팅
    if True:
        # 관세 부가 여부 체크
        is_customsDuty = useMoney >= 200

        # 옵션 엑셀 세팅
        Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

        # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
        UpdateOptionsFromExcel(is_customsDuty)

    # 이미지 등록(대표, 추가)
    ProductRegistration.IamgeRegistration_v2(imgCount)

    # HTML 으로 등록
    SetHTML(arraySizesAndImgUrls, data.details, "", True)

    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
    Util.SleepTime(5)
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", -80, 5):
        Util.SleepTime(3)

        # 새로운 빈 행 생성
        new_row = pd.DataFrame([np.nan] * len(dfAdd.columns)).T
        new_row.columns = dfAdd.columns
        # 원하는 위치에 빈 행 삽입
        insert_index = 2  # 3번째 행에 삽입하려면 인덱스는 1입니다.
        dfAdd = pd.concat(
            [dfAdd.iloc[:insert_index], new_row, dfAdd.iloc[insert_index:]]
        ).reset_index(drop=True)

        Util.CsvSave(dfAdd, xlFileAdd)

        # 상품 url
        Util.GoToTheAddressWindow()
        Util.SleepTime(0.5)
        addurl = Util.CopyToClipboardAndGet()
        Util.Debug(f"addurl : {addurl}")
        dfAdd.at[2, COLUMN.B.name] = addurl
        # 크롬 탭 닫기
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(0.5)
        # 상품 번호
        addUrlSplitArray = addurl.split("/")
        if len(addUrlSplitArray) > 0:
            dfAdd.at[2, COLUMN.A.name] = addUrlSplitArray[-1]

        dfAdd.at[2, COLUMN.C.name] = url
        # 상품명 기재
        dfAdd.at[2, COLUMN.T.name] = title
        # 가격
        dfAdd.at[2, COLUMN.U.name] = useMoney
        # 브랜드
        brand = dfAddBefore.at[dfAddBefore.index[0], COLUMN.A.name]
        dfAdd.at[2, COLUMN.E.name] = brand

        # 색이름 리스트 값
        colorNames = []
        for item in arraySizesAndImgUrls:
            colorNames.append(item[Util.Array_ColroName])
        str_saveColorList = Util.JoinArrayToString(colorNames)
        Util.Debug(f"str_saveColorList : {str_saveColorList}")

        dfAdd.at[2, COLUMN.F.name] = str_saveColorList

        # 색 이름 과 사아즈 리스트 값(이중 배열)
        str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
        Util.Debug(f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}")

        dfAdd.at[2, COLUMN.G.name] = str_saveColorNameDoubleArray

        System.xl_J_(dfAdd, 2, "신규 등록", True)

        Util.CsvSave(dfAdd, xlFileAdd)

        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        Util.TelegramSend(
            f"등록한 스토어 주소 : {addurl} 구매 url: {url} title : {title}  useMoney : {useMoney}"
        )

        returnValue = AddOneProduct_Data_Ugg()
        returnValue.addCount = True
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue
    else:
        Util.TelegramSend("++++++++++++++ 이름이 입력 안됬음 왜지??")
        # 이름이 입력 안됬음  왜지??
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\취소", 5, 5)
        Util.SleepTime(1)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품취소 유실 확인", 5, 5)
        Util.SleepTime(1)

        returnValue = AddOneProduct_Data_Ugg()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = False
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

def AddOneProduct_Common(
    url, data, dfAddBefore, dfAdd, xlFileAddBefore, xlFileAdd, addOneProductSuccess, firstName, category
) -> AddOneProduct_Data_Common:
    # UGG에 사이즈 정보로 정보 취합
    useMoney = data.useMoney

    # 이중 배열
    arraySizesAndImgUrls = data.arraySizesAndImgUrls
    
    # 상품 이름
    title = f"{firstName} {data.title}"
    filtered_rows = dfAdd[dfAdd[System.COLUMN.T.name] == title]
    if len(filtered_rows) >= 1:
        if len(arraySizesAndImgUrls) >= 1:
            title += "(" + data.arraySizesAndImgUrls[0][Util.Array_ColroName] + ")"
        count = 2
        while True:
            filtered_rows = dfAdd[dfAdd[System.COLUMN.T.name] == title]
            if len(filtered_rows) >= 1:
                # 1개 이상의 행이 일치합니다.
                title += f"v{count}"
                count += 1
            else:
                # 일치하는 행이 없습니다.
                break

    Util.FolderToDelete(EnvData.g_DefaultPath() + r"\DownloadImage")

    if len(arraySizesAndImgUrls) == 0:
        Util.TelegramSend(f"len(arraySizesAndImgUrls) == 0 url : {url}")
        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        returnValue = AddOneProduct_Data_Common()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

    imgCount: int = 0
    if len(arraySizesAndImgUrls) >= 1:
        imgUrls = arraySizesAndImgUrls[0][Util.Array_UrlList]
        for i in range(len(imgUrls)):
            if i < 10: # 스마스 스토어 사진 추가 이미지 9까지만 밖에 안되서
                if True == Util.DownloadImageUrl(imgUrls[i], i, 750, 1000):
                    imgCount += 1

    if imgCount == 0:
        Util.TelegramSend(f"imgCount == 0 url : {url}")
        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        returnValue = AddOneProduct_Data_Common()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

    Util.SleepTime(1)
    webbrowser.open("https://sell.smartstore.naver.com/#/products/create")
    # 전에 있던 탭 창 삭제
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
        # 전에 있던 탭 창 삭제
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
    ProductRegistration.ProductCategory(
        dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.B.name]
    )

    # 상품명 입력
    ProductRegistration.ProductTitle(title)

    # 판매가 입력
    UpdateAndReturnSalePrice(data.korMony)

    # 옵션 세팅
    if True:
        # 관세 부가 여부 체크
        is_customsDuty = useMoney >= 200

        # 옵션 엑셀 세팅
        Util.SetExcelOption(arraySizesAndImgUrls, is_customsDuty)

        # 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
        UpdateOptionsFromExcel(is_customsDuty)

    # 이미지 등록(대표, 추가)
    ProductRegistration.IamgeRegistration_v2(imgCount)

    # HTML 으로 등록
    SetHTML(arraySizesAndImgUrls, data.details, data.htmlEndAdd, True)

    Util.SleepTime(1)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\저장하기", 5, 5)
    Util.SleepTime(5)
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품관리", -80, 5):
        Util.SleepTime(3)

        # 새로운 빈 행 생성
        new_row = pd.DataFrame([np.nan] * len(dfAdd.columns)).T
        new_row.columns = dfAdd.columns
        # 원하는 위치에 빈 행 삽입
        insert_index = 2  # 3번째 행에 삽입하려면 인덱스는 1입니다.
        dfAdd = pd.concat(
            [dfAdd.iloc[:insert_index], new_row, dfAdd.iloc[insert_index:]]
        ).reset_index(drop=True)

        Util.CsvSave(dfAdd, xlFileAdd)

        # 상품 url
        Util.GoToTheAddressWindow()
        Util.SleepTime(0.5)
        addurl = Util.CopyToClipboardAndGet()
        Util.Debug(f"addurl : {addurl}")
        dfAdd.at[2, COLUMN.B.name] = addurl
        # 크롬 탭 닫기
        Util.KeyboardKeyHotkey("ctrl", "w")
        Util.SleepTime(0.5)
        # 상품 번호
        addUrlSplitArray = addurl.split("/")
        if len(addUrlSplitArray) > 0:
            dfAdd.at[2, COLUMN.A.name] = addUrlSplitArray[-1]

        dfAdd.at[2, COLUMN.C.name] = url
        # 상품명 기재
        dfAdd.at[2, COLUMN.T.name] = title
        # 가격
        dfAdd.at[2, COLUMN.U.name] = useMoney
        # 카테고리
        dfAdd.at[2, COLUMN.V.name] = category
        # 브랜드
        brand = dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.A.name]
        dfAdd.at[2, COLUMN.E.name] = brand

        # 색이름 리스트 값
        colorNames = []
        for item in arraySizesAndImgUrls:
            colorNames.append(item[Util.Array_ColroName])
        str_saveColorList = Util.JoinArrayToString(colorNames)
        Util.Debug(f"str_saveColorList : {str_saveColorList}")

        dfAdd.at[2, COLUMN.F.name] = str_saveColorList

        # 색 이름 과 사아즈 리스트 값(이중 배열)
        str_saveColorNameDoubleArray = Util.DoubleArrayToString(arraySizesAndImgUrls)
        Util.Debug(f"str_saveColorNameDoubleArray : {str_saveColorNameDoubleArray}")

        dfAdd.at[2, COLUMN.G.name] = str_saveColorNameDoubleArray

        System.xl_J_(dfAdd, 2, "신규 등록", True)

        Util.CsvSave(dfAdd, xlFileAdd)

        # 등록해야 될 것에서 삭제
        dfAddBefore = dfAddBefore.iloc[1:]
        Util.CsvSave(dfAddBefore, xlFileAddBefore)

        Util.TelegramSend(
            f"등록한 스토어 주소 : {addurl} 구매 url: {url} title : {title}  useMoney : {useMoney}"
        )

        returnValue = AddOneProduct_Data_Common()
        returnValue.addCount = True
        returnValue.addOneProductSuccess = True
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue
    else:
        Util.TelegramSend("++++++++++++++ 이름이 입력 안됬음 왜지??")
        # 이름이 입력 안됬음  왜지??
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\취소", 5, 5)
        Util.SleepTime(1)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\상품취소 유실 확인", 5, 5)
        Util.SleepTime(1)

        returnValue = AddOneProduct_Data_Common()
        returnValue.addCount = False
        returnValue.addOneProductSuccess = False
        returnValue.dfAddBefore = dfAddBefore
        returnValue.dfAdd = dfAdd
        return returnValue

# 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
def AddDataFromExcel_Ugg():
    Util.TelegramSend(
        "추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 시작"
    )

    xlFileAddBefore = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    dfAddBefore = pd.read_csv(xlFileAddBefore, encoding="cp949")
    rowCountAddBefore = dfAddBefore.shape[0]

    xlFileAdd = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    dfAdd = pd.read_csv(xlFileAdd, encoding="cp949")

    krwUsd = Util.KRWUSD()

    # 모두 반복
    addCount = 0
    count = 1
    data = None
    for _ in range(rowCountAddBefore):
        addOneProductSuccess = True
        while True:
            Util.TelegramSend(f"{count}/{rowCountAddBefore}")
            data = AddOneProduct_Ugg(
                dfAddBefore if data is None else data.dfAddBefore,
                dfAdd if data is None else data.dfAdd,
                xlFileAddBefore,
                xlFileAdd,
                addOneProductSuccess,
                krwUsd,
            )
            dfAddBefore = dfAddBefore if data is None else data.dfAddBefore
            addOneProductSuccess = data.addOneProductSuccess
            if addOneProductSuccess:
                if data.addCount:
                    addCount += 1
                count += 1
                break

    Util.TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 끝")

    return addCount

# 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
def AddDataFromExcel_Common(GetData, exchangeRate, firstName):
    Util.TelegramSend(
        "추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 시작"
    )

    xlFileAddBefore = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    dfAddBefore = pd.read_csv(xlFileAddBefore, encoding="cp949")
    rowCountAddBefore = dfAddBefore.shape[0]

    xlFileAdd = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    dfAdd = pd.read_csv(xlFileAdd, encoding="cp949")

    # 모두 반복
    addCount = 0
    count = 1
    data = None
    for _ in range(rowCountAddBefore):
        addOneProductSuccess = True
        while True:
            Util.TelegramSend(f"{count}/{rowCountAddBefore}")
            url = dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.C.name]
            category = dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.B.name]
            data = AddOneProduct_Common(
                url,
                GetData(url, exchangeRate, category),
                dfAddBefore if data is None else data.dfAddBefore,
                dfAdd if data is None else data.dfAdd,
                xlFileAddBefore,
                xlFileAdd,
                addOneProductSuccess,
                firstName,
                category,
            )
            dfAddBefore = dfAddBefore if data is None else data.dfAddBefore
            addOneProductSuccess = data.addOneProductSuccess
            if addOneProductSuccess:
                if data.addCount:
                    addCount += 1
                count += 1
                break

    Util.TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 끝")

    return addCount


# 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
def AddDataFromExcel_BananarePublic():
    Util.TelegramSend(
        "추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 시작"
    )

    xlFileAddBefore = EnvData.g_DefaultPath() + r"\엑셀\추가 할 것들.CSV"
    dfAddBefore = pd.read_csv(xlFileAddBefore, encoding="cp949")
    rowCountAddBefore = dfAddBefore.shape[0]

    xlFileAdd = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    dfAdd = pd.read_csv(xlFileAdd, encoding="cp949")

    krwUsd = Util.KRWUSD()

    # 모두 반복
    addCount = 0
    count = 1
    data = None
    for _ in range(rowCountAddBefore):
        addOneProductSuccess = True
        while True:
            Util.TelegramSend(f"{count}/{rowCountAddBefore}")
            url = dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.C.name]
            category = dfAddBefore.at[dfAddBefore.index[0], COLUMN_Add.B.name]
            data = AddOneProduct_Common(
                url,
                GetData_BananarePublic(url, krwUsd, category),
                dfAddBefore if data is None else data.dfAddBefore,
                dfAdd if data is None else data.dfAdd,
                xlFileAddBefore,
                xlFileAdd,
                addOneProductSuccess,
                firstName_BananarePublic,
                category
            )
            dfAddBefore = dfAddBefore if data is None else data.dfAddBefore
            addOneProductSuccess = data.addOneProductSuccess
            if addOneProductSuccess:
                if data.addCount:
                    addCount += 1
                count += 1
                break

    Util.TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 끝")

    return addCount


def GetData_Ugg(url, exchangeRate, onlyUseMoney=False) -> Data_Ugg:
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    # 1145990은 url의 끝에 .html 전에 있는 값
    productNumber = ""
    urlSplitArray = url.split("/")
    if len(urlSplitArray) > 0:
        productNumber = urlSplitArray[-1].replace(".html", "")
    else:
        productNumber = url

    webbrowser.open(url)
    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
    # WinWait, ugg
    Util.SleepTime(10)
    htmlElementsData: str = System.GetElementsData()
    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(1)

    # 로봇인지 체크
    match = re.search(
        r"geo.captcha",
        htmlElementsData,
    )
    if match:
        returnValue = Data_Ugg()
        returnValue.isCheckRobot = True
        return returnValue

    # 상품 이름
    title = ""
    match = re.search(
        r'<div\s+class\s*=\s*"sticky-toolbar__content">\s*<span>([^<]*)</span>',
        htmlElementsData,
    )
    if match:
        title = match.group(1)
        title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")

    Util.Debug(f"title : {title}")

    # 상품 정보
    details = ""
    match = re.search(
        r'"description":"(.*?)"',
        htmlElementsData,
    )
    if match:
        details = match.group(1)

    Util.Debug(f"details : {details}")

    startPos = 0
    useMoney = 0
    korMony = 0
    # <div class="sticky-toolbar__content"
    match = re.search(r'<div\s+class\s*=\s*"sticky-toolbar__content"', htmlElementsData)
    if match:
        startPos = match.start()
    # aria-labelledby="size" 까지에서 찾기
    match = re.search(r'aria-labelledby\s*=\s*"size"', htmlElementsData[startPos:])
    if match:
        endPos = match.end() + startPos
    if startPos and endPos:
        contentValue = htmlElementsData[startPos:endPos]
        useMoneys = Util.GetRegExMatcheGroup1List(contentValue, r"\$(.+)")
        if len(useMoneys) > 0:
            if(Util.IsFloat(useMoneys[-1].replace(",", ""))):
                useMoney = float(useMoneys[-1].replace(",", ""))
                Util.Debug(f"useMoney : {useMoney}")

        korMony: int = Util.GetKorMony(useMoney, exchangeRate)

        if onlyUseMoney == False:
            urlEndColorNames = Util.GetRegExMatcheGroup1List(
                contentValue,
                r'<span data-attr-value="([^"]*)" class="color-value swatch swatch-circle',
            )
            colorNames = Util.GetRegExMatcheGroup1List(
                contentValue, r'data-attr-color-swatch="[^"]*"\s*title="([^"]*)"'
            )

            if len(urlEndColorNames) == len(colorNames):
                for index in range(len(urlEndColorNames)):
                    # .html?dwvar_1145990_color=BCDR 제일 뒤에 색 정보 적어서 url 열 수 있음
                    # 색 위치로 클릭하는 것보다 url 열는 것이 더 낫다고 생각됨
                    colorUrl = (
                        f"{url}?dwvar_{productNumber}_color={urlEndColorNames[index]}"
                    )

                    Util.Debug(f"urlEndColorNames[{index}] : {urlEndColorNames[index]}")

                    Util.Debug(f"colorNames[{index}] : {colorNames[index]}")

                    webbrowser.open(colorUrl)
                    # "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
                    # WinWait, ugg
                    Util.SleepTime(10)
                    colorUrlHtmlElementsData: str = System.GetElementsData()
                    # Ctrl + W를 눌러 현재 Chrome 탭 닫기
                    Util.KeyboardKeyHotkey("ctrl", "w")
                    Util.SleepTime(1)

                    # 이미지 url 알아오는 것
                    if True:
                        imgBigUrls = []
                        imgUrls = Util.GetRegExMatcheGroup1List(
                            colorUrlHtmlElementsData,
                            r'<img[^>]+data-srcset="([^"]+)"[^>]+>',
                        )
                        for imgUrl in imgUrls:
                            splitArray = imgUrl.split(" , ")
                            if len(splitArray) > 0:
                                # 제일 끝에 것으로 하는 이유는 이미지 제일 큰 Url 이기 때문에
                                imgBigUrl_match = re.search(
                                    r"https:(.+)\.(png|jpg)", splitArray[-1]
                                )
                                if imgBigUrl_match:
                                    imgBigUrls.append(imgBigUrl_match.group(0))

                    # 구매 가능한 사이즈 구하기
                    if True:
                        sizes = []
                        # <div class="sticky-toolbar__content"> 는 뒤에 줄에 이는 곳 부터 찾기
                        match = re.search(
                            r'<div\s+class\s*=\s*"sticky-toolbar__content"',
                            colorUrlHtmlElementsData,
                        )
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
                            match = re.search(
                                r'options-select\s*"\s*value="([^"]*)"', line
                            )
                            if match:
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
                                        if extractedValue.startswith("0") and (
                                            len(extractedValue) < 2
                                            or extractedValue[1] != "."
                                        ):
                                            extractedValue = extractedValue[1:]

                                        sizeData = f"US_{extractedValue}({Util.GetKorSize_Ugg(extractedValue)})"

                                        Util.Debug(f"size : {sizeData}")
                                        sizes.append(sizeData)
                                    else:
                                        sizes.append(defaultExtractedValue)

                    if len(sizes) > 0 and len(imgBigUrls) > 0:
                        arraySizesAndImgUrls.append([colorNames[index][:25], sizes, imgBigUrls])
    returnValue = Data_Ugg()
    returnValue.useMoney = float(useMoney)
    returnValue.korMony = float(korMony)
    returnValue.arraySizesAndImgUrls = arraySizesAndImgUrls
    returnValue.title = Util.TranslateToKorean(title)
    returnValue.isCheckRobot = False
    returnValue.details = Util.TranslateToKorean(details)
    return returnValue

def GetData_Zara(url, exchangeRate, onlyUseMoney=False) -> Data_Zara:
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    htmlElementsData: str = System.GetElementsData_Zara_v2(url)

    # 상품 이름
    title = ""
    match = re.search(
        r'\},"name":"([^"]+)","detail"',
        htmlElementsData,
    )
    if match:
        title = match.group(1)
        title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")

    Util.Debug(f"title : {title}")

    # 세부 사항
    details = ""
    match = re.search(
        r'"\}\],"description":"(.*?)","rawDescription"',
        htmlElementsData,
    )
    if match:
        details = match.group(1)

    Util.Debug(f"details : {details}")

    useMoney = 0
    korMony = 0
    match = re.search(
        r'<div class="product-detail-info__price-amount price">.*?>\$ (.*?)</span>',
        htmlElementsData,
    )
    if match:
        if(Util.IsFloat(match.group(1).replace(",", ""))):
            useMoney = float(match.group(1).replace(",", ""))
            Util.Debug(f"useMoney : {useMoney}")

    korMony: int = Util.GetKorMony(useMoney, exchangeRate)

    if onlyUseMoney == False:
        match = re.search(r'<p class="product-color-extended-name product-detail-info__color" data-qa-qualifier="product-detail-info-color">(.*?) \|', htmlElementsData)
        if match:
            oneColorName = match.group(1)
            matchComingSoon = re.search("<span>Coming soon</span>", htmlElementsData)
            matchOutOfStock = re.search("<span>OUT OF STOCK</span>", htmlElementsData)
            if matchComingSoon and matchOutOfStock:
                sizes: list = []
                sizeDatas: list = Util.GetRegExMatcheGroup1List(
                    htmlElementsData,
                    r'<div class="product-size-info__main-label" data-qa-qualifier="product-size-info-main-label">(\d+)(½)?</div></div></div></div></li>',
                )
                for size in sizeDatas:
                    korSize = Util.GetKorSize_Zara(size)
                    if korSize != 0:
                        sizes.append(
                            f"US_{size}({korSize})"
                        )
                    else:
                        sizes.append(size)
                        
                imgUrls: list = Util.GetRegExMatcheGroup1List(
                    htmlElementsData,
                    r'<img class="media-image__image media__wrapper--media" alt=".*?srcset="(.*?) ',
                )
                imgUrls = list([string.split("&amp;")[0] for string in imgUrls])
                arraySizesAndImgUrls.append([oneColorName[:25], sizes, imgUrls])
        else:
            colorNames: list = Util.GetRegExMatcheGroup1List(
                htmlElementsData,
                r'<span class="screen-reader-text">(.*?)</span></div><',
            )
            for colorName in colorNames:
                if colorName != "":
                    colorUrlHtmlElementsData = System.GetElementsData_Zara_v2(url, colorName)
                    
                    matchComingSoon = re.search("<span>Coming soon</span>", colorUrlHtmlElementsData)
                    matchOutOfStock = re.search("<span>OUT OF STOCK</span>", colorUrlHtmlElementsData)
                    if not matchComingSoon and not matchOutOfStock:
                        sizes: list = []
                        sizeDatas: list = Util.GetRegExMatcheGroup1List(
                            colorUrlHtmlElementsData,
                            r'<div class="product-size-info__main-label" data-qa-qualifier="product-size-info-main-label">(\d+)(½)?</div></div></div></div></li>',
                        )
                        for size in sizeDatas:
                            korSize = Util.GetKorSize_Zara(size)
                            if korSize != 0:
                                sizes.append(
                                    f"US_{size}({korSize})"
                                )
                            else:
                                sizes.append(size)
                        
                        imgUrls: list = Util.GetRegExMatcheGroup1List(
                            colorUrlHtmlElementsData,
                            r'<img class="media-image__image media__wrapper--media" alt=".*?srcset="(.*?) ',
                        )
                        imgUrls = list([string.split("&amp;")[0] for string in imgUrls])
                        arraySizesAndImgUrls.append([colorName[:25], sizes, imgUrls])

    returnValue = Data_Zara()
    returnValue.useMoney = float(useMoney)
    returnValue.korMony = float(korMony)
    returnValue.arraySizesAndImgUrls = arraySizesAndImgUrls
    returnValue.title = Util.TranslateToKorean(title)
    returnValue.isSoldOut = False
    returnValue.details = Util.TranslateToKorean(details)
    returnValue.fabricAndCare = ""
    return returnValue

def GetData_BananarePublic(url, exchangeRate, category, onlyUseMoney=False) -> Data_BananarePublic:
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    # webbrowser.open(url)
    # Util.SleepTime(10)
    # htmlElementsData: str = System.GetElementsData()
    # # Ctrl + W를 눌러 현재 Chrome 탭 닫기
    # Util.KeyboardKeyHotkey("ctrl", "w")
    # Util.SleepTime(1)
    
    htmlElementsData: str = System.GetElementsData_v2(url)

    # 상품 이름
    title = ""
    match = re.search(
        r'\\"productTitle\\":\\"(.*?)\\"',
        htmlElementsData,
    )
    if match:
        title = match.group(1)
        title = title.replace("*", " ").replace("?", " ").replace("<", " ").replace(">", " ").replace("\\", " ")

    Util.Debug(f"title : {title}")

    # 세부 사항
    details = ""
    match = re.search(
        r'"\},"description":"(.*?)#',
        htmlElementsData,
    )
    if match:
        details = match.group(1)

    Util.Debug(f"details : {details}")

    # 패브릭&케어
    fabricAndCare = ""
    match = re.search(
        r'\\"fabric\\":\{\\"bulletAttributes\\":\[\\"(.*?)"',
        htmlElementsData,
    )
    if match:
        fabricAndCare = match.group(1)

    Util.Debug(f"fabricAndCare : {fabricAndCare}")

    useMoney = 0
    korMony = 0
    contentValue = htmlElementsData
    useMoneys = Util.GetRegExMatcheGroup1List(
        contentValue, r'\\"localizedCurrentPrice\\":\\"\$([0-9.]+)\\"'
    )
    if len(useMoneys) > 0:
        if(Util.IsFloat(useMoneys[-1].replace(",", ""))):
            useMoney = float(useMoneys[-1].replace(",", ""))
            Util.Debug(f"useMoney : {useMoney}")

    korMony: int = Util.GetKorMony(useMoney, exchangeRate)
    
    htmlEndAdd = "<div><span>-</span></div>"
    topCategories = ["여성의류 블라우스", "여성의류 재킷", "여성의류 코트", "여성의류 조끼", "여성의류 티셔츠", "여성의류 정장세트"]
    if " 여성신발 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">US</th><th style="text-align: center;">한국</th></tr><tr bgcolor="#ffffff" align="center"><td>5   </td><td>210</td></tr><tr bgcolor="#ffffff" align="center"><td>5.5 </td><td>215</td></tr><tr bgcolor="#ffffff" align="center"><td>6   </td><td>220</td></tr><tr bgcolor="#ffffff" align="center"><td>6.5 </td><td>225</td></tr><tr bgcolor="#ffffff" align="center"><td>7   </td><td>230</td></tr><tr bgcolor="#ffffff" align="center"><td>7.5 </td><td>235</td></tr><tr bgcolor="#ffffff" align="center"><td>8   </td><td>240</td></tr><tr bgcolor="#ffffff" align="center"><td>8.5 </td><td>245</td></tr><tr bgcolor="#ffffff" align="center"><td>9   </td><td>250</td></tr><tr bgcolor="#ffffff" align="center"><td>9.5 </td><td>255</td></tr><tr bgcolor="#ffffff" align="center"><td>10  </td><td>260</td></tr><tr bgcolor="#ffffff" align="center"><td>10.5</td><td>265</td></tr><tr bgcolor="#ffffff" align="center"><td>11  </td><td>270</td></tr></table>'
    elif " 반지 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">직경 (mm)</th></tr><tr bgcolor="#ffffff" align="center"><td> S </td><td> 5 </td><td> 16 </td></tr><tr bgcolor="#ffffff" align="center"><td> S </td><td> 6 </td><td> 17 </td></tr><tr bgcolor="#ffffff" align="center"><td> M </td><td> 7 </td><td> 18 </td></tr><tr bgcolor="#ffffff" align="center"><td> L </td><td> 8 </td><td> 19 </td></tr></table>'
    elif any(c in category for c in topCategories):
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">가슴 (인치)</th><th style="text-align: center;">허리 (인치)</th><th style="text-align: center;">엉덩이 (인치)</th><th style="text-align: center;">소매 길이 (인치)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>30.5</td><td>24  </td><td>33.5</td><td>30  </td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>31.5</td><td>25  </td><td>34.5</td><td>30  </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>32.5</td><td>26  </td><td>35.5</td><td>31  </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>33.5</td><td>27  </td><td>36.5</td><td>31  </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>34.5</td><td>28  </td><td>37.5</td><td>31.5</td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>35.5</td><td>29  </td><td>38.5</td><td>31.5</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>36.5</td><td>30  </td><td>39.5</td><td>32  </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>37.5</td><td>31  </td><td>40.5</td><td>32  </td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>39  </td><td>32.5</td><td>42  </td><td>32.5</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>40.5</td><td>34  </td><td>43.5</td><td>32.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>42.5</td><td>36  </td><td>45.5</td><td>33.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>44.5</td><td>38  </td><td>47.5</td><td>33.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>47.5</td><td>41  </td><td>50.5</td><td>34  </td></tr></table>'
        htmlEndAdd += "<div><span>-</span></div>"
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">가슴 (cm)</th><th style="text-align: center;">허리 (cm)</th><th style="text-align: center;">엉덩이 (cm)</th><th style="text-align: center;">소매 길이 (cm)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>77.5 </td><td>61  </td><td>85   </td><td>76  </td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>80   </td><td>63.5</td><td>87.5 </td><td>76  </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>82.5 </td><td>66  </td><td>90   </td><td>79  </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>85   </td><td>68.5</td><td>93   </td><td>79  </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>87.5 </td><td>71  </td><td>95   </td><td>80  </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>90   </td><td>73.5</td><td>98   </td><td>80  </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>92.5 </td><td>76  </td><td>100  </td><td>81  </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>95   </td><td>79  </td><td>103  </td><td>81  </td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>99   </td><td>82.5</td><td>106.5</td><td>82.5</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>103  </td><td>86.5</td><td>110.5</td><td>82.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>108  </td><td>91.5</td><td>115.5</td><td>85  </td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>113  </td><td>96.5</td><td>120.5</td><td>85  </td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>120.5</td><td>104 </td><td>128  </td><td>86.5</td></tr></table>'
    elif "여성의류 원피스" in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">가슴 (인치)</th><th style="text-align: center;">허리 (인치)</th><th style="text-align: center;">엉덩이 (인치)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>30.5</td><td>24  </td><td>33.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>31.5</td><td>25  </td><td>34.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>32.5</td><td>26  </td><td>35.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>33.5</td><td>27  </td><td>36.5</td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>34.5</td><td>28  </td><td>37.5</td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>35.5</td><td>29  </td><td>38.5</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>36.5</td><td>30  </td><td>39.5</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>37.5</td><td>31  </td><td>40.5</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>39  </td><td>32.5</td><td>42  </td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>40.5</td><td>34  </td><td>43.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>42.5</td><td>36  </td><td>45.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>44.5</td><td>38  </td><td>47.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>46.5</td><td>40  </td><td>49.5</td></tr></table>'
        htmlEndAdd += "<div><span>-</span></div>"
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">가슴 (cm)</th><th style="text-align: center;">허리 (cm)</th><th style="text-align: center;">엉덩이 (cm)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>77 </td><td>61 </td><td>95 </td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>80 </td><td>64 </td><td>98 </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>83 </td><td>66 </td><td>85 </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>85 </td><td>69 </td><td>90 </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>88 </td><td>71 </td><td>93 </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>90 </td><td>74 </td><td>88 </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>93 </td><td>76 </td><td>100</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>95 </td><td>79 </td><td>103</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>99 </td><td>83 </td><td>107</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>103</td><td>86 </td><td>110</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>108</td><td>91 </td><td>116</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>113</td><td>97 </td><td>121</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>118</td><td>102</td><td>126</td></tr></table>'
    elif "여성의류 바지" in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">수치</th><th style="text-align: center;">허리(인치)</th><th style="text-align: center;">허리(cm)</th><th style="text-align: center;">엉덩이(인치)</th><th style="text-align: center;">엉덩이(cm)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>23</td><td>24  </td><td>61  </td><td>33.5</td><td>85   </td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>24</td><td>25  </td><td>63.5</td><td>34.5</td><td>87.5 </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>25</td><td>26  </td><td>66  </td><td>35.5</td><td>90   </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>26</td><td>27  </td><td>68.5</td><td>36.5</td><td>93   </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>27</td><td>28  </td><td>71  </td><td>37.5</td><td>95   </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>28</td><td>29  </td><td>73.5</td><td>38.5</td><td>98   </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>29</td><td>30  </td><td>76  </td><td>39.5</td><td>100  </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>30</td><td>31  </td><td>79  </td><td>40.5</td><td>103  </td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>31</td><td>32.5</td><td>82  </td><td>42  </td><td>106.5</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>32</td><td>34  </td><td>86.5</td><td>43.5</td><td>110.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>33</td><td>36  </td><td>91.5</td><td>45.5</td><td>115.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>34</td><td>38  </td><td>96.5</td><td>47.5</td><td>120.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>35</td><td>41  </td><td>104 </td><td>50.5</td><td>128  </td></tr> </table>'
    elif "여성의류 스커트" in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">허리 (인치)</th><th style="text-align: center;">엉덩이 (인치)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>24  </td><td>33.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>25  </td><td>34.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>26  </td><td>35.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>27  </td><td>36.5</td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>28  </td><td>37.5</td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>29  </td><td>38.5</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>30  </td><td>39.5</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>31  </td><td>40.5</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>32.5</td><td>42  </td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>34  </td><td>43.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>36  </td><td>45.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>38  </td><td>47.5</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>40  </td><td>49.5</td></tr></table>'
        htmlEndAdd += "<div><span>-</span></div>"
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">허리 (cm)</th><th style="text-align: center;">엉덩이 (cm)</th></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>000</td><td>61 </td><td>85 </td></tr><tr bgcolor="#ffffff" align="center"><td>XXS</td><td>00 </td><td>64 </td><td>88 </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>0  </td><td>66 </td><td>90 </td></tr><tr bgcolor="#ffffff" align="center"><td>XS </td><td>2  </td><td>69 </td><td>93 </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>4  </td><td>71 </td><td>95 </td></tr><tr bgcolor="#ffffff" align="center"><td>S  </td><td>6  </td><td>74 </td><td>98 </td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>8  </td><td>76 </td><td>100</td></tr><tr bgcolor="#ffffff" align="center"><td>M  </td><td>10 </td><td>79 </td><td>103</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>12 </td><td>83 </td><td>107</td></tr><tr bgcolor="#ffffff" align="center"><td>L  </td><td>14 </td><td>86 </td><td>110</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>16 </td><td>91 </td><td>116</td></tr><tr bgcolor="#ffffff" align="center"><td>XL </td><td>18 </td><td>97 </td><td>121</td></tr><tr bgcolor="#ffffff" align="center"><td>XXL</td><td>20 </td><td>102</td><td>126</td></tr></table>'
    elif " 벨트 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">바지 벨트 (중앙 구멍)(인치)</th><th style="text-align: center;">바지 벨트 (조절 범위)(인치)</th><th style="text-align: center;">허리 벨트 (중앙 구멍)(인치)</th><th style="text-align: center;">허리 벨트 (조절 범위)(인치)</th></tr><tr bgcolor="#ffffff" align="center"><td> XXS </td><td> 0-2  </td><td> 29 </td><td> 27-31 </td><td> 27 </td><td> 25-29 </td></tr><tr bgcolor="#ffffff" align="center"><td> XS  </td><td> 4-6  </td><td> 31 </td><td> 29-33 </td><td> 29 </td><td> 27-31 </td></tr><tr bgcolor="#ffffff" align="center"><td> S   </td><td> 8-10 </td><td> 33 </td><td> 31-35 </td><td> 31 </td><td> 29-33 </td></tr><tr bgcolor="#ffffff" align="center"><td> M   </td><td> 12-14</td><td> 35 </td><td> 33-37 </td><td> 33 </td><td> 31-35 </td></tr><tr bgcolor="#ffffff" align="center"><td> L   </td><td> 16-18</td><td> 38 </td><td> 36-40 </td><td> 36 </td><td> 34-38 </td></tr><tr bgcolor="#ffffff" align="center"><td> XL  </td><td> 20   </td><td> 42 </td><td> 40-44 </td><td> 40 </td><td> 38-42 </td></tr><tr bgcolor="#ffffff" align="center"><td> XXL </td><td> 22   </td><td> 46 </td><td> 44-48 </td><td> 44 </td><td> 42-46 </td></tr></table>'
        htmlEndAdd += "<div><span>-</span></div>"
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">크기</th><th style="text-align: center;">바지 벨트 (중앙 구멍)(cm)</th><th style="text-align: center;">바지 벨트 (조절 범위)(cm)</th><th style="text-align: center;">허리 벨트 (중앙 구멍)(cm)</th><th style="text-align: center;">허리 벨트 (조절 범위)(cm)</th></tr><tr bgcolor="#ffffff" align="center"><td> XXS </td><td> 0-2   </td><td> 74  </td><td> 69-79   </td><td> 69  </td><td> 64-74   </td></tr><tr bgcolor="#ffffff" align="center"><td> XS  </td><td> 4-6   </td><td> 79  </td><td> 74-84   </td><td> 74  </td><td> 69-79   </td></tr><tr bgcolor="#ffffff" align="center"><td> S   </td><td> 8-10  </td><td> 84  </td><td> 79-89   </td><td> 79  </td><td> 74-84   </td></tr><tr bgcolor="#ffffff" align="center"><td> M   </td><td> 12-14 </td><td> 89  </td><td> 84-94   </td><td> 84  </td><td> 79-89   </td></tr><tr bgcolor="#ffffff" align="center"><td> L   </td><td> 16-18 </td><td> 96  </td><td> 91-102  </td><td> 91  </td><td> 86-96   </td></tr><tr bgcolor="#ffffff" align="center"><td> XL  </td><td> 20    </td><td> 107 </td><td> 102-112 </td><td> 102 </td><td> 96-107  </td></tr><tr bgcolor="#ffffff" align="center"><td> XXL </td><td> 22    </td><td> 117 </td><td> 112-122 </td><td> 112 </td><td> 107-117 </td></tr></table>'
    elif " 모자 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">둘레 (인치)</th><th style="text-align: center;">둘레 (cm)</th></tr><tr bgcolor="#ffffff" align="center"><td> S </td><td> 22 </td><td> 55.88 </td></tr><tr bgcolor="#ffffff" align="center"><td> M </td><td> 23 </td><td> 58.42 </td></tr><tr bgcolor="#ffffff" align="center"><td> L </td><td> 24 </td><td> 60.97 </td></tr></table>'
    elif " 장갑 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">길이 (인치)</th><th style="text-align: center;">길이 (cm)</th></tr><tr bgcolor="#ffffff" align="center"><td> S </td><td> 8 ⅛ </td><td> 21 </td></tr><tr bgcolor="#ffffff" align="center"><td> M </td><td> 8 ½ </td><td> 22 </td></tr><tr bgcolor="#ffffff" align="center"><td> L </td><td> 8 ⅞ </td><td> 23 </td></tr> </table>'
    elif " 타이즈 " in category:
        htmlEndAdd += '<table border="1" cellpadding="10" cellspacing="0" width="50%" align="center" style="margin-left: auto; margin-right: auto;"><tr bgcolor="#f2f2f2"><th style="text-align: center;">크기</th><th style="text-align: center;">키 (Ft)</th><th style="text-align: center;">몸무게 (Lbs)</th><th style="text-align: center;">키 (cm)</th><th style="text-align: center;">몸무게 (kg)</th></tr><tr bgcolor="#ffffff" align="center"><td> PETITE </td><td> 4.8 - 5.25 </td><td> 90 - 125  </td><td> 147 - 160 </td><td> 41 - 57  </td></tr><tr bgcolor="#ffffff" align="center"><td> S/M    </td><td> 5 - 5.6    </td><td> 90 - 145  </td><td> 152 - 170 </td><td> 41 - 66  </td></tr><tr bgcolor="#ffffff" align="center"><td> M/L    </td><td> 5.6+       </td><td> 145 - 175 </td><td> 170+      </td><td> 66 - 79  </td></tr><tr bgcolor="#ffffff" align="center"><td> TALL   </td><td> 5.6 - 6    </td><td> 175 - 245 </td><td> 170 - 183 </td><td> 79 - 111 </td></tr></table>'
    
    if onlyUseMoney == False:
        matcheGroup1And2 = Util.GetRegExMatcheGroup1And2List(
            contentValue,
            r'\\"businessCatalogItemId\\":\\"(\d+)\\".*?\\"colorName\\":\\"(.*?)\\"',
        )
        # 중복 제거하면서 순서 유지
        colorNames = [item[1] for item in matcheGroup1And2]
        for i in range(len(matcheGroup1And2)):
            colorName = matcheGroup1And2[i][1]
            count = colorNames.count(colorName)
            if (
                count != 0
            ):
                colorNames = [value for value in colorNames if value != colorName]
                productNumber = matcheGroup1And2[i][0]
                colorUrlHtmlElementsData: str = System.GetElementsData_v2(f"https://bananarepublic.gap.com/browse/product.do?pid={productNumber}")

                sizes: list = []
                match = re.search("Size:One Size", colorUrlHtmlElementsData)
                if match:
                    sizes.append("One Size")
                else:
                    sizeDatas: list = Util.GetRegExMatcheGroup1List(
                        colorUrlHtmlElementsData,
                        r'aria-label="Size:(.*?)"',
                    )
                    for i in range(len(sizeDatas)):
                        sizeData = sizeDatas[i]
                        if "out of stock" not in sizeData:
                            if " 여성신발 " in category:
                                korSize = Util.GetKorSize_BananarePublic(sizeData)
                                if korSize != 0:
                                    sizes.append(
                                        f"US_{sizeData}({korSize})"
                                    )
                                else:
                                    sizes.append(sizeData)
                            else:
                                sizes.append(sizeData)

                imgBigUrls: list = []
                matchs: list = Util.GetRegExMatcheGroup1List(
                    colorUrlHtmlElementsData,
                    r'" src="/(.*?).jpg" width=',
                )
                for i in range(len(matchs)):
                    imgBigUrls.append(f"https://bananarepublic.gap.com/{matchs[i]}.jpg")
                arraySizesAndImgUrls.append([colorName[:25], sizes, imgBigUrls])

    returnValue = Data_BananarePublic()
    returnValue.useMoney = float(useMoney)
    returnValue.korMony = float(korMony)
    returnValue.arraySizesAndImgUrls = arraySizesAndImgUrls
    returnValue.title = Util.TranslateToKorean(title)
    returnValue.isSoldOut = False
    returnValue.details = Util.TranslateToKorean(details)
    returnValue.fabricAndCare = Util.TranslateToKorean(fabricAndCare)
    returnValue.htmlEndAdd = htmlEndAdd
    return returnValue

def GetData_Mytheresa(url, exchangeRate) -> Data_Mytheresa:
    # 사이즈 정보로 정보 취합
    useMoney = 0
    korMony = 0

    # 이중 배열
    arraySizesAndImgUrls = []

    # 상품 이름
    title: str = ""

    # 상세 정보
    details: str = ""

    isSoldOut = False

    if True:
        # webbrowser.open(url)
        # # "mytheresa"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
        # # WinWait, mytheresa
        # Util.SleepTime(10)
        # htmlElementsData: str = System.GetElementsData()
        # # Ctrl + W를 눌러 현재 Chrome 탭 닫기
        # Util.KeyboardKeyHotkey("ctrl", "w")
        # Util.SleepTime(1)

        # match = re.search(r'PriceSpecification": {\s+"price": (\d+)', htmlElementsData)
        # if match:
        #     useMoney = float(match.group(1))
            
        # v2 일때
        htmlElementsData: str = System.GetElementsData_v2(url)
        
        match = re.search(r">Sold Out<", htmlElementsData)
        if match:
            isSoldOut = True

        if isSoldOut == False:
            match = re.search(r"error404__title", htmlElementsData)
            if match:
                isSoldOut = True
            
        if isSoldOut == False:
            # HTML 파서를 사용하여 파싱
            parser = etree.HTMLParser()
            tree = etree.fromstring(htmlElementsData, parser)
            # XPath를 사용하여 타이틀 요소를 선택하고 텍스트를 추출
            if tree is not None:
                if tree.xpath("//title/text()"):
                    title = tree.xpath("//title/text()")[0]
                else:
                    isSoldOut = True
            else:
                isSoldOut = True
            title = str(title).replace("| Mytheresa", "")
            
        if isSoldOut == False:
            # 할인 일때 먼저 체크 하고 
            match = re.search(r'"pricing__prices__value pricing__prices__value--discount"><span class="pricing__prices__price"> <!-- -->€ (.*?)</span></span></div><div', htmlElementsData)
            if match:
                if(Util.IsFloat(match.group(1).replace(",", ""))):
                    useMoney = float(match.group(1).replace(",", ""))
            else:
                match = re.search(r'"productinfo__price"><div class="pricing">.*?<!-- -->€ (.*?)<', htmlElementsData)
                if match:
                    if(Util.IsFloat(match.group(1).replace(",", ""))):
                        useMoney = float(match.group(1).replace(",", ""))

            korMony: int = Util.GetKorMony(float(useMoney), float(exchangeRate))

            # 세부 사항
            details = ""
            match = re.search(
                r'"description": "(.*?)"',
                htmlElementsData,
            )
            if match:
                details = match.group(1)

            Util.Debug(f"details : {details}")

            # 재화 타입
            match = re.search(r'"priceCurrency":\s*"([A-Z]{3})"', htmlElementsData)
            if match:
                priceCurrency = match.group(1)

            sizeLines = Util.GetRegExMatcheGroup1List(
                htmlElementsData, r'<span class="sizeitem__label">(.*?)</span>'
            )

            sizes = []
            for sizeLine in sizeLines:
                sizeData = sizeLine + f"({Util.GetKorSize_Mytheresa(sizeLine)})"
                Util.Debug(f"size : {sizeData}")
                sizes.append(sizeData)

            colorName = "One Color"
            imgBigUrls = []
            arraySizesAndImgUrls.append([colorName, sizes, imgBigUrls])

    returnValue = Data_Mytheresa()
    returnValue.useMoney = float(useMoney)
    returnValue.korMony = float(korMony)
    returnValue.arraySizesAndImgUrls = arraySizesAndImgUrls
    returnValue.title = Util.TranslateToKorean(title)
    # returnValue.sizesLength = len(sizes)
    returnValue.isSoldOut = isSoldOut
    returnValue.details = Util.TranslateToKorean(details)
    return returnValue

# 스마트 스토어 수정 화면까지 이동
def ManageAndModifyProducts(df, row) -> ManageAndModifyProductsData:
    values = ManageAndModifyProductsData()
    values.isNoProduct = False
    values.isNoNetwork = False

    Util.SleepTime(1)
    webbrowser.open("https://sell.smartstore.naver.com/#/products/origin-list")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "tab")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "w")
    Util.SleepTime(2)

    # 확인은 오류가 발생했습니다 라는 팝업이 나오는 경우가 있어서
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\확인", 5, 5, 1):
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
        pyperclip.copy(round(float(df.at[row, COLUMN.A.name])))
        Util.SleepTime(0.5)
        # 상품번호 붙여넣기
        Util.KeyboardKeyHotkey("ctrl", "v")
        Util.SleepTime(0.5)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 조회\검색", 0, 0)
        Util.SleepTime(1)
        if False == Util.ClickAtWhileFoundImage(
            r"스마트 스토어\상품 조회\수정", 0, 0, 5
        ):
            values.isNoProduct = True
            return values
        Util.SleepTime(2)

    # "스마트 스토어\네트워크 불안정 느낌표"
    if Util.ClickAtWhileFoundImage(r"스마트 스토어\네트워크 불안정", 0, 0, 1):
        values.isNoNetwork = True
        return values

    if Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\KC인증", 0, 0, 2):
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\KC인증 닫기", 10, 10, 2)

    # 안전기준 팝업 등이 나오면 끄기 위함
    Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\확인", 5, 5, 1)
    Util.SleepTime(0.5)

    return values


# 품절
def SoldOut(df, row):
    # 품절
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\옵션", 0, 0, -500)
    Util.SleepTime(1)
    Util.WheelAndMoveAtWhileFoundImage(
        r"스마트 스토어\상품 수정\옵션에서 선택형", 0, 0, -500
    )
    Util.SleepTime(1)
    optionEnd = pyautogui.position()
    Util.SleepTime(0.5)
    Util.MoveAtWhileFoundImage(
        r"스마트 스토어\상품 수정\옵션", 0, 0, 2, 1, optionEnd.x - 50, optionEnd.y - 150
    )
    Util.SleepTime(1)
    optionStart = pyautogui.position()
    Util.SleepTime(0.5)
    Util.MouseMove(optionEnd.x + 1000, optionEnd.y + 100)
    Util.ClickAtWhileFoundImage(
        r"스마트 스토어\상품 수정\옵션에서 선택형에서 설정함 상태",
        0,
        0,
        2,
        1,
        optionStart.x,
        optionStart.y,
        optionEnd.x + 1000,
        optionEnd.y + 100,
    )
    Util.WheelAndMoveAtWhileFoundImage(
        r"스마트 스토어\상품 수정\재고수량에 개", 0, 0, 500
    )
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

    System.xl_J_(df, row, "품절 상태로 변경 완료", True)

    Util.TelegramSend(f"품절 상태로 변경 완료 row({row}) ")


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
    pyperclip.copy(int(korMony))
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)


# 품절 완료 된 것 엑셀에서 제거
def RemoveCompletedSoldOutItems(df):
    startCount = 2
    while True:
        isFind = False
        for i in range(startCount, len(df)):
            if "품절 상태로 변경 완료" in df.at[i, COLUMN.J]:
                # 행 삭제
                df.drop(index=i, inplace=True)
                isFind = True
                startCount = i
                break

        if not isFind:
            break

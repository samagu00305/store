from datetime import datetime
import Util, EnvData, GlobalData
import urllib.request
import numpy as np
import webbrowser, requests, subprocess, pyautogui, time, pyperclip, re, os, cv2, openpyxl
from enum import Enum
import tkinter as tk
import threading
import psutil
from googletrans import Translator

Array_ColroName = 0
Array_SizeList = 1
Array_UrlList = 2


# 현재 미국 환율 정보 출력
def KRWUSD():
    response = requests.get(
        "https://quotation-api-cdn.dunamu.com/v1/forex/recent?codes=FRX.KRWUSD"
    )

    if response.status_code == 200:
        basePrice = re.search(r'"basePrice":(\d+\.\d+)', response.text)
        if basePrice:
            value = float(basePrice.group(1))
            return value
        else:
            Util.TelegramSend(
                "******** Error  ---- KRWUSD() ==== 0  ----  Error  Value not found in response"
            )
            return 0
    else:
        Util.TelegramSend(
            f"******** Error  ---- KRWUSD() ==== 0  ----  Error response.status_code : {response.status_code}"
        )
        return 0


# 현재 유료 환율 정보 출력
def KRWEUR():
    response = requests.get(
        "https://quotation-api-cdn.dunamu.com/v1/forex/recent?codes=FRX.KRWEUR"
    )
    if response.status_code == 200:
        basePrice = re.search(r'"basePrice":(\d+\.\d+)', response.text)
        if basePrice:
            value = float(basePrice.group(1))
            return value
        else:
            Util.TelegramSend(
                "******** Error  ---- KRWUSD() ==== 0  ----  Error  Value not found in response"
            )
            return 0
    else:
        Util.TelegramSend(
            f"******** Error  ---- KRWUSD() ==== 0  ----  Error response.status_code : {response.status_code}"
        )
        return 0


def MouseMove(x, y):
    pyautogui.moveTo(x, y)


# 현재 위치에서 마우스 클릭
def NowMouseClick():
    pyautogui.click()


# 현재 위치에서 마우스 오른쪽 클릭
def NowMouseClickRight():
    pyautogui.click(button="right")


def KeyboardKeyPress(key):
    pyautogui.press(key)


def KeyboardKeyHotkey(key1, key2):
    pyautogui.hotkey(key1, key2)


def MouseWheelScroll(wheelMove):
    pyautogui.scroll(wheelMove)


class Size:
    def __init__(self):
        self.width = ""
        self.height = ""


def ScreenSize():
    screenSize = pyautogui.size()
    returnValue = Size()
    returnValue.width = screenSize.width
    returnValue.height = screenSize.height
    return returnValue


def FoundImage(imageName):
    return (
        Util.FindImage_Byref(imageName).resultType
        == Enum_FIND_IMAGE_RESULT_TYPE.Success
    )


# 이미지 찾을 때 까지 대기
def WhileFoundImage(
    imageName, findMaxCount=10, delayTime=1, searchStart_x=0, searchStart_y=0
):
    findCount = 0
    while True:
        if (
            Util.FindImage_Byref(imageName, searchStart_x, searchStart_y).resultType
            == Enum_FIND_IMAGE_RESULT_TYPE.Success
        ):
            Util.SleepTime(0.5)
            return True
        Util.SleepTime(delayTime)
        findCount += 1
        if findMaxCount <= findCount:
            Util.Debug(
                f"Error WhileFoundImage - {findMaxCount}번을 찾았으나 실패했습니다. imageName : {imageName}"
            )
            return False


# 이미지 찾아서 마우스 이동
def MoveAtWhileFoundImage(
    imageName,
    addX=0,
    addY=0,
    findMaxCount=10,
    delayTime=1,
    searchStart_x=0,
    searchStart_y=0,
):
    findCount = 0
    while True:  # 무한 루프
        findImageResult = Util.FindImage_Byref(imageName, searchStart_x, searchStart_y)
        if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            Util.Debug(f"x:{findImageResult.x + addX}  y{findImageResult.y + addY}")
            Util.MouseMove(findImageResult.x + addX, findImageResult.y + addY)
            Util.SleepTime(0.5)
            return True
        Util.SleepTime(delayTime)
        findCount += 1
        if findMaxCount <= findCount:
            Util.Debug(
                f"Error ClickAtWhileFoundImage - {findMaxCount}번을 찾았으나 실패했습니다. imageName : {imageName}"
            )
            return False


# 이미지 찾아서 클릭
def ClickAtWhileFoundImage(
    imageName,
    addX=0,
    addY=0,
    findMaxCount=10,
    delayTime=1,
    searchStart_x=0,
    searchStart_y=0,
    searchEnd_x=-1,
    searchEnd_y=-1,
) -> bool:
    findCount = 0
    while True:  # 무한 루프
        findImageResult = Util.FindImage_Byref(
            imageName, searchStart_x, searchStart_y, searchEnd_x, searchEnd_y
        )
        if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            Util.MouseMove(findImageResult.x + addX, findImageResult.y + addY)
            Util.SleepTime(0.5)
            Util.NowMouseClick()
            return True
        Util.SleepTime(delayTime)
        findCount += 1
        if findMaxCount <= findCount:
            Util.Debug(
                f"Error ClickAtWhileFoundImage - {findMaxCount}번을 찾았으나 실패했습니다. imageName : {imageName}"
            )
            return False


# 이미지 찾아서 더블클릭
def DoubleClickAtWhileFoundImage(imageName, addX=0, addY=0, threshold=0.7):
    while True:
        findImageResult = Util.FindImage_Byref(imageName, threshold=threshold)
        if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX = findImageResult.x + addX
            imagePosY = findImageResult.y + addY
            Util.MouseMove(imagePosX + addX, imagePosY)
            Util.SleepTime(0.5)
            pyautogui.click(x=imagePosX, y=imagePosY, button="left", clicks=2)
            return
        Util.SleepTime(1)


# 이미지 찾아서 Drag
def DragAtFoundImage(imageName1, addX1, addY1, imageName2, addX2, addY2):
    while True:  # 무한 루프
        findImageResult1 = Util.FindImage_Byref(imageName1)
        if findImageResult1.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX = findImageResult1.x + addX1
            imagePosY = findImageResult1.y + addY1
            Util.MouseMove(imagePosX, imagePosY)
            Util.SleepTime(0.5)
            pyautogui.mouseDown(button="left")
            break
        Util.SleepTime(1)

    while True:  # 무한 루프
        findImageResult2 = Util.FindImage_Byref(imageName2)
        if findImageResult2.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX = findImageResult2.x + addX2
            imagePosY = findImageResult2.y + addY2
            Util.MouseMove(imagePosX, imagePosY)
            Util.SleepTime(0.5)
            pyautogui.mouseUp()
            break
        Util.SleepTime(1)


# 아래로 스크롤 하면서 이미지 찾아서 클릭
def WheelAndClickAtWhileFoundImage(
    imageName, addX=0, addY=0, wheelMove=-1000, inputCount=-1
):
    count = 0
    while True:
        findImageResult = Util.FindImage_Byref(imageName)
        if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX = findImageResult.x + addX
            imagePosY = findImageResult.y + addY
            Util.MouseMove(imagePosX, imagePosY)
            Util.SleepTime(0.5)
            pyautogui.click(x=imagePosX, y=imagePosY, button="left", clicks=1)
            return

        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        Util.MouseWheelScroll(wheelMove)

        Util.SleepTime(1)

        if inputCount != -1:
            count += 1
            if count >= inputCount:
                return

            if count == 10:
                Util.TelegramSend(
                    f"******** Error  WheelAndClickAtWhileFoundImage()   imageName : {imageName}"
                )


# 아래로 스크롤 하면서 이미지 찾아서 클릭
def WheelAndClickAtWhileFoundImage_v2(
    imageName_1, imageName_2, addX=0, addY=0, wheelMove=-1000
):
    count = 0
    while True:
        findImageResult1 = Util.FindImage_Byref(imageName_1)
        if findImageResult1.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX_1 = findImageResult1.x + addX
            imagePosY_1 = findImageResult1.y + addY
            Util.MouseMove(imagePosX_1, imagePosY_1)
            Util.SleepTime(0.5)
            Util.NowMouseClick()
            return 1

        findImageResult2 = Util.FindImage_Byref(imageName_2)
        if findImageResult2.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX_2 = findImageResult2.x + addX
            imagePosY_2 = findImageResult2.y + addY
            Util.MouseMove(imagePosX_2, imagePosY_2)
            Util.SleepTime(0.5)
            Util.NowMouseClick()
            return 2

        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        Util.MouseWheelScroll(wheelMove)

        Util.SleepTime(1)

        count += 1
        if count == 10:
            Util.TelegramSend(
                f"******** Error  WheelAndClickAtWhileFoundImage_v2()   imageName_1 : {imageName_1}    imageName_2 : {imageName_2}"
            )

    return 0


# 아래로 스크롤 하면서 이미지 찾아서 이동
def WheelAndMoveAtWhileFoundImage(imageName, addX=0, addY=0, wheelMove=-1000):
    count = 0
    while True:
        findImageResult = Util.FindImage_Byref(imageName)
        if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX = findImageResult.x + addX
            imagePosY = findImageResult.y + addY
            Util.MouseMove(imagePosX, imagePosY)
            Util.SleepTime(0.5)
            return

        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        Util.MouseWheelScroll(wheelMove)

        Util.SleepTime(1)

        count += 1
        if count == 10:
            Util.TelegramSend(
                f"******** Error  WheelAndMoveAtWhileFoundImage()   imageName : {imageName}"
            )


# 아래로 스크롤 하면서 이미지 찾아서 이동
def WheelAndMoveAtWhileFoundImage_v2(
    imageName_1, imageName_2, addX=0, addY=0, wheelMove=-1000
):
    count = 0
    while True:
        findImageResult1 = Util.FindImage_Byref(imageName_1)
        if findImageResult1.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX_1 = findImageResult1.x + addX
            imagePosY_1 = findImageResult1.y + addY
            Util.MouseMove(imagePosX_1, imagePosY_1)
            Util.SleepTime(0.5)
            return 1

        findImageResult2 = Util.FindImage_Byref(imageName_2)
        if findImageResult2.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            imagePosX_2 = findImageResult2.x + addX
            imagePosY_2 = findImageResult2.y + addY
            Util.MouseMove(imagePosX_2, imagePosY_2)
            Util.SleepTime(0.5)
            return 2

        # 스크롤 시작 위치에서 아래로 이동하여 스크롤링
        Util.MouseWheelScroll(wheelMove)

        Util.SleepTime(1)

        count += 1
        if count == 10:
            Util.TelegramSend(
                f"******** Error  WheelAndClickAtWhileFoundImage_v2()   imageName_1 : {imageName_1}    imageName_2 : {imageName_2}"
            )

    return 0


# # 이미지를 화면 전체에서 검색해서 검색 정보를 Byref로 값을 넘김
# def FindImage_Byref(
#     imageName,
#     isFindImage,
#     imagePosX,
#     imagePosY,
#     searchStart_x=0,
#     searchStart_y=0,
#     searchEnd_x=-1,
#     searchEnd_y=-1,
# ):
#     Util.Debug(f"이미지 찾기 시작({imageName})")

#     if searchEnd_x == -1:
#         searchEnd_x = pyautogui.size()[0]  # 화면 너비
#     if searchEnd_y == -1:
#         searchEnd_y = pyautogui.size()[1]  # 화면 높이

#     Util.Debug(
#         f"{imageName} searchStart_x:{searchStart_x} searchStart_y:{searchStart_y} searchEnd_x:{searchEnd_x} searchEnd_y:{searchEnd_y}"
#     )

#     # 이미지 검색
#     image_location = pyautogui.locateOnScreen(
#         imageName + ".png",
#         region=(
#             searchStart_x,
#             searchStart_y,
#             searchEnd_x - searchStart_x,
#             searchEnd_y - searchStart_y,
#         ),
#     )

#     print(
#         f"{imageName} ErrorLevel:{'Error' if not image_location else 'Success'} FoundX:{image_location[0] if image_location else None} FoundY:{image_location[1] if image_location else None}"
#     )

#     if image_location:
#         isFindImage.value = 0  # 성공 시 에러 레벨은 0
#         imagePosX.value = image_location[0]
#         imagePosY.value = image_location[1]
#     else:
#         isFindImage.value = 1  # 실패 시 에러 레벨은 1
#         imagePosX.value = 0
#         imagePosY.value = 0


def FindImage_Byref(
    imageName,
    searchStart_x=0,
    searchStart_y=0,
    searchEnd_x=-1,
    searchEnd_y=-1,
    threshold=0.7,
):
    Util.Debug(f"이미지 찾기 시작({imageName})")

    if searchEnd_x == -1:
        searchEnd_x = Util.ScreenSize().width  # 화면 너비
    if searchEnd_y == -1:
        searchEnd_y = Util.ScreenSize().height  # 화면 높이

    Util.Debug(
        f"{imageName} searchStart_x:{searchStart_x} searchStart_y:{searchStart_y} searchEnd_x:{searchEnd_x} searchEnd_y:{searchEnd_y}"
    )
    image_path = f"{EnvData.g_DefaultPath()}/Image/{imageName}.png"

    # 이미지 검색
    findImageResult = Util.FindImage(
        image_path,
        searchStart_x,
        searchStart_y,
        region=(
            searchStart_x,
            searchStart_y,
            searchEnd_x - searchStart_x,
            searchEnd_y - searchStart_y,
        ),
        threshold=threshold,
    )

    Util.Debug(
        f"{imageName} resultTyp:{findImageResult.resultType.name}"
        + (
            f" FoundX:{findImageResult.x} FoundY:{findImageResult.y}"
            if findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success
            else ""
        )
    )

    return findImageResult


# 이미지 서치해서 있는지 알려주는 함수
def IsImageSearch(imageName, searchStart_x=0, searchStart_y=0):
    Util.Debug(f"IsImageSearch의 이미지 검색 시작 imageName : {imageName}")
    findImageResult = Util.FindImage(
        imageName,
        searchStart_x,
        searchStart_y,
        region=(
            searchStart_x,
            searchStart_y,
            Util.ScreenSize().width,
            Util.ScreenSize().height,
        ),
    )
    Util.Debug(f"{imageName} result: {findImageResult.resultType}")
    return findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success


# 이미지 서치해서 있는지 알려주는 함수
# imageSearchStartPoint = [0, 0]
def IsImageSearchV2(imageName, searchStart_x=0, searchStart_y=0):
    findImageResult = Util.FindImage(
        imageName,
        searchStart_x,
        searchStart_y,
        region=(
            searchStart_x,
            searchStart_y,
            Util.ScreenSize().width,
            Util.ScreenSize().height,
        ),
    )
    Util.Debug(f"{imageName} result: {findImageResult.resultType}")
    return findImageResult.resultType == Enum_FIND_IMAGE_RESULT_TYPE.Success


# # 클립보드에서_중복되는_것중에_최대_값_추출
# def GetClipboardOverlapMaxValue(clipboardData):

# 	maxValue = 0

# 	resultArray = [] # 값을 저장할 배열

# 	pos = 1
# 	while pos = RegExMatch(clipboardData, "/w_(\d+)(?=\/)/", match, pos)
# 		resultArray.append(match1) # 매치된 값을 배열에 추가
# 		pos += StrLen(match) # 다음 검색 위치 업데이트

# 	uniqueElements =  # 중복 인지 여부 요소를 담을 객체 생성

# 	# 중복되지 않는 요소를 객체에 기록
# 	for _, value in resultArray
# 		uniqueElements[value] = uniqueElements.HasKey(value) ? 2 : 1 # 중복된 요소는 2, 그 외는 1로 설정

# 	# 중복되지 않는 요소 제외하고 배열 재구성
# 	overlappedValueArray = []
# 	for _, value in resultArray
# 		if uniqueElements[value] == 2: # 중복된 요소만 선택하여 새 배열에 추:
# 			overlappedValueArray.append(value)

# 	# 배열에서 최대값 찾기
# 	for _, value in resultArray
# 		if value > maxValue:
# 			maxValue = value

# 	return maxValue


# # 16진수 -> 10진수로 변경
# def HEX2DEC(rgb):

# 	Util.Debug(f"HEX2DEC rgb : " + rgb)
# 	formattedNumber = Format(":d", rgb) # 숫자를 10진수로 형식화 로 나중에 테스트 하기기
# 	Util.Debug(f"HEX2DEC Format formattedNumber : " + formattedNumber)
# 	formattedNumber = "0x" + formattedNumber
# 	formattedNumber += 0
# 	Util.Debug(f"HEX2DEC Format formattedNumber end : " + formattedNumber)

# 	SetFormat, IntegerFast, d # 10진수로 정수 형식으로 설정
# 	decimalCode = "0x" + rgb
# 	decimalCode += 0
# 	Util.Debug(f"HEX2DEC SetFormat decimalCode : " + decimalCode)
# 	return decimalCode


# 배열에 중복값 제거
def RemoveDuplicatesFromArray(arr):
    uniqueArray = []
    for _, value in arr:
        if value not in uniqueArray:
            uniqueArray.append(value)
    return uniqueArray


# 값이 배열에 있는지 확인하는 함수
def IsValueInArray(value, arr):
    for _, val in arr:
        if val == value:
            return True
    return False


def Debug(value, isShowPopup=True):
    value = str(value)
    nowTime = Util.GetFormattedCurrentDateTime()
    # 문자열을 UTF-8로 인코딩하여 바이트열로 변환하고 연결
    result = nowTime.encode("utf-8") + value.encode("utf-8")
    print(result.decode("utf-8"))  # 바이트열을 다시 문자열로 디코딩하여 출력
    # Tcl_AsyncDelete: async handler deleted by the wrong thread 가 발생해서 주석
    # if isShowPopup == True:
    #     ShowPopup(value, 3)
    return


# 현재 실행 중인 팝업 스레드를 저장하는 변수
current_progress_thread = None


def print_progress_thread(output_value, duration):
    global current_progress_thread
    if current_progress_thread:
        # 이전 팝업 스레드가 있으면 종료
        current_progress_thread.join()
    progress_thread = threading.Thread(
        target=print_progress, args=(output_value, duration)
    )
    progress_thread.daemon = True  # 메인 프로세스 종료 시 함께 종료되도록 설정
    progress_thread.start()
    current_progress_thread = progress_thread
    return progress_thread


def print_progress(output_value, duration):
    # Tkinter 윈도우 생성
    root = tk.Tk()
    root.attributes("-alpha", 0.0)  # 윈도우를 투명하게 만듦
    root.geometry("200x100")  # 윈도우 크기 설정

    # 화면 크기 구하기
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 말풍선 생성
    balloon = tk.Toplevel(root)
    balloon.overrideredirect(True)  # 말풍선 윈도우에 테두리와 타이틀 바 숨김
    balloon.attributes("-alpha", 0.9)  # 말풍선을 투명하게 만들지 않음
    balloon.attributes("-topmost", True)  # 말풍선을 최상위로 설정

    # 화면 왼쪽 하단에 말풍선 위치 설정
    x_position = 0
    y_position = screen_height - 70
    balloon.geometry("+{}+{}".format(x_position, y_position))

    # 말풍선에 텍스트 레이블 추가
    label = tk.Label(balloon, text=output_value, padx=5, pady=5)
    label.pack()

    # 일정 시간이 지난 후에 말풍선을 닫음
    root.after(duration * 1000, root.destroy)

    # Tkinter 윈도우 실행
    root.mainloop()


# 팝업을 보여주는 함수
def ShowPopup(text, duration):
    global current_progress_thread
    progress_thread = print_progress_thread(text, duration)
    current_progress_thread = progress_thread


# 크롬 현재 탭 닫기
def NowChromeTabExit():
    Util.KeyboardKeyHotkey("ctrl", "w")  # Ctrl + W로 현재 탭 닫기


# Clipboard에 복사 됨
def ChromeTranslateGoogle(textToTranslate):
    # 번역할 텍스트 복사
    pyperclip.copy(textToTranslate)
    # 웹 브라우저 열기 및 Google 번역 페이지로 이동
    webbrowser.open("https://translate.google.com/")
    # Google Chrome이 실행되고 그 창이 활성화될 때까지 대기
    Util.SleepTime(1)
    WhileFoundImage("구글 번역\구글 번역 화면 로딩 끝")
    Util.KeyboardKeyHotkey("ctrl", "v")  # 번역한 텍스트 붙여넣기
    ClickAtWhileFoundImage("구글 번역\번역 된 것 복사 버튼", 10, 10)
    NowChromeTabExit()
    Util.SleepTime(0.5)
    return pyperclip.paste()


# 크롬에서 특정 탭으로 이동하는 기능(제목 포함 하면 멈춤)
def WhileFindChromeTab(tabTitle):
    while True:
        Util.KeyboardKeyHotkey("ctrl", "tab")  # Ctrl + Tab을 눌러 다음 탭으로 이동
        Util.SleepTime(1)  # 1초 대기하여 페이지 로딩을 기다립니다.
        # 현재 탭의 제목 가져오기
        findImageResult = Util.FindImage("title.png")
        if findImageResult.resultTyp == Enum_FIND_IMAGE_RESULT_TYPE.Success:
            Util.MouseMove(
                findImageResult.x, findImageResult.y
            )  # 제목이 화면에 보이면 마우스를 해당 위치로 이동
            Util.SleepTime(0.5)  # 마우스 이동 후에는 잠시 대기하여 안정화
            break  # 원하는 탭이 활성화되었으므로 루프를 종료합니다.


# 통관 고유부호 검증
# name = "나철환"
# ecm = "P180002538686"
# phoneNumber = "01091700607"
def IsCheckPcc(name, ecm, phoneNumber):
    WhileFindChromeTab("통관고유부호 검증")  # https://gsiexpress.com/pcc_chk.php
    ClickAtWhileFoundImage("통관고유부호\통관 고유부호 검증", 0, 150)
    pyperclip.copy(name + "/" + ecm + "/" + phoneNumber)
    Util.KeyboardKeyHotkey("ctrl", "v")  # 번역한 텍스트 붙여넣기
    ClickAtWhileFoundImage("통관고유부호\통관부호 검증 확인")
    Util.SleepTime(1)
    return IsImageSearchV2("통관고유부호\검증 결과 정상")


# 크롬 창에서 주소창으로 이동
def GoToTheAddressWindow():
    Util.KeyboardKeyHotkey("alt", "d")


# def GetYesshipAddress():

# 	FormatTime, currentDate,, yyyyMMdd # 현재 날짜를 년월일 형식(yyyyMMdd)으로 얻기

# 	# 현재 날짜에서 2달을 빼고 2달 전의 날짜 구하기
# 	currentYear = SubStr(currentDate, 1, 4)
# 	currentMonth = SubStr(currentDate, 5, 2)
# 	currentDay = SubStr(currentDate, 7, 2)

# 	# 2달 전의 월 계산
# 	if currentMonth <= 2:
# 		currentYear = currentYear - 1
# 		previousMonth = currentMonth + 10 # 12월인 경우 10을 더함 else:
# 		previousMonth = currentMonth - 2

# 	previousDate = Format(":04d:02d:02d", currentYear, previousMonth, currentDay)

# 	# 예스쉽 입고완료에서 검색옵션을 주문번호 조회
# 	GoToTheAddressWindow()
# 	return "https://www.yesship.kr:8440/page/deliveryorder_list.asp?txtstartdate=" + previousDate + "&txtenddate=" + currentDate + "&gubun2="


# # 예스쉽 입고완료에서 검색옵션을 주문번호 조회
# def IsFindYesshipByOrderNumber(gubun, orderNumber):

# 	GoToTheAddressWindow()
# 	pyperclip.copy(GetYesshipAddress() + gubun + "&gubun=pnumber&keyword=" + orderNumber + "&page=1")
# 	pyautogui.hotkey('ctrl', 'v') # 번역한 텍스트 붙여넣기
# 	pyautogui.press('enter')
# 	Util.SleepTime(3)
# 	return !IsImageSearch("예스쉽\데이터가 없습니다")


# # 통관(세관) 시작 여부 체크
# def IsStartCustomsClearance(hblNo):

# 	oHTTP = ComObjCreate("WinHttp.WinHttpRequest.5.1")
# 	oHTTP.Open("GET", "https://unipass.customs.go.kr:38010/ext/rest/cargCsclPrgsInfoQry/retrieveCargCsclPrgsInfo?crkyCn=h280n293k172p003i070b030e0&hblNo=" + hblNo + "&blYy=2023", False)
# 	oHTTP.Send()
# 	responseBody = oHTTP.ResponseText
# 	return InStr(responseBody, "<hblNo>" + hblNo + "</hblNo>")


# # CJ 운송장 받았는지 여부 체크 만드는 중
# def IsStartCJ(hblNo):
#     GoToTheAddressWindow()
#     clipboard_content = (
#         "https://trace.cjlogistics.com/next/tracking.html?wblNo=" + hblNo
#     )
#     pyautogui.hotkey("ctrl", "v")  # 번역한 텍스트 붙여넣기
#     pyautogui.press("enter")
#     Util.SleepTime(3)
#     return not IsImageSearch("유효하지 않는 운송장번호 입니다")

#     # oHTTP = ComObjCreate("WinHttp.WinHttpRequest.5.1")
#     # oHTTP.Open("GET", "https://trace.cjlogistics.com/next/tracking.html?wblNo=433032447342", False)
#     # oHTTP.Send()
#     # responseBody = oHTTP.ResponseText


# 배열을 , 구분해서 string로 반환
def JoinArrayToString(array) -> str:
    outputString = ""
    for index, item in enumerate(array, start=1):
        outputString += ("," if index > 1 else "") + item
    return outputString


# 입력 예 -"aaa(205,260,270),bbb(260,261)"
def StringToDoubleArray(input):
    input_str = input_str.replace(" ", "_")
    result = []

    pattern = r"(\w+)\(([^)]+)\)"
    pos = 0
    while True:
        match = re.search(pattern, input_str[pos:])
        if match:
            key = match.group(1)
            values = match.group(2).split(",")

            keyValue = [key, values]
            result.append(keyValue)

            pos += match.end()
        else:
            break

    return result


def DoubleArrayToString(array) -> str:
    formattedString = ""
    for index, item in enumerate(array, start=1):
        key = item[0]
        values = item[1]

        valuesString = ",".join(values)

        formattedString += f"{key}({valuesString})"
        if index < len(array):
            formattedString += ","

    return formattedString


# 현재날짜및시간반환
def GetFormattedCurrentDateTime() -> str:
    return datetime.now().strftime("%Y/%m/%d/%H:%M:%S")


# 옵션 엑셀 세팅
def SetExcelOptionByString(stringDoubleArray):
    result = StringToDoubleArray(stringDoubleArray)
    Util.SetExcelOption(result, False)


# 옵션 엑셀 세팅
def SetExcelOption(doubleArray, is_customsDuty):

    # # 과세 부가 되는 기준을 정해서 옵션에 추가 할지 여부 정하기
    # # 31만원 이상이면 무조건 과세 부가
    # # 26 ~ 30만원 사이는 과세가 부가 될 수도 있습니다. 정확한 것을 판매자에게 문의 주세요 라고 옵션에 적기
    # # 25만원 이하는 무조건 과세 아님

    # 사이즈 것을 옵션에 넣을까 고민 중
    # 사이즈 신중하게 결정해주세요(사이즈로 교환 및 환불 불가 합니다.)

    xlFile = os.path.join(
        EnvData.g_DefaultPath(),
        "엑셀",
        f"OptionCombinationTemplate{'_' if is_customsDuty != False else '2_'}"
        + ".xlsx",
    )
    wb = openpyxl.load_workbook(filename=xlFile)  # 엑셀 파일 열기
    ws = wb.active
    rowCount = ws.max_row
    allCount = 0
    for index, item in enumerate(doubleArray):
        colorName = item[Array_ColroName]
        sizes = item[Array_SizeList]
        for index2, size in enumerate(sizes):
            allCount += 1
            ws.cell(row=(allCount + 1), column=1).value = colorName
            ws.cell(row=(allCount + 1), column=2).value = size
            if is_customsDuty != False:
                ws.cell(row=(allCount + 1), column=3).value = (
                    "관부가세(23%) 수취인 부담(통관시 납부)"
                )
            ws.cell(
                row=(allCount + 1), column=(4 if is_customsDuty != False else 3)
            ).value = 0
            ws.cell(
                row=(allCount + 1), column=(5 if is_customsDuty != False else 4)
            ).value = 300
            ws.cell(
                row=(allCount + 1), column=(6 if is_customsDuty != False else 5)
            ).value = f"{index}{index2}"
            ws.cell(
                row=(allCount + 1), column=(7 if is_customsDuty != False else 6)
            ).value = "Y"

    # 이전 것 초과 된 행 삭제
    for _ in range(rowCount - (allCount + 1)):
        ws.delete_rows(allCount + 2)

    wb.save(xlFile)


def CopyToClipboardAndGet():
    pyperclip.copy("")
    Util.SleepTime(0.1)
    pyautogui.hotkey("ctrl", "c")
    Util.SleepTime(0.1)
    return pyperclip.paste()


def IsArray(array, checkValue) -> bool:
    for item in array:
        if item == checkValue:
            return True
    return False


def GetRegExMatcheList(value, regexPattern, startPos=1):
    matcheList = []

    pos = startPos
    while True:
        match = re.search(regexPattern, value[pos:])
        if match:
            matcheList.append(match.group())
            pos += match.end()
        else:
            break

    for index, matche in enumerate(matcheList, start=1):
        Util.Debug(f"matcheList[{index}] : {matche}")

    return matcheList


def GetRegExMatcheGroup1List(value, regexPattern, startPos=1):
    matche1List = []

    pos = startPos
    while True:
        match = re.search(regexPattern, value[pos:])
        if match:
            matche1List.append(match.group(1))
            pos += match.end()
        else:
            break

    for index, matche1 in enumerate(matche1List, start=1):
        Util.Debug(f"matche1List[{index}] : {matche1}")

    return matche1List


def DownloadImageUrl(url, saveName) -> bool:
    # 저장할 파일의 경로 및 사용자가 원하는 파일 이름
    savePath = EnvData.g_DefaultPath() + "\\DownloadImage\\" + str(saveName) + ".png"

    # 이미지 다운로드
    try:
        urllib.request.urlretrieve(url, savePath)
        Util.Debug(f"이미지 다운로드 url : {url}")
    except Exception as e:
        Util.Debug(f"이미지 다운로드 실패 Error: {e}")
        return False

    return True


def FolderToDelete(folderPath):
    # 폴더 안의 모든 파일 삭제
    for root, dirs, files in os.walk(folderPath):
        for file in files:
            filePath = os.path.join(root, file)
            os.remove(filePath)
    Util.Debug(f"폴더 안의 모든 파일 삭제  folderPath : " + folderPath)


def ExcelSystemKill():
    # 열려있는 엑셀 닫기(이유 : 열려 있는 엑셀 사용시 읽기 전용으로 해서 저장시 에러 나기 때문에
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"])
    # 시간 3초는 닫는데 걸리는 시간(제대로 안 닫히면 엑셀 여는에 문제가 있기 때문에)
    Util.SleepTime(3)


def SleepTime(second):
    time.sleep(second)


def TelegramSend(Message, isDebug=True):
    if isDebug == True:
        Util.Debug(f"TelegramSend : " + Message)
    ChatID = EnvData.TelegramSend_ChatID()
    Token = EnvData.TelegramSend_Token()
    # 매개변수 설정
    Param = {"chat_id": ChatID, "text": Message}

    # URL 설정
    URL = "https://api.telegram.org/bot" + Token + "/sendmessage"

    # POST 요청 보내기
    response = requests.post(URL, data=Param)


def GetKorMony(mony, exchangeRate) -> int:
    # 물건 원가
    korCostPrice = mony * exchangeRate
    if korCostPrice == 0 and mony != 0:
        Util.TelegramSend(
            f"******** Error  ---- GetKorMony() ==== 0   mony : {mony}    exchangeRate : {exchangeRate} ----  Error  "
        )
        return 0
    else:
        # 수익율
        marginPrice = korCostPrice * GlobalData.g_MarginRate()
        # 배송비
        courierPrice = GlobalData.g_CourierPrice()
        # -3은 천원 단위로 반올림 하기 위함
        outValue = round(korCostPrice + marginPrice + courierPrice, -3)
        if outValue == 0 and mony != 0:
            Util.TelegramSend(
                f"******** Error  ---- GetKorMony() ==== {outValue}   mony : {mony}    exchangeRate : {exchangeRate} ----  Error  "
            )
        return int(outValue)


def GetUggKorSize(usSize):
    match usSize:
        case "4":
            return 210
        case "5":
            return 220
        case "5.5":
            return 225
        case "6":
            return 230
        case "6.5":
            return 235
        case "7":
            return 240
        case "7.5":
            return 245
        case "8":
            return 250
        case "8.5":
            return 255
        case "9":
            return 260
        case "9.5":
            return 265
        case "10":
            return 270
        case "10.5":
            return 275
        case "11":
            return 280
        case "11.5":
            return 285
        case "12":
            return 290
        case "13":
            return 300
        case "14":
            return 310
    return 0


def GetMytheresaKorSize(size):
    match size:
        case "EU 33":
            return "210"
        case "EU 33.5":
            return "215-"
        case "EU 34":
            return "215+"
        case "EU 34.5":
            return "220"
        case "EU 35":
            return "225-"
        case "EU 35.5":
            return "225+"
        case "EU 36":
            return "230"
        case "EU 36.5":
            return "235-"
        case "EU 37":
            return "235+"
        case "EU 37.5":
            return "240"
        case "EU 38":
            return "245-"
        case "EU 38.5":
            return "245+"
        case "EU 39":
            return "250"
        case "EU 39.5":
            return "255-"
        case "EU 40":
            return "255+"
        case "EU 40.5":
            return "260"
        case "EU 41":
            return "265-"
        case "EU 41.5":
            return "265+"
        case "EU 42":
            return "270"
        case "EU 42.5":
            return "275-"
        case "EU 43":
            return "275+"
        case "EU 43.5":
            return "280"
        case "EU 44":
            return "285-"
        case "EU 44.5":
            return "285+"
    return ""


class Enum_FIND_IMAGE_RESULT_TYPE(Enum):
    Success = "Success"
    Fail_NoRead = "Fail_NoRead"
    Fail_NoFind = "Fail_NoFind"
    Fail_NoFile = "Fail_NoFile"


class Module_FindImageResult:
    def __init__(self):
        self.resultType = ""
        self.x = ""
        self.y = ""
        self.title = ""


def FindImage(
    image_path, searchStart_x=0, searchStart_y=0, region={}, threshold=0.7
) -> Module_FindImageResult:
    # 이미지 파일 존재 확인
    if not os.path.exists(image_path):
        returnValue = Module_FindImageResult()
        returnValue.resultType = Enum_FIND_IMAGE_RESULT_TYPE.Fail_NoFile
        returnValue.x = 0
        returnValue.y = 0
        return returnValue

    # 이미지 경로를 문자열로 변환하여 읽기(한글주소가 되게 하기 위함)
    img_array = np.fromfile(image_path, np.uint8)
    template = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
    if template is None:
        returnValue = Module_FindImageResult()
        returnValue.resultType = Enum_FIND_IMAGE_RESULT_TYPE.Fail_NoRead
        returnValue.x = 0
        returnValue.y = 0
        return returnValue

    # 화면 캡처
    screenshot = pyautogui.screenshot(region=region)
    screenshot = cv2.cvtColor(np.array(screenshot), cv2.IMREAD_COLOR)

    # 이미지 매칭 수행
    res = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)

    loc = np.where(res >= threshold)

    # for pt in zip(*loc[::-1]):
    #     # 매칭된 좌표에서 잘린 부분 추출
    #     roi = screenshot[pt[1]:pt[1] + template.shape[0], pt[0]:pt[0] + template.shape[1]]
    #     # 템플릿 이미지와의 매칭 확인
    #     res_match = cv2.matchTemplate(roi, template, cv2.TM_CCOEFF_NORMED)
    #     max_match = np.max(res_match)
    #     if max_match >= threshold:
    #         # 추가 검증 후 성공으로 판단
    #         return Module_FindImageResult(
    #             Enum_FIND_IMAGE_RESULT_TYPE.Success,
    #             searchStart_x + pt[0],
    #             searchStart_y + pt[1],
    #         )

    # return Module_FindImageResult(Enum_FIND_IMAGE_RESULT_TYPE.Fail_NoFind, 0, 0)

    if loc[0].size > 0 and loc[1][0] != 0 and loc[1][0] != 0:
        # 일치하는 이미지가 발견되면 좌표를 반환
        returnValue = Module_FindImageResult()
        returnValue.resultType = Enum_FIND_IMAGE_RESULT_TYPE.Success
        returnValue.x = searchStart_x + loc[1][0]
        returnValue.y = searchStart_y + loc[0][0]
        return returnValue
    else:
        returnValue = Module_FindImageResult()
        returnValue.resultType = Enum_FIND_IMAGE_RESULT_TYPE.Fail_NoFind
        returnValue.x = 0
        returnValue.y = 0
        return returnValue


def save_and_close_open_excel_files():
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if "EXCEL.EXE" in proc.info["name"]:  # 엑셀 프로세스인지 확인
                for conn in proc.connections():
                    if (
                        conn.laddr and conn.laddr.port
                    ):  # 엑셀 파일에 연결된 프로세스 확인
                        file_path = conn.laddr.ip  # 파일 경로 또는 IP 주소
                        # 파일 저장
                        try:
                            wb = openpyxl.load_workbook(file_path)
                            wb.save(f"{os.path.basename(file_path)}_backup.xlsx")
                            print(f"저장 및 닫기 완료: {file_path}")
                            wb.close()  # 엑셀 파일 닫기
                        except Exception as e:
                            print(f"저장 및 닫기 중 에러 발생: {file_path}", e)
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass


def TranslateToKorean(enText: str) -> str:
    returnValue = ""
    if enText:
        Util.Debug(f"번역 시작 : {enText}")
        trans = Translator()
        result = trans.translate(enText, dest="ko", src="en")
        Util.Debug(f"원  문({result.src}): {result.origin}")
        Util.Debug(f"번역문({result.dest}) : {result.text}")
        returnValue = result.text
    return returnValue

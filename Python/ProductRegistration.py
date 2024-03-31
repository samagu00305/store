import Util
import pyautogui
import pyperclip
import EnvData


# 상품 등록시 기본 세팅 들
def ProductRegistrationDefaultSettings():
    Util.WheelAndClickAtWhileFoundImage(
        r"스마트 스토어\상품 수정\상품 주요정보", 10, 10
    )
    Util.SleepTime(0.5)
    Util.WheelAndClickAtWhileFoundImage(
        r"스마트 스토어\상품 수정\구매대행 라디오 버튼", 13, 13
    )
    Util.SleepTime(0.5)
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\국산", 10, 10)
    Util.SleepTime(0.5)
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\기타", 10, 10)
    Util.SleepTime(0.5)
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\에누리", 10, 10)
    Util.WheelAndClickAtWhileFoundImage(
        r"스마트 스토어\상품 수정\다나와", 10, 10, 100, 10
    )
    Util.SleepTime(0.5)
    Util.WheelAndMoveAtWhileFoundImage(r"스마트 스토어\제일 위 상품등록", 0, 0, 10000)


# 이미지 등록(대표, 추가)
def IamgeRegistration_v2(imgCount: int):
    # 대표 이미지 등록
    Util.WheelAndClickAtWhileFoundImage(r"스마트 스토어\상품 수정\대표이미지", 250, 20)
    Util.SleepTime(0.5)
    Util.ClickAtWhileFoundImage(r"스마트 스토어\내 사진", 0, 0)
    Util.SleepTime(1.5)
    for _ in range(5):
        pyautogui.hotkey("tab")
        Util.SleepTime(0.2)
    pyautogui.hotkey("enter")
    Util.SleepTime(0.3)
    pyperclip.copy(EnvData.g_DefaultPath() + r"\DownloadImage")
    Util.SleepTime(0.5)
    Util.KeyboardKeyHotkey("ctrl", "v")
    Util.SleepTime(0.5)
    pyautogui.hotkey("enter")
    Util.SleepTime(0.3)
    for _ in range(4):
        pyautogui.hotkey("tab")
        Util.SleepTime(0.2)
    Util.SleepTime(0.5)
    pyautogui.hotkey("right")
    Util.SleepTime(0.5)
    pyautogui.hotkey("left")
    Util.SleepTime(0.5)
    pyautogui.hotkey("enter")
    Util.SleepTime(2)

    pyautogui.hotkey("esc")
    Util.SleepTime(1)

    if Util.MoveAtWhileFoundImage(r"스마트 스토어\내 사진", 0, 0, 2):
        # 이미지 넣을 수 없는 것임(너무 큼)
        return False

    if imgCount > 1:
        # 추가 이미지 등록
        Util.ClickAtWhileFoundImage(r"스마트 스토어\상품 수정\추가이미지", 0, 0)
        Util.SleepTime(1)
        pyautogui.scroll(-50)  # wheelMove 틱 스크롤 다운
        Util.SleepTime(1)
        Util.WheelAndClickAtWhileFoundImage(
            r"스마트 스토어\상품 수정\추가이미지", 250, 20, -500
        )
        Util.SleepTime(0.5)
        Util.ClickAtWhileFoundImage(r"스마트 스토어\내 사진", 0, 0)
        Util.SleepTime(1.5)
        for _ in range(5):
            pyautogui.hotkey("tab")
            Util.SleepTime(0.2)
        pyautogui.hotkey("enter")
        Util.SleepTime(0.3)
        pyperclip.copy(EnvData.g_DefaultPath() + r"\DownloadImage")
        Util.SleepTime(0.5)
        Util.KeyboardKeyHotkey("ctrl", "v")
        Util.SleepTime(0.5)
        pyautogui.hotkey("enter")
        Util.SleepTime(0.3)
        for _ in range(4):
            pyautogui.hotkey("tab")
            Util.SleepTime(0.2)
        Util.SleepTime(0.5)
        pyautogui.hotkey("right")
        Util.SleepTime(0.5)
        pyautogui.hotkey("left")
        Util.SleepTime(0.5)
        pyautogui.hotkey("delete")
        Util.SleepTime(0.5)
        pyautogui.hotkey("ctrl", "a")
        Util.SleepTime(0.5)
        pyautogui.hotkey("enter")

    return True

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


class COLUMN_Add11(Enum):
    A = "A"  # 네이버 단지 번호
    B = "B"  #
    C = "C"  #
    D = "D"  #


def GetElementsData_v4(url: str, sleepTime) -> str:
    html = ""

    # 웹 드라이버 초기화
    driver = webdriver.Chrome()

    # 웹 페이지 열기
    driver.get(url)
    Util.SleepTime(sleepTime)

    try:
        ClickButton(driver, "단지정보")
        ClickButton(driver, "시세/실거래가")
        # ClickButton(driver, "동호수/공시가격")

        # 추가적인 대기 시간을 둘 수 있음 (예: 동적 콘텐츠 로딩을 위해)
        Util.SleepTime(2)
        html = driver.page_source

    except Exception as e:
        print("에러가 발생했습니다.:", e)

    finally:
        # 웹 브라우저 종료
        driver.quit()

    return html


def ClickButton(driver, buttonName):
    buttons = driver.find_elements(
        By.XPATH,
        f"//div[@class='complex_detail_link']//button[contains(text(), '{buttonName}')]",
    )

    # '단지정보' 버튼이 존재하는지 확인
    if len(buttons) > 0:
        button = buttons[0]  # 첫 번째 매칭되는 버튼을 선택
        driver.execute_script("arguments[0].click();", button)

        Util.SleepTime(1)
    else:
        print(f"{buttonName} 버튼을 찾을 수 없습니다.")


def GetElementsData_v5(url: str, sleepTime) -> str:
    html = ""

    # 웹 드라이버 초기화
    driver = webdriver.Chrome()

    # 웹 페이지 열기
    driver.get(url)
    Util.SleepTime(sleepTime)

    try:
        # 스크롤하고자 하는 요소 찾기
        scrollable_element = driver.find_element(By.CLASS_NAME, "item_list--article")

        # 무한 스크롤 처리를 위한 반복문
        while True:
            # 스크롤 이전의 높이
            prev_height = driver.execute_script(
                "return arguments[0].scrollHeight", scrollable_element
            )

            # 스크롤을 끝까지 내림
            driver.execute_script(
                "arguments[0].scrollTop = arguments[0].scrollHeight", scrollable_element
            )

            # 페이지 로딩 대기
            Util.SleepTime(2)  # 실제 사이트에 따라 대기 시간 조정 필요

            # 스크롤 이후의 높이
            new_height = driver.execute_script(
                "return arguments[0].scrollHeight", scrollable_element
            )

            # 스크롤 이전 높이와 이후 높이가 같다면, 더 이상 로딩되는 내용이 없다는 의미이므로 반복 종료
            if prev_height == new_height:
                break

        # 추가적인 대기 시간을 둘 수 있음 (예: 동적 콘텐츠 로딩을 위해)
        Util.SleepTime(2)
        html = driver.page_source

    except Exception as e:
        print("버튼을 찾을 수 없습니다:", e)

    finally:
        # 웹 브라우저 종료
        driver.quit()

    return html

import tkinter as tk
from tkinter import messagebox
import Util
import System
import traceback

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


def category():
    xlFileAddBefore = (
        EnvData.g_DefaultPath() + r"\엑셀\스마트스토어상품_20240421_140228.csv"
    )
    dfAddBefore = pd.read_csv(xlFileAddBefore)

    xlFile = EnvData.g_DefaultPath() + r"\엑셀\마구싸5_구매루트.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    df = df.astype(str)
    lastRow = df.shape[0]

    row = 1

    while True:
        row += 1

        if row >= lastRow:
            break

        value = df.at[row, System.COLUMN.A.name]
        if value.isdigit():  # 문자열이 숫자로 표현되어 있는지 확인
            numeric_value = int(value)
        indexList = dfAddBefore[
            dfAddBefore["상품번호(스마트스토어)"] == numeric_value
        ].index.tolist()
        if len(indexList) != 0:
            data = dfAddBefore.index[indexList[0]]
            category = dfAddBefore.at[data, "대분류"]
            if not pd.isna(dfAddBefore.at[data, "중분류"]):
                category += " " + dfAddBefore.at[data, "중분류"]
            if not pd.isna(dfAddBefore.at[data, "소분류"]):
                category += " " + dfAddBefore.at[data, "소분류"]
            if not pd.isna(dfAddBefore.at[data, "세분류"]):
                category += " " + dfAddBefore.at[data, "세분류"]

            df.at[row, System.COLUMN.V.name] = category
            System.SaveWorksheet(df)

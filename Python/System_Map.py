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


def SetCsvNewProductURLs_Common():
    xlFile = EnvData.g_DefaultPath() + r"\엑셀\추가 할 네이버 단지 번호.CSV"
    df = pd.read_csv(xlFile, encoding="cp949")
    lastRow = df.shape[0]

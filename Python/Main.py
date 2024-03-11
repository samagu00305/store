import cv2
import numpy as np
import openpyxl
import pyautogui
import tkinter as tk
from tkinter import messagebox
import os
import Util
import EnvData
import requests
from typing import NamedTuple
from enum import Enum
import System


def show_start_popup():
    System.UpdateStoreWithColorInformation(-1)

    Util.SleepTime(5)
    Util.TelegramSend("Test")

    # # 현재 스크립트의 디렉토리 경로
    # script_dir = os.path.dirname(os.path.abspath(__file__))

    # # 이미지 파일의 상대 경로
    # image_path = os.path.join(script_dir, "html.png")

    # if os.path.exists(image_path) == False:
    #     print("파일이 존재하지 않습니다.")
    #     return

    # # 이미지를 화면에서 찾기
    # result = find_image_on_screen(image_path)
    # if isinstance(result, tuple):
    #     print("일치하는 이미지가 화면에서 발견되었습니다. 위치:", result)
    # else:
    #     print(result)


class Point:
    def __init__(self, x, y, z):
        self.x = x
        self.y = y
        self.y = z


class Size(NamedTuple):
    x: int
    y: int
    z: int


def show_exit_popup():
    if messagebox.askyesno("종료하기", "정말 종료하시겠습니까?"):
        root.destroy()


# Tkinter 애플리케이션 생성
root = tk.Tk()

# 창 제목 설정
root.title("자동")

# 창 크기 지정 (너비x높이)
root.geometry("200x100")

root.update_idletasks()
width = root.winfo_width()
height = root.winfo_height()
x = (root.winfo_screenwidth() // 2) - (width // 2)
y = (root.winfo_screenheight() // 2) - (height // 2)
root.geometry(f"{width}x{height}+{x}+{y}")


# 시작하기 버튼 생성 및 가운데 정렬
start_button = tk.Button(root, text="시작하기", command=show_start_popup)
start_button.place(relx=0.3, rely=0.5, anchor="center")

# 종료하기 버튼 생성 및 가운데 정렬
exit_button = tk.Button(root, text="종료하기", command=show_exit_popup)
exit_button.place(relx=0.7, rely=0.5, anchor="center")

# Tkinter 애플리케이션 실행
root.mainloop()

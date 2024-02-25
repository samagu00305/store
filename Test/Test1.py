
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pyperclip  # pyperclip 패키지를 사용하여 클립보드에 접근

import tkinter as tk
from tkinter import messagebox

# Tkinter 창 생성
root = tk.Tk()
root.withdraw()  # 창 숨기기


# messagebox.showinfo('제목', '000')

# 사용할 범위(scope) 설정
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

global_var = 10  # 전역 변수 정의
    
def get_credentials():

    global global_var  # 함수 내에서 전역 변수 사용을 선언
    creds = None

    # 저장된 토큰 파일 불러오기
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    # 토큰이 유효하지 않으면 새로운 토큰 가져오기
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_800061949795-ef3q3g3v2abna89rburue93o0o0ga84k.apps.googleusercontent.com.json', SCOPES)  # credentials.json은 Google API 콘솔에서 다운로드 받은 클라이언트 정보 파일
            creds = flow.run_local_server(port=0)

        # 토큰을 클립보드에 복사
        pyperclip.copy(creds.to_json())
        # messagebox.showinfo('제목', '토큰이 클립보드에 복사되었습니다.')

    return creds

# OAuth 2.0 인증 획득
credentials = get_credentials()


# 메시지 박스 팝업
messagebox.showinfo('제목', '내용을 입력하세요.')





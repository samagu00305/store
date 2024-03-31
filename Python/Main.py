import tkinter as tk
from tkinter import messagebox
import Util
import System
import traceback


def show_start_popup():
    try:
        System.CloseExcelProcesses()

        # System.SetCsvBananarePublicNewProductURLs()

        addCount = System.AddDataFromExcel_BananarePublic()

        # 신규 등록 할 UGG 목록을 CSV에 정리
        # System.SetCsvUGGNewProductURLs()

        # System.CloseExcelProcesses()

        # addCount = System.AddDataFromExcel_Ugg()

        # System.UpdateStoreWithColorInformation(-1)

        System.CloseExcelProcesses()

        Util.SleepTime(5)
        Util.TelegramSend("Test")
    except:
        stack_trace_str = traceback.format_exc()
        Util.TelegramSend(str(stack_trace_str), False)
        # save_and_close_open_excel_files()
        raise


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

close_excel_button = tk.Button(
    root, text="엑셀 다 끄기", command=System.CloseExcelProcesses
)
close_excel_button.place(relx=0.5, rely=0.2, anchor="center")


# 종료하기 버튼 생성 및 가운데 정렬
exit_button = tk.Button(root, text="종료하기", command=show_exit_popup)
exit_button.place(relx=0.7, rely=0.5, anchor="center")

# Tkinter 애플리케이션 실행
root.mainloop()

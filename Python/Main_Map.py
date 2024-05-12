import tkinter as tk
from tkinter import messagebox
import Util
import System
import System_Map
import traceback
import pandas as pd


def show_start_popup():
    try:
        Util.TelegramSend(f"시작")

        pd.set_option("future.no_silent_downcasting", True)

        System.CloseExcelProcesses()

        xlFile = System.EnvData.g_DefaultPath() + r"\엑셀\존재하는 네이버 단지 번호.CSV"
        df = pd.read_csv(xlFile, encoding="cp949")
        lastRow = df.shape[0]

        if lastRow != 0:
            start = int(df.at[df.index[lastRow - 1], System_Map.COLUMN_Add11.A.name])
        allCount = lastRow
        for i in range(start, 300000):
            index = i + 1
            if index % 2 == 0:
                Util.TelegramSend(f"index : {index}")
            url = f"https://new.land.naver.com/complexes/{index}"
            htmlElementsData: str = System.GetElementsData_v3(url, 1)
            match = System.re.search(r"단지정보", htmlElementsData)
            if match:
                allCount += 1
                df.loc[allCount, System_Map.COLUMN_Add11.A.name] = index
                Util.Debug(f"정보가 존재 합니다.(index : {index})")
            else:
                Util.Debug(f"정보가 안 존재 합니다.(index : {index})")
            if index % 50 == 0:
                Util.CsvSave(df, xlFile)
                Util.TelegramSend(f"Save index : {index}")

        Util.CsvSave(df, xlFile)

        System.CloseExcelProcesses()

        Util.TelegramSend(f"End")
        Util.DiscordSend("End")
        Util.SleepTime(5)

    except:
        stack_trace_str = traceback.format_exc()
        Util.TelegramSend(str(stack_trace_str), False)
        Util.DiscordSend(str(stack_trace_str), False)
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

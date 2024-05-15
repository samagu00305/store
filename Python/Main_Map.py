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

        for i in range(1, 300000):
            index = i + 1
            if index % 2 == 0:
                Util.TelegramSend(f"index : {index}")
            url = f"https://new.land.naver.com/complexes/{index}"
            htmlElementsData: str = System_Map.GetElementsData_v4(url, 1)
            match = System.re.search(r"단지정보", htmlElementsData)
            if match:
                # 도로명 주소
                roadName = ""
                match = System.re.search(
                    r'"도로명"></i>(.*?)<',
                    htmlElementsData,
                )
                if match:
                    roadName = match.group(1)

                # 전세가
                jeonse = ""
                match = System.re.search(
                    r'<dt class="title">전세가</dt><dd class="data">(.*?)<',
                    htmlElementsData,
                )
                if match:
                    jeonse = match.group(1)
                    jeonse = jeonse.replace(",", "").replace("억", "0000")
                    jeonseList = jeonse.split("~")[0].split()
                    jeonseCount = 0
                    for jeonseData in jeonseList:
                        if jeonseData != "-":
                            jeonseCount += int(jeonseData)

                # 최저 매매가
                purchasePrice = ""
                match = System.re.search(
                    r'<dt class="title">매매가</dt><dd class="data">(.*?)(<|~)',
                    htmlElementsData,
                )
                if match:
                    purchasePrice = match.group(1)
                    purchasePrice = purchasePrice.replace(",", "").replace("억", "0000")
                    purchasePriceList = purchasePrice.split("~")[0].split()
                    purchasePriceCount = 0
                    for purchasePriceData in purchasePriceList:
                        if purchasePriceData != "-":
                            purchasePriceCount += int(purchasePriceData)

                # 매매가 대비 전세가 비율 추출
                sale_to_jeonse_ratio = ""
                match = System.re.search(
                    r'>매매가 대비 전세가</th>.*?"type_result">(.*?)%</td></tr><tr class="">',
                    htmlElementsData,
                )
                if match:
                    sale_to_jeonse_ratio = match.group(1)
                    sale_to_jeonse_ratio = sale_to_jeonse_ratio.replace("%", "")
                    sale_to_jeonse_ratio_parts = sale_to_jeonse_ratio.split("~")[
                        0
                    ].split()
                    sale_to_jeonse_ratio_sum = 0
                    for sale_to_jeonse_ratio_part in sale_to_jeonse_ratio_parts:
                        if sale_to_jeonse_ratio_part != "-":
                            sale_to_jeonse_ratio_sum += int(sale_to_jeonse_ratio_part)

                if jeonseCount != 0 and purchasePriceCount != 0:
                    if jeonseCount >= purchasePriceCount:
                        aa = (purchasePriceCount / 100) * 70
                        if aa <= 20000:
                            if sale_to_jeonse_ratio_sum >= 85:
                                Util.DiscordSend(
                                    f"전세가와 비슷한 것 {jeonseCount} >= {purchasePriceCount}   url : {url} 도로명 : {roadName}"
                                )
                                Util.TelegramSend(
                                    f"전세가와 비슷한 것 {jeonseCount} >= {purchasePriceCount}   url : {url} 도로명 : {roadName}"
                                )

        # xlFile = System.EnvData.g_DefaultPath() + r"\엑셀\존재하는 네이버 단지 번호.CSV"
        # df = pd.read_csv(xlFile, encoding="cp949")
        # lastRow = df.shape[0]

        # if lastRow > 0:
        #     for i in range(1, lastRow):
        #         index = int(df.at[df.index[i - 1], System_Map.COLUMN_Add11.A.name])
        #         url = f"https://new.land.naver.com/complexes/{index}"
        #         htmlElementsData: str = System_Map.GetElementsData_v4(url, 1)
        #         match = System.re.search(r"단지정보", htmlElementsData)
        #         if match:
        #             # 도로명 주소
        #             roadName = ""
        #             match = System.re.search(
        #                 r'"도로명"></i>(.*?)<',
        #                 htmlElementsData,
        #             )
        #             if match:
        #                 roadName = match.group(1)

        #             # 전세가
        #             jeonse = ""
        #             match = System.re.search(
        #                 r'<dt class="title">전세가</dt><dd class="data">(.*?)<',
        #                 htmlElementsData,
        #             )
        #             if match:
        #                 jeonse = match.group(1)
        #                 jeonse = jeonse.replace(",", "").replace("억", "0000")
        #                 jeonseList = jeonse.split("~")[0].split()
        #                 jeonseCount = 0
        #                 for jeonseData in jeonseList:
        #                     if jeonseData != "-":
        #                         jeonseCount += int(jeonseData)

        #             # 최저 매매가
        #             purchasePrice = ""
        #             match = System.re.search(
        #                 r'<dt class="title">매매가</dt><dd class="data">(.*?)(<|~)',
        #                 htmlElementsData,
        #             )
        #             if match:
        #                 purchasePrice = match.group(1)
        #                 purchasePrice = purchasePrice.replace(",", "").replace(
        #                     "억", "0000"
        #                 )
        #                 purchasePriceList = purchasePrice.split("~")[0].split()
        #                 purchasePriceCount = 0
        #                 for purchasePriceData in purchasePriceList:
        #                     if purchasePriceData != "-":
        #                         purchasePriceCount += int(purchasePriceData)

        #             if jeonseCount != 0 and purchasePriceCount != 0:
        #                 if jeonseCount >= purchasePriceCount:
        #                     Util.TelegramSend(
        #                         f"전세가와 비슷한 것{jeonseCount} >= {purchasePriceCount}   url : {url}"
        #                     )
        #                     df.loc[i - 1, System_Map.COLUMN_Add11.C.name] = (
        #                         "전세가와 비슷한 것"
        #                     )

        #             df.loc[i - 1, System_Map.COLUMN_Add11.B.name] = roadName
        #             df.loc[0, System_Map.COLUMN_Add11.C.name] = i
        #             Util.CsvSave(df, xlFile)

        #             match = System.re.search(
        #                 r"해당되는 매물이 없습니다.", htmlElementsData
        #             )
        #             if match:
        #                 Util.Debug(f"해당되는 매물이 없습니다.(index : {index})")
        #             else:
        #                 Util.Debug(f"해당되는 매물이 존재합니다.(index : {index})")

        # if lastRow != 0:
        #     start = int(df.at[df.index[lastRow - 1], System_Map.COLUMN_Add11.A.name])
        # allCount = lastRow
        # for i in range(start, 300000):
        #     index = i + 1
        #     if index % 2 == 0:
        #         Util.TelegramSend(f"index : {index}")
        #     url = f"https://new.land.naver.com/complexes/{index}"
        #     htmlElementsData: str = System.GetElementsData_v3(url, 1)
        #     match = System.re.search(r"단지정보", htmlElementsData)
        #     if match:
        #         allCount += 1
        #         df.loc[allCount, System_Map.COLUMN_Add11.A.name] = index
        #         Util.Debug(f"정보가 존재 합니다.(index : {index})")
        #     else:
        #         Util.Debug(f"정보가 안 존재 합니다.(index : {index})")
        #     if index % 50 == 0:
        #         Util.CsvSave(df, xlFile)
        #         Util.TelegramSend(f"Save index : {index}")

        # Util.CsvSave(df, xlFile)

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

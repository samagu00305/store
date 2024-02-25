#Include EnvData.ahk
#Include GlobalData.ahk
#Include GitUtil.ahk
#Include Util.ahk
#Include ProductRegistration.ahk
#Include Test\Test_Ing.ahk
#Include PhoneControlUtil.ahk
#Include System.ahk

Gui, Add, Text, x30 y5 w110 h20, 프로그램 ; 프로그램 제목 표시
Gui, Add, Text, x60 y25 w50 h20 vA, 준비 !!!
Gui, Add, Text, x300 y50 w400 h50 vB, 진행중 상태
Gui, Add, Button, x20 y80 w110 h20, 시작
Gui, Add, Button, x20 y110 w110 h20, 종료
Gui, Show

isTest := false

return

Button시작:
    {	
        ;// 상품 신규 등록 및 최신화
        while (true)
        {
            ExcelSystemKill()

            ;// 신규 등록 할 UGG 목록을 엑셀에 정리
            SetXlsxUGGNewProductURLs()

            ExcelSystemKill()
            
            ;// 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
            AddDataFromExcel_Ugg()
            
            ExcelSystemKill()

            ;// git에 엑셀 파일 최신으로 올림
            GitSyncPushExcelFile()
                
            ;// 등록 된 상품 최신화
            UpdateStoreWithColorInformation(true)

            ;// 가격 갱신 및 사이즈 갱신(엑셀로)
            ; UpdateStoreWithColorInformationMoney_Mytheresa()

            ExcelSystemKill()

            ;// git에 엑셀 파일 최신으로 올림
            GitSyncPushExcelFile()
        }
        
        GuiControl,,B,시작
        SleepTime(1)
        GuiControl,,B, 1
        SleepTime(1)
        GuiControl,,B, 2
        SleepTime(1)
        GuiControl,,B, 3
        
        ;// 입고대기, 입고완료, 출고완료
        ; yesshipStatus := "없음"
        ; if(IsFindYesshipByOrderNumber("IW", "2311261670"))
        ; {
        ; 	yesshipStatus := "입고대기"
        ; }
        ; else
        ; {
        ; 	if(IsFindYesshipByOrderNumber("IC", "2311261670"))
        ; 	{
        ; 		yesshipStatus := "입고완료"
        ; 	}
        ; 	else
        ; 	{
        ; 		if(IsFindYesshipByOrderNumber("OC", "2311261670"))
        ; 		{
        ; 			yesshipStatus := "출고완료"
        ; 		}
        ; 	}
        ; }
        
        ; Switch yesshipStatus
        ; {
        ; 	Case "입고대기":
        ; 		break
        
        ; 	Case "입고완료":
        ; 		;// 엑셀에 예스쉽 주문번호로 찾아서 예스쉽 상태를 입고완료로 변경
        ; 		break
        
        ; 	Case "출고완료":
        ; 		;// 1. 엑셀에 예스쉽 주문번호로 찾아서 예스쉽 상태를 출고완료로 변경
        ; 		;// 2. 엑셀에 예스쉽 주문번호로 찾아서 예스쉽에 운송장 번호를 엑셀에 입력
        ; 		;// 3. 운송장 번호로
        ; 		;// 통관(세관) 시작 여부 체크
        ; 		IsStartCustomsClearance(11)
        ; 		break
        ; }
        
        ; Send, "IW" ;// 입고대기
        ; Send, "OC" ;// 출고완료
        
        ;// 예스쉽 입고완료로 이동
        ; GoToTheAddressWindow()
        ; Clipboard := "https://www.yesship.kr:8440/page/deliveryorder_list.asp?txtstartdate=20230603&txtenddate=20231203&gubun2=IW"
        ; Send, ^v  ; 번역한 텍스트 붙여넣기
        ; Send, {Enter}
        
        GuiControl,,B,시작
        SleepTime(1)
        GuiControl,,B, 1
        SleepTime(1)
        GuiControl,,B, 2
        SleepTime(1)
        GuiControl,,B, 3
        
        ;RunWait, python "C:\get_exchange_rate_v2.py", OutputVar
        ;RunWait, python "C:\Users\samagu0030\Desktop\get_exchange_rate_v2.py", OutputVar
        ;RunWait, python "C:\get_exchange_rate.py", OutputVar1, Hide
    }
return

Button종료:
    {
        
        ; isCheckPcc = IsCheckPcc("나철환", P180002538686, 01091700607)
        
        phonenNmber := 01091700607
        
        ;// 해외배송 시작 문자 보내기
        PhoneControlUtil_SendMessage(phonenNmber, g_StartInternationalDeliveryMessage())
        ;// 통관 시작 문자 보내기
        PhoneControlUtil_SendMessage(phonenNmber, g_StartCustomsClearanceMessage())
        ;// 국내 배송 시작 문자 보내기
        PhoneControlUtil_SendMessage(phonenNmber, g_StartKoreaDeliveryMessage())
        
        ;WhileFindChromeTab("Google 번역")
        
        ; 번역할 텍스트 복사
        ;ChromeTranslateGoogle("Best Seller")
        
        ; 번역된 텍스트 복사
        ;Send, ^a^c
        ;SleepTime(0.1)  ; 클립보드 처리를 위한 대기 시간
        
        ;translatedText := DeepLTranslate("나라", "EN")  ; 번역할 언어 코드 입력
        
        ;IsCoIorByPosionRange(0x000000, g_FindPosionType_upR(), 10)
        ;// 7ae40e64-3bef-73d2-c357-e5565374bf69:fx
        
      
        isTest := false
        ExitApp
    }
return
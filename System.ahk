#Include EnvData.ahk
#Include GlobalData.ahk
#Include Util.ahk
#Include ProductRegistration.ahk
#Include Test\Test_Ing.ahk
#Include PhoneControlUtil.ahk

;// 상품 번호 칸
xl_A(row)
{
	return "A" . row
}

;// 상품 url 칸
xl_B(row)
{
	return "B" . row
}

;// 상품 구매 url 칸
xl_C(row)
{
	return "C" . row
}

;// 브랜드 칸
xl_E(row)
{
	return "E" . row
}

;// 색 RGB(16진수) 리스트 칸
xl_F(row)
{
	return "F" . row
}

;// 색명(사이즈 리스트) 칸
xl_G(row)
{
	return "G" . row
}

;// 업데이트 시간 칸
xl_H(row)
{
	return "H" . row
}

;// 체크 시간 칸
xl_I(row)
{
	return "I" . row
}

;// 체크 상태 칸
xl_J(row)
{
	return "J" . row
}

;// 이전 색RGB(16진수) 리스트 칸
xl_K(row)
{
	return "K" . row
}

;// 이전 색명(사아즈 리스트) 칸
xl_L(row)
{
	return "L" . row
}

xl_O(row)
{
	return "O" . row
}

;// 마지막 실행 시켰던 라인
xl_P(row)
{
	return "P" . row
}

;// 마지막 실행 시켰던 라인의 시간 입력
xl_Q(row)
{
	return "Q" . row
}

;// 상품 이름
xl_T(row)
{
	return "T" . row
}

;// 상품 원본 가격
xl_U(row)
{
	return "U" . row
}

SaveWorksheet(xlWorkbook)
{
	xlWorkbook.SaveCopyAs(g_DefaultPath() . "\엑셀\마구싸5 구매루트_복제.xlsx")
	xlWorkbook.Save()
}

GetElementsData()
{
	Send, {F12}
	SleepTime(2)
	if(WhileFoundImage("크롬\Elements에 html"))
	{
		MoveAtWhileFoundImage("크롬\Elements에 html", 5, 5)
		SleepTime(0.5)
		NowMouseClickRight()
		SleepTime(3)
		CoordMode, Mouse, Screen
		MouseGetPos, currentX, currentY
		MoveAtWhileFoundImage("크롬\Elements에 html의 copy", 5, 5,10, 1, currentX, currentY)
		SleepTime(1)
		MoveAtWhileFoundImage("크롬\Elements에 html의 copy에 copy element", 5, 5,10, 1, currentX, currentY)
		SleepTime(1)
		NowMouseClick()
		SleepTime(3)
		outElementsData := Clipboard
		SleepTime(0.5)
		return outElementsData
	}
	return ""
}

ExcelSystemKill()
{
	;// 열려있는 엑셀 닫기(이유 : 열려 있는 엑셀 사용시 읽기 전용으로 해서 저장시 에러 나기 때문에
	Run, taskkill /F /IM "EXCEL.EXE"
	;// 시간 3초는 닫는데 걸리는 시간(제대로 안 닫히면 엑셀 여는에 문제가 있기 때문에)
	SleepTime(3)
}

;// 등록 된 상품 최신화
UpdateStoreWithColorInformation()
{
	TelegramSend("등록 된 상품 최신화 -- 시작")
	xlFile := g_DefaultPath() . "\엑셀\마구싸5 구매루트.xlsx"
	;// 엑셀 파일 열기
	xl := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xl.Visible := false
	xl.DisplayAlerts := false
	xlWorkbook := xl.Workbooks.Open(xlFile, 0, false, , "")
	xlWorksheet := xlWorkbook.Sheets(1)

	; UsedRange을 통해 데이터가 있는 범위 가져오기
	xlUsedRange := xlWorksheet.UsedRange
	lastRow := xlUsedRange.Rows.Count

	row := Round(xlWorksheet.Range(xl_P("1")).value)

	krwUsd := KRWUSD()
	krwEur := KRWEUR()

	while (true)
	{
		++row

		if(row > lastRow)
		{
			break
		}

		Debug("row(" . row . ") / lastRow(" . lastRow . ")")
		if (Mod(row, 10) = 0)
		{
			TelegramSend("__ row(" . row . ") / lastRow(" . lastRow . ")" . GetFormattedCurrentDateTime())
		}

		if(InStr(xlWorksheet.Range(xl_J(row)).value, "품절 상태로 변경 완료"))
		{
			xlWorksheet.Range(xl_P("1")).value := row
			xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
			SaveWorksheet(xlWorkbook)
			continue
		}

		url := xlWorksheet.Range(xl_C(row)).value

		if(InStr(url, "www.ugg.com"))
		{
			TelegramSend("row(" . row . ") / lastRow(" . lastRow . ")" . GetFormattedCurrentDateTime())

			isUpdateProduct := UpdateProductInfo_UGG(xlWorkbook, xlWorksheet, url, row, krwUsd)
			if(isUpdateProduct)
			{
				xlWorksheet.Range(xl_P("1")).value := row
				xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
				SaveWorksheet(xlWorkbook)
			}
			else
			{
				--row
			}

			continue
		}

		if(InStr(url, "www.mytheresa.com"))
		{
			TelegramSend("row(" . row . ") / lastRow(" . lastRow . ")" . GetFormattedCurrentDateTime())

			isUpdateProduct := UpdateProductInfoMoney_Mytheresa(xlWorkbook, xlWorksheet, url, row, krwEur)
			if(isUpdateProduct)
			{
				xlWorksheet.Range(xl_P("1")).value := row
				xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
				SaveWorksheet(xlWorkbook)
			}
			else
			{
				--row
			}

			continue
		}
	}

	;// true를 전달하여 저장 여부 설정
	xlWorkbook.Close(true)
	xl.Quit()

	TelegramSend("등록 된 상품 최신화 -- 끝")
}

UpdateStoreWithColorInformationMoney_Mytheresa()
{
	xlFile := g_DefaultPath() . "\엑셀\마구싸5 구매루트.xlsx"
	;// 엑셀 파일 열기
	xl := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xl.Visible := false
	xl.DisplayAlerts := false
	xlWorkbook := xl.Workbooks.Open(xlFile, 0, false, , "")
	xlWorksheet := xlWorkbook.Sheets(1)

	; UsedRange을 통해 데이터가 있는 범위 가져오기
	xlUsedRange := xlWorksheet.UsedRange
	lastRow := xlUsedRange.Rows.Count

	row := Round(xlWorksheet.Range(xl_P("1")).value)

	krwEur := KRWEUR()

	while (true)
	{
		++row

		if(row > lastRow)
		{
			break
		}

		Debug("row(" . row . ") / lastRow(" . lastRow . ")")
		if (Mod(row, 10) = 0)
		{
			TelegramSend("__ row(" . row . ") / lastRow(" . lastRow . ")" . GetFormattedCurrentDateTime())
		}

		; 웹 브라우저 열기 및 상품 url로 이동
		url := xlWorksheet.Range(xl_C(row)).value
		if(!InStr(url, "www.mytheresa.com"))
		{
			xlWorksheet.Range(xl_P("1")).value := row
			xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
			SaveWorksheet(xlWorkbook)
			continue
		}

		if(InStr(xlWorksheet.Range(xl_J(row)).value, "품절 상태로 변경 완료"))
		{
			xlWorksheet.Range(xl_P("1")).value := row
			xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
			SaveWorksheet(xlWorkbook)
			continue
		}

		TelegramSend("row(" . row . ") / lastRow(" . lastRow . ")" . GetFormattedCurrentDateTime())

		isUpdateProduct := UpdateProductInfoMoney_Mytheresa(xlWorkbook, xlWorksheet, url, row, krwEur)
		if(isUpdateProduct)
		{
			xlWorksheet.Range(xl_P("1")).value := row
			xlWorksheet.Range(xl_Q("1")).value := GetFormattedCurrentDateTime()
			SaveWorksheet(xlWorkbook)
		}
		else
		{
			--row
		}
	}

	;// true를 전달하여 저장 여부 설정
	xlWorkbook.Close(true)
	xl.Quit()
}

UpdateProductInfo_UGG(xlWorkbook, xlWorksheet, url, row, krwUsd)
{
	data := GetUggData(url, krwUsd)

	;// UGG에 사이즈 정보로 정보 취합
	useMoney := data.useMoney

	;// 이중 배열
	arraySizesAndImgUrls := data.arraySizesAndImgUrls

	;// 기존 것과 같은지 비교(같으면 스마트 스토어에 하지 않기 위함)
	before_SaveColorList := xlWorksheet.Range(xl_F(row)).value
	Debug("before_SaveColorList : " . before_SaveColorList)

	;// 기존 색 이름 과 사아즈를 변수로 저장
	before_SaveColorNameDoubleArray := xlWorksheet.Range(xl_G(row)).value
	Debug("before_SaveColorNameDoubleArray : " . before_SaveColorNameDoubleArray)

	;// 색이름 리스트 값
	colorNames := []
	Loop, % arraySizesAndImgUrls.Length()
	{
		colorNames.Push(arraySizesAndImgUrls[A_Index][1])
	}
	str_saveColorList := JoinArrayToString(colorNames)
	Debug("str_saveColorList : " . str_saveColorList)

	;// 색 이름 과 사아즈 리스트 값(이중 배열)
	str_saveColorNameDoubleArray := DoubleArrayToString(arraySizesAndImgUrls)
	Debug("str_saveColorNameDoubleArray : " . str_saveColorNameDoubleArray)

	if(str_saveColorNameDoubleArray = "" || useMoney = 0)
	{ ;// 색이 없은 경우 자체가 연결 되지 않거나 물건 자체가 없어졌을 경우
		;// 스마트 스토어 수정 화면까지 이동
		if(ManageAndModifyProducts(xlWorksheet, row) == false)
		{
			return false
		}

		;// 품절
		SoldOut(xlWorkbook, xlWorksheet, row)
	}
	else
	{
		if(before_SaveColorNameDoubleArray = str_saveColorNameDoubleArray && xlWorksheet.Range(xl_U(row)).value = useMoney )
		{ ;// 이전과 정보가 변함이 없을 경우(이전과 동일하다고 적고 다음으로 넘어감)
			xl_J_(xlWorkbook, xlWorksheet, row, "이전과 동일합니다.")
		}
		else
		{ ;// 이전과 달라졌음
			xl_J_(xlWorkbook, xlWorksheet, row, "이전과 동일하지 않아서 변경 하려고 합니다.")

			;// 스마트 스토어 수정 화면까지 이동
			if(ManageAndModifyProducts(xlWorksheet, row) == false)
			{
				return false
			}

			;// 가격 변동이 있으면 변경
			if(xlWorksheet.Range(xl_U(row)).value != useMoney)
			{
				;// 판매가 입력
				UpdateAndReturnSalePrice(data.korMony)
			}

			if(before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray)
			{
				;// 관세 부가 여부 체크
				customsDuty := useMoney >= 200

				;// 옵션 엑셀 세팅
				SetExcelOption(arraySizesAndImgUrls, customsDuty)

				xl_J_(xlWorkbook, xlWorksheet, row, "이전과 동일하지 않아서 변경 하려고 합니다.(옵션 엑셀 세팅 완료)")

				;// 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
				UpdateOptionsFromExcel(customsDuty)

				;// 색은 그대로 인 상태에서 사이즈 숫자만 바꿔서 상세 페이지 갱신 하지 않도록 처리
				if(before_SaveColorList != str_saveColorList)
				{
					;// HTML 으로 등록
					SetHTML(arraySizesAndImgUrls)
				}
			}

			SleepTime(1)
			ClickAtWhileFoundImage("스마트 스토어\상품 수정\저장하기", 5, 5)
			; SleepTime(5)
			; ClickAtWhileFoundImage("스마트 스토어\상품 수정\상품관리", 5, 5)
			; SleepTime(1)

			if(before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray && xlWorksheet.Range(xl_U(row)).value != useMoney)
			{
				xlWorksheet.Range(xl_U(row)).value := useMoney

				;// 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
				if(true)
				{
					;// 색 이름 리스트 표시
					xlWorksheet.Range(xl_F(row)).value := str_saveColorList
					Debug("str_saveColorList : " . str_saveColorList)

					;// 색 이름과 사아즈 리스트 표시
					xlWorksheet.Range(xl_G(row)).value := str_saveColorNameDoubleArray
					Debug("str_saveColorNameDoubleArray : " . str_saveColorNameDoubleArray)

					xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(이전과 동일하지 않아)(이전 값 등록 전)", true)

					;// 이전 색 이름 리스트 표시
					xlWorksheet.Range(xl_K(row)).value := before_SaveColorList
					Debug("before_SaveColorList : " . before_SaveColorList)

					;// 이전 색 이름과 사아즈 리스트 표시
					xlWorksheet.Range(xl_L(row)).value := before_SaveColorNameDoubleArray
					Debug("before_SaveColorNameDoubleArray : " . before_SaveColorNameDoubleArray)

					xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(이전과 동일하지 않아)(가격과 사이즈)")
				}
			}
			else
			{
				;// 가격 변동이 있으면 변경
				if(xlWorksheet.Range(xl_U(row)).value != useMoney)
				{
					xlWorksheet.Range(xl_U(row)).value := useMoney

					xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(가격만 변동)")
				}

				if(before_SaveColorNameDoubleArray != str_saveColorNameDoubleArray)
				{
					;// 입력 - (색 이름 리스트, 색 이름과 사아즈 리스트, 갱신 시간, 체크 시간, 체크 상태, 이전 색RGB(16진수) 리스트, 이전 색명(사아즈 리스트))
					if(true)
					{
						;// 색 이름 리스트 표시
						xlWorksheet.Range(xl_F(row)).value := str_saveColorList
						Debug("str_saveColorList : " . str_saveColorList)

						;// 색 이름과 사아즈 리스트 표시
						xlWorksheet.Range(xl_G(row)).value := str_saveColorNameDoubleArray
						Debug("str_saveColorNameDoubleArray : " . str_saveColorNameDoubleArray)

						xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(이전과 동일하지 않아)(이전 값 등록 전)", true)

						;// 이전 색 이름 리스트 표시
						xlWorksheet.Range(xl_K(row)).value := before_SaveColorList
						Debug("before_SaveColorList : " . before_SaveColorList)

						;// 이전 색 이름과 사아즈 리스트 표시
						xlWorksheet.Range(xl_L(row)).value := before_SaveColorNameDoubleArray
						Debug("before_SaveColorNameDoubleArray : " . before_SaveColorNameDoubleArray)

						xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(이전과 동일하지 않아)")
					}
				}
			}
		}
	}

	return true
}

UpdateProductInfoMoney_Mytheresa(xlWorkbook, xlWorksheet, url, row, krwEur)
{
	data := GetMytheresaData(url, krwEur)

	;// 스마트 스토어 수정 화면까지 이동
	if(ManageAndModifyProducts(xlWorksheet, row) == false)
	{
		return false
	}

	if(data.isSoldOut)
	{
		;// 품절
		SoldOut(xlWorkbook, xlWorksheet, row)
	}
	else
	{
		if (data.sizesLength = 0)
		{
			useMoney := data.useMoney

			;// 판매가 입력
			UpdateAndReturnSalePrice(data.korMony)

			SleepTime(1)
			ClickAtWhileFoundImage("스마트 스토어\상품 수정\저장하기", 5, 5)

			if(data.korMony != 0)
			{
				xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(가격만 변동)")
			}
			else
			{
				xl_J_(xlWorkbook, xlWorksheet, row, "가격이 0이 나왔습니다.")
			}
		}
		else
		{
			useMoney := data.useMoney
			arraySizesAndImgUrls := data.arraySizesAndImgUrls

			;// 관세 부가 여부 체크
			customsDuty := useMoney >= 150

			;// 옵션 엑셀 세팅
			SetExcelOption(arraySizesAndImgUrls, customsDuty)

			;// 1. 가격 세팅
			;// 2. 엑셀로 옵셥 세팅
			if(true)
			{
				;// 판매가 입력
				UpdateAndReturnSalePrice(data.korMony)

				;// 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
				UpdateOptionsFromExcel(customsDuty)

				SleepTime(1)
				ClickAtWhileFoundImage("스마트 스토어\상품 수정\저장하기", 5, 5)
				; SleepTime(5)
				; ClickAtWhileFoundImage("스마트 스토어\상품 수정\상품관리", 5, 5)
				; SleepTime(1)
			}

			if(data.korMony != 0)
			{
				xl_J_(xlWorkbook, xlWorksheet, row, "변경 완료(가격과 사이즈 변동)")
			}
			else
			{
				xl_J_(xlWorkbook, xlWorksheet, row, "가격이 0이 나왔습니다.")
			}
		}
	}

	return true
}

xl_J_(xlWorkbook, xlWorksheet, row, value, updateTime := false)
{
	if(updateTime)
	{
		;// 갱신 시간 표시
		xlWorksheet.Range(xl_H(row)).value := GetFormattedCurrentDateTime()
	}
	;// 체크 시간 표시
	xlWorksheet.Range(xl_I(row)).value := GetFormattedCurrentDateTime()
	;// 체크 상태 표시
	xlWorksheet.Range(xl_J(row)).value := value

	SaveWorksheet(xlWorkbook)
}

;// UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
GetNewProductURLs_UGG(name, url, filterUrls)
{
	TelegramSend("GetNewProductURLs_UGG()  " . name . " -- 시작")
	Run, chrome.exe %url%
	; "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
	WinWait, ugg
	SleepTime(10)

	;// 웹 제일 끝까지 스코롤 한다.
	while (true)
	{
		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		; DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, 100, uint, 0) ; 100 틱 스크롤 djq
		; SleepTime(1)
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, -10000, uint, 0) ; -10000 틱 스크롤 다운
		SleepTime(1.5)
		Send {Up}
		SleepTime(1)
		Send {Down}
		Send {Down}
		SleepTime(1)
		if(ClickAtWhileFoundImage("크롬\오른쪽 스트롤바가 제일 아래인 이미지", 0, 0, 1, 1, A_ScreenWidth - 200, A_ScreenHeight - 200) || ClickAtWhileFoundImage("크롬\오른쪽 스트롤바가 제일 아래인 이미지_v2", 0, 0, 1, 1, A_ScreenWidth - 200, A_ScreenHeight - 200)|| ClickAtWhileFoundImage("크롬\오른쪽 스트롤바가 제일 아래인 이미지_v3", 0, 0, 1, 1, A_ScreenWidth - 200, A_ScreenHeight - 200))
		{
			;// 상품 더 보기가 있는지 체크
			if(ClickAtWhileFoundImage("UGG\상품 리스트\상품 더 보기 버튼", 0, 0, 3))
			{
				SleepTime(5)
			}
			else
			{ ;// 상품이 더이상 없음
				Break
			}
		}
	}

	htmlElementsData := GetElementsData()
	;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
	Send ^w
	SleepTime(1)

	productUrls := []
	;// <a href=" 과 " class="js-pdp-link image-link pdp-link"> 중간에 있는 값
	productUrlLines := GetRegExMatche1List(htmlElementsData, "<a href=""(.*?)"" class=""js-pdp-link image-link pdp-link"">")
	Loop, % productUrlLines.Length()
	{
		splitList := StrSplit(productUrlLines[A_Index], ".html")
		if(splitList.Length() > 0)
		{
			productUrls.Push("https://www.ugg.com" . splitList[1] . ".html")
		}
		else
		{
			productUrls.Push(productUrlLines[A_Index])
		}
	}
	uniqueArr := []
	for index, value in productUrls
	{
		for index2, value2 in filterUrls
		{
			if (value = value2)
			{
				uniqueArr.Push(Value)
				break
			}
		}
	}

	for index, value in uniqueArr
	{
		ArrayRemove(productUrls, value)
	}

	TelegramSend("GetNewProductURLs_UGG()  " . name . " -- 끝")

	return [name, productUrls]
}

ArrayRemove(ByRef arr, value) {
	for index, element in arr
	{
		if (element = value)
		{
			arr.Remove(index)
			break
		}
	}
}

;// 신규 등록 할 UGG 목록을 엑셀에 정리
SetXlsxUGGNewProductURLs()
{
	TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 시작")
	xlFile := g_DefaultPath() . "\엑셀\마구싸5 구매루트.xlsx"
	;// 엑셀 파일 열기
	xl := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xl.Visible := false
	xl.DisplayAlerts := false
	xlWorkbook := xl.Workbooks.Open(xlFile, 0, false, , "")
	xlWorksheet := xlWorkbook.Sheets(1)

	; UsedRange을 통해 데이터가 있는 범위 가져오기
	xlUsedRange := xlWorksheet.UsedRange
	lastRow := xlUsedRange.Rows.Count

	Debug("start xlsx ugg url")
	; C 열의 데이터를 배열에 저장
	filterUrls := []
	Loop, % lastRow
	{
		url := xlWorksheet.Cells(A_Index, "C").Value
		if(InStr(url, "www.ugg.com"))
		{
			filterUrls.Push(url)
		}
	}
	TelegramSend("end xlsx ugg url Length : " . filterUrls.Length())

	xlWorkbook.Close()
	xl.Quit()

	Run, chrome.exe "https://www.ugg.com/women-footwear"
	; "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
	WinWait, ugg
	SleepTime(10)

	Debug("11111111111")
	;// UGG 현재 웹 창의 전체 상품 URL 리스트 정보 가져옴
	uggProductUrls := []
	uggProductUrls.Push(GetNewProductURLs_UGG("패션잡화 여성신발 부츠 미들부츠", "https://www.ugg.com/women-footwear/?prefn1=type&prefv1=boots%7Cclassic-boots%7Ccold-weather-boots", filterUrls)) ;// 부츠(미들부츠)
	uggProductUrls.Push(GetNewProductURLs_UGG("패션잡화 여성신발 샌들 뮬", "https://www.ugg.com/women-footwear/?prefn1=type&prefv1=dress-shoes%7Csandals", filterUrls)) ;// 샌들(뮬)
	uggProductUrls.Push(GetNewProductURLs_UGG("패션잡화 여성신발 슬리퍼", "https://www.ugg.com/women-footwear/?prefn1=type&prefv1=clogs%7Cslippers", filterUrls)) ;// 슬리퍼
	uggProductUrls.Push(GetNewProductURLs_UGG("패션잡화 여성신발 운동화 러닝화", "https://www.ugg.com/women-footwear/?prefn1=type&prefv1=sneakers", filterUrls)) ;// 운동화

	Send ^w
	SleepTime(1)

	Debug("22222222222222")
	xlFile := g_DefaultPath() . "\엑셀\추가 할 것들.xlsx"
	;// 엑셀 파일 열기
	xl := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xl.Visible := false
	xl.DisplayAlerts := false
	xlWorkbook := xl.Workbooks.Open(xlFile, 0, false, , "")
	xlWorksheet := xlWorkbook.Sheets(1)
	rowCount := xlWorksheet.UsedRange.Rows.Count

	; 모든 행을 삭제합니다.
	Loop % rowCount {
		xlWorksheet.Rows(1).Delete() ; 각 반복에서 첫 번째 행을 삭제합니다.
	}

	allCount := 0
	for index, item in uggProductUrls {
		for index2, item2 in item[2] {
			allCount++
			xlWorksheet.Range("A" . allCount).value := "UGG"
			xlWorksheet.Range("B" . allCount).value := item[1]
			xlWorksheet.Range("C" . allCount).value := item2
		}
	}

	xlWorkbook.Close(true)
	xl.Quit()

	TelegramSend("신규 등록 할 UGG 목록을 엑셀에 정리 -- 끝")
}

;// HTML 으로 등록
SetHTML(arraySizesAndImgUrls, isAdd := false)
{
	;// 상세설명 찾아서 그 아래로 ONE 원형 검색이 존재 하는데 체크(이때는 아래로 조금씩 내리기)
	WheelAndMoveAtWhileFoundImage("스마트 스토어\상품 수정\상세 설명")
	findIndex := 0
	while (true) ; 무한 루프
	{
		if(MoveAtWhileFoundImage("스마트 스토어\상품 수정\녹색 상세설명", 0, 0, 1))
		{
			findIndex := 1
			break
		}
		else
		{
			if(MoveAtWhileFoundImage("스마트 스토어\상품 수정\녹색 상세설명_v2", 0, 0, 1))
			{
				findIndex := 2
				break
			}
			else
			{
				DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, -500, uint, 0)
				SleepTime(1)
			}
		}
	}

	if(MoveAtWhileFoundImage("스마트 스토어\상품 수정\HTML 작성", 0, 0, 1))
	{
		NowMouseClick()
		if(isAdd = false)
		{
			SleepTime(1)
			ClickAtWhileFoundImage("스마트 스토어\상품 수정\확인", 5, 5)
		}
	}
	SleepTime(1)
	if(findIndex == 1)
	{
		ClickAtWhileFoundImage("스마트 스토어\상품 수정\녹색 상세설명", 100, -150, 1)
	}
	else if(findIndex == 2)
	{
		ClickAtWhileFoundImage("스마트 스토어\상품 수정\녹색 상세설명_v2", 100, -150, 1)
	}
	SleepTime(1)
	Send ^a
	SleepTime(1)
	Send {Delete}
	SleepTime(1)
	;// html 내용 작성
	if(true)
	{
		htmlData := "<div style=""text-align: center;"">"
		htmlData .= "<img src=""https://nacharhan.github.io/photo/2.png""/>"
		htmlData .= "`r`n"
		Loop % arraySizesAndImgUrls.MaxIndex()
		{
			colorName := arraySizesAndImgUrls[A_Index][1]
			imgUrls := arraySizesAndImgUrls[A_Index][3]

			htmlData .= "<div style=""text-align: center;"">"
			htmlData .= "<div><span style=""font-size: 30px;"">" . colorName . "</span></div>"
			htmlData .= "`r`n"
			Loop % imgUrls.MaxIndex()
			{
				htmlData .= "<div style=""text-align: center;"">"
				htmlData .= "<img src=""" . imgUrls[A_Index] . """/>"
				htmlData .= "`r`n"
			}
		}

		htmlData .= "<div style=""text-align: center;"">"
		htmlData .= "<img src=""https://nacharhan.github.io/photo/11.png""/>"
		htmlData .= "`r`n"
	}
	Clipboard := htmlData
	SleepTime(1)
	Send ^v
	SleepTime(1)
}

;// 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
UpdateOptionsFromExcel(customsDuty)
{
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\옵션", 0, 0, -500)
	SleepTime(1)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\엑셀 일괄등록", 0, 0)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\확인", 0, 0, 2)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\엑셀 일괄등록하기", 0, 0)
	SleepTime(4)
	ClickAtWhileFoundImage("스마트 스토어\열기\즐겨찾기에 엑셀 폴더", 40, 5)
	SleepTime(1)
	if(customsDuty = 1)
	{
		DoubleClickAtWhileFoundImage("스마트 스토어\열기\옵션 세팅된 엑셀", 5, 5)
	}
	else
	{
		DoubleClickAtWhileFoundImage("스마트 스토어\열기\옵션 세팅된 엑셀2", 5, 5)
	}
	SleepTime(1)
}

AddOneProduct_Ugg(xlWorkbookAddBefore, xlWorkbookAdd, addOneProductSuccess, krwUsd)
{
	xlWorksheetAddBefore := xlWorkbookAddBefore.Sheets(1)
	xlWorksheetAdd := xlWorkbookAdd.Sheets(1)

	url := xlWorksheetAddBefore.Range("C1").value

	data := GetUggData(url, krwUsd)

	;// UGG에 사이즈 정보로 정보 취합
	useMoney := data.useMoney

	;// 이중 배열
	arraySizesAndImgUrls := data.arraySizesAndImgUrls

	;// 상품 이름
	title := data.title

	FolderToDelete(g_DefaultPath() . "\DownloadImage")

	if(arraySizesAndImgUrls.Length() = 0)
	{
		TelegramSend("arraySizesAndImgUrls.Length() = 0   url : " . url)
		;// 등록해야 될 것에서 삭제
		xlWorksheetAddBefore.Rows(1).Delete()
		xlWorkbookAddBefore.Save()
		return true
	}

	if(arraySizesAndImgUrls.Length() >= 1){
		imgUrls := arraySizesAndImgUrls[1][3]
		Loop % imgUrls.MaxIndex(){
			DownloadImageUrl(imgUrls[A_Index], A_Index)
		}
	}

	SleepTime(1)
	Run, chrome.exe "https://sell.smartstore.naver.com/#/products/create"
	SleepTime(0.5)
	Send ^{Tab}
	SleepTime(0.5)
	Send ^w
	SleepTime(2)

	if(ClickAtWhileFoundImage("스마트 스토어\로그인하기", 5, 5, 1))
	{
		SleepTime(2)
	}
	if(ClickAtWhileFoundImage("스마트 스토어\로그인", 5, 5, 1))
	{
		SleepTime(1)
		Run, chrome.exe "https://sell.smartstore.naver.com/#/products/create"
		SleepTime(0.5)
		Send ^{Tab}
		SleepTime(0.5)
		Send ^w
		SleepTime(2)
	}

	SleepTime(2)
	Send, {Esc}
	SleepTime(1)

	; if(!addOneProductSuccess)
	; {
	; 	ClickAtWhileFoundImage("스마트 스토어\상품 수정\이전 내용 불러오기 확인", -80, 5, 5)
	; }

	;// 상품 등록시 기본 세팅 들
	ProductRegistrationDefaultSettings()

	;// 카테고리명 입력
	MoveAtWhileFoundImage("스마트 스토어\상품 수정\카테고리명 선택", 0, 50)
	SleepTime(0.5)
	NowMouseClick()
	Clipboard := xlWorksheetAddBefore.Range("B1").value
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)
	Send, {Enter}
	SleepTime(2)

	;// 상품명 입력
	MoveAtWhileFoundImage("스마트 스토어\상품 수정\상품명", 50, 85)
	SleepTime(1)
	NowMouseClick()
	SleepTime(0.5)
	Clipboard := title
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)

	;// 판매가 입력
	UpdateAndReturnSalePrice(data.korMony)

	;// 옵션 세팅
	if(true)
	{
		;// 관세 부가 여부 체크
		customsDuty := useMoney >= 200

		;// 옵션 엑셀 세팅
		SetExcelOption(arraySizesAndImgUrls, customsDuty)

		;// 상품 수정에서 옵션을 엑셀 파일로 일괄 등록
		UpdateOptionsFromExcel(customsDuty)
	}

	;// 이미지 등록(대표, 추가)
	IamgeRegistration_v2()

	;// HTML 으로 등록
	SetHTML(arraySizesAndImgUrls, true)

	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\저장하기", 5, 5)
	SleepTime(5)
	if(ClickAtWhileFoundImage("스마트 스토어\상품 수정\상품관리", -80, 5))
	{
		SleepTime(3)

		;// 동록한 엑셀에 기록(앞에 추가)
		xlWorksheetAdd.Rows(2).Insert()

		;// 상품 url
		GoToTheAddressWindow()
		SleepTime(0.5)
		addurl := CopyToClipboardAndGet()
		Debug("addurl : " . addurl)
		xlWorksheetAdd.Range(xl_B("2")).value := addurl
		;// 크롬 탭 닫기
		Send, ^w
		SleepTime(0.5)
		;// 상품 번호
		addUrlSplitArray := StrSplit(addurl, "/")
		if(addUrlSplitArray.Length() > 0)
		{
			xlWorksheetAdd.Range(xl_A("2")).value := addUrlSplitArray[addUrlSplitArray.Length()]
		}

		xlWorksheetAdd.Range(xl_C("2")).value := url
		;// 상품명 기재
		xlWorksheetAdd.Range(xl_T("2")).value := title
		;// 가격
		xlWorksheetAdd.Range(xl_U("2")).value := useMoney
		;// 브랜드
		brand := xlWorksheetAddBefore.Range("A2").value
		xlWorksheetAdd.Range(xl_E("2")).value := brand

		;// 색이름 리스트 값
		colorNames := []
		Loop, % arraySizesAndImgUrls.Length()
		{
			colorNames.Push(arraySizesAndImgUrls[A_Index][1])
		}
		str_saveColorList := JoinArrayToString(colorNames)
		Debug("str_saveColorList : " . str_saveColorList)

		xlWorksheetAdd.Range(xl_F("2")).value := str_saveColorList

		;// 색 이름 과 사아즈 리스트 값(이중 배열)
		str_saveColorNameDoubleArray := DoubleArrayToString(arraySizesAndImgUrls)
		Debug("str_saveColorNameDoubleArray : " . str_saveColorNameDoubleArray)

		xlWorksheetAdd.Range(xl_G("2")).value := str_saveColorNameDoubleArray

		xl_J_(xlWorkbookAdd, xlWorksheetAdd, 2, "신규 등록", true)

		xlWorkbookAdd.Save()

		;// 등록해야 될 것에서 삭제
		xlWorksheetAddBefore.Rows(1).Delete()
		xlWorkbookAddBefore.Save()

		return true
	}
	else
	{
		TelegramSend("++++++++++++++  이름이 입력 안됬음  왜지??")
		;// 이름이 입력 안됬음  왜지??
		ClickAtWhileFoundImage("스마트 스토어\상품 수정\취소", 5, 5)
		SleepTime(1)
		ClickAtWhileFoundImage("스마트 스토어\상품 수정\상품취소 유실 확인", 5, 5)
		SleepTime(1)

		return false
	}
}

;// 추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기
AddDataFromExcel_Ugg()
{
	TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 시작")
	xlFileAddBefore := g_DefaultPath() . "\엑셀\추가 할 것들.xlsx"
	;// 엑셀 파일 열기
	xlAddBefore := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xlAddBefore.Visible := false
	xlAddBefore.DisplayAlerts := false
	xlWorkbookAddBefore := xlAddBefore.Workbooks.Open(xlFileAddBefore, 0, false, , "")
	xlWorksheetAddBefore := xlWorkbookAddBefore.Sheets(1)

	xlFileAdd := g_DefaultPath() . "\엑셀\마구싸5 구매루트.xlsx"
	;// 엑셀 파일 열기
	xlAdd := ComObjCreate("Excel.Application")
	;// Excel을 보이게 설정 할지 여부
	xlAdd.Visible := false
	xlAdd.DisplayAlerts := false
	xlWorkbookAdd := xlAdd.Workbooks.Open(xlFileAdd, 0, false, , "")

	; UsedRange을 통해 데이터가 있는 범위 가져오기
	rowCountAddBefore := xlWorksheetAddBefore.UsedRange.Rows.Count

	krwUsd := KRWUSD()

	;// 모두 반복
	count := 1
	Loop % rowCountAddBefore {
		addOneProductSuccess := true
		while(true)
		{
			Debug(count . "/" . rowCountAddBefore)
			TelegramSend(count . "/" . rowCountAddBefore)
			addOneProductSuccess := AddOneProduct_Ugg(xlWorkbookAddBefore, xlWorkbookAdd, addOneProductSuccess, krwUsd)
			if(addOneProductSuccess)
			{
				++count
				break
			}
		}
	}

	xlAdd.Quit()
	xlAddBefore.Quit()
	TelegramSend("추가할 엑셀 정보를 가지고 실제로 스마트스토어에 등록하기 -- 끝")
}

;// url에 필요 정보 가져오기
GetUggData(url, exchangeRate, onlyUseMoney := false)
{
	;// UGG에 사이즈 정보로 정보 취합
	useMoney := 0
	korMony := 0

	;// 이중 배열
	arraySizesAndImgUrls := []

	;// 상품 이름
	title := ""

	if(true)
	{
		;// 1145990은 url의 끝에 .html 전에 있는 값
		urlSplitArray := StrSplit(url, "/")
		if(urlSplitArray.Length() > 0)
		{
			productNumber := StrReplace(urlSplitArray[urlSplitArray.Length()], ".html", "")
		}
		else
		{
			productNumber := url
		}

		Run, chrome.exe %url%
		; "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
		WinWait, ugg
		SleepTime(10)
		htmlElementsData := GetElementsData()
		;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
		Send ^w
		SleepTime(1)

		;// 상품 이름
		if (RegExMatch(htmlElementsData, "<div\s+class\s*=\s*""sticky-toolbar__content"">\s*<span>([^<]*)</span>", match))
		{
			title := match1
			Debug("title : " . title)
		}

		;// <div class="sticky-toolbar__content"
		startPos := RegExMatch(htmlElementsData, "<div\s+class\s*=\s*""sticky-toolbar__content""", , 1)
		;// aria-labelledby="size" 까지에서 찾기
		endPos := RegExMatch(htmlElementsData, "aria-labelledby\s*=\s*""size""", , startPos)
		if (startPos && endPos)
		{
			contentValue := SubStr(htmlElementsData, startPos, endPos - startPos)
			useMoneys := GetRegExMatche1List(contentValue, "\$(.+)")
			if(useMoneys.Length() > 0)
			{
				useMoney := useMoneys[useMoneys.Length()] + 0
				Debug("useMoney : " . useMoney)
			}

			korMony := GetKorMony(useMoney, exchangeRate)

			if(onlyUseMoney = false)
			{
				urlEndColorNames := GetRegExMatche1List(contentValue, "<span data-attr-value=""([^""]*)"" class=""color-value swatch swatch-circle")
				colorNames := GetRegExMatche1List(contentValue, "data-attr-color-swatch=""[^""]*""\s*title=""([^""]*)""")

				if(urlEndColorNames.Length() = colorNames.Length())
				{
					Loop, % urlEndColorNames.Length()
					{
						;// .html?dwvar_1145990_color=BCDR 제일 뒤에 색 정보 적어서 url 열 수 있음
						;// 색 위치로 클릭하는 것보다 url 열는 것이 더 낫다고 생각됨
						colorUrl := url . "?dwvar_" . productNumber . "_color=" . urlEndColorNames[A_Index]

						Debug("urlEndColorNames[" . A_Index . "] : " . urlEndColorNames[A_Index])
						Debug("colorNames[" . A_Index . "] : " . colorNames[A_Index])

						Run, chrome.exe %colorUrl%
						; "ugg"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
						WinWait, ugg
						SleepTime(10)
						colorUrlHtmlElementsData := GetElementsData()
						;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
						Send ^w
						SleepTime(1)

						;// 이미지 url 알아오는 것
						if(true)
						{
							imgBigUrls := []
							imgUrls := GetRegExMatche1List(colorUrlHtmlElementsData, "<img[^>]+data-srcset=""([^""]+)""[^>]+>")
							Loop, % imgUrls.Length()
							{
								splitArray := StrSplit(imgUrls[A_Index], " , ")
								if(splitArray.Length() > 0)
								{
									;// 제일 끝에 것으로 하는 이유는 이미지 제일 큰 Url 이기 때문에
									if(RegExMatch(splitArray[splitArray.Length()], "https:(.+)\.png", match))
									{
										imgBigUrls.Push(match)
									}
								}
							}
						}

						;// 구매 가능한 사이즈 구하기
						if(true)
						{
							sizes := []
							;// <div class="sticky-toolbar__content"> 는 뒤에 줄에 이는 곳 부터 찾기
							startPos := RegExMatch(colorUrlHtmlElementsData, "<div\s+class\s*=\s*""sticky-toolbar__content""", match, 1)
							;// 정규식으로 특정 줄에 options-select과 https://www과 _1145990_ 과 data-attr-value 이 포함되는 줄이 있는지 체크
							sizeLines := GetRegExMatcheList(colorUrlHtmlElementsData, "options-select\s*.*value=""(https:\/\/www\.ugg\.com\/on\/demandware\.store\/Sites-UGG-US-Site\/en_US\/Product-Variation\?dwvar_" productNumber "_color=[^""]+)"".*data-attr-value=""([^""]+)""", startPos)
							Loop, % sizeLines.Length()
							{
								;// options-select 와 value= 사이에 값이 없는 것만 찾기
								;// 사이에 값이 있으면 그 사이즈는 없는 재고가 없는 것 임
								if (RegExMatch(sizeLines[A_Index], "options-select\s*""\s*value=""([^""]*)"""))
								{
									if (RegExMatch(sizeLines[A_Index], "data-attr-value=""([^""]+)""", match))
									{
										extractedValue := match1
										; 숫자로 판단되면 앞의 0 제거
										if (extractedValue ~= "^[0-9]+(\.[0-9]+)?$")
										{
											extractedValue := +extractedValue
											if (extractedValue ~= "^[0-9]+$")
											{
												; 숫자로 변환하여 앞의 0 제거
												extractedValue := extractedValue + 0
											}
											else
											{
												; 앞의 0이 있는 경우 제거
												if (SubStr(extractedValue, 1, 1) = "0" && SubStr(extractedValue, 2, 1) != ".")
												{
													extractedValue := SubStr(extractedValue, 2)
												}
											}

											sizeData := "US_" . extractedValue . "(" . GetUggKorSize(extractedValue) . ")"
											Debug("size : " . sizeData)
											sizes.Push(sizeData)
										}
										else
										{
											sizes.Push(extractedValue)
										}
									}
								}
							}
						}

						if(sizes.Length() > 0 && imgBigUrls.Length() > 0)
						{
							; 문자열의 길이가 25보다 큰지 확인
							if (StrLen(colorNames[A_Index]) > 25) {
								; 25자까지만 잘라내기
								colorName := SubStr(colorNames[A_Index], 1, 25)
							}
							else{
								colorName := colorNames[A_Index]
							}
							arraySizesAndImgUrls.Push([colorName, sizes, imgBigUrls])
						}
					}
				}
			}
		}
	}

	values := {}
	values.useMoney := useMoney
	values.korMony := korMony
	values.arraySizesAndImgUrls := arraySizesAndImgUrls
	values.title := title
	return values
}

GetMytheresaData(url, exchangeRate)
{
	;// 사이즈 정보로 정보 취합
	useMoney := 0
	korMony := 0

	;// 이중 배열
	arraySizesAndImgUrls := []

	;// 상품 이름
	title := ""

	isSoldOut := false

	if(true)
	{
		Run, chrome.exe %url%
		; "mytheresa"이라는 문자열을 포함하는 Chrome 창이 나타날 때까지 대기
		; WinWait, mytheresa
		SleepTime(10)
		htmlElementsData := GetElementsData()
		;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
		Send ^w
		SleepTime(1)

		if (RegExMatch(htmlElementsData, """PriceSpecification"": {\s+""price"": (\d+)", match))
		{
			useMoney := match1
		}

		korMony := GetKorMony(useMoney, exchangeRate)

		;// 상품 이름
		if (RegExMatch(htmlElementsData, """priceCurrency"":\s*""([A-Z]{3})""", match))
		{
			priceCurrency := match1
		}

		if (RegExMatch(htmlElementsData, ">Sold Out<", match))
		{
			isSoldOut := true
		}

		if (RegExMatch(htmlElementsData, "error404__title", match))
		{
			isSoldOut := true
		}

		sizeLines := GetRegExMatche1List(htmlElementsData, "<span class=""sizeitem__label"">(.*?)</span>")

		sizes := []
		Loop, % sizeLines.Length()
		{
			sizeData := sizeLines[A_Index] . "(" . GetMytheresaKorSize(sizeLines[A_Index]) . ")"
			Debug("size : " . sizeData)
			sizes.Push(sizeData)
		}

		colorName := "One Color"
		imgBigUrls := []
		arraySizesAndImgUrls.Push([colorName, sizes, imgBigUrls])
	}

	values := {}
	values.useMoney := useMoney
	values.korMony := korMony
	values.arraySizesAndImgUrls := arraySizesAndImgUrls
	values.title := title
	values.sizesLength := sizes.Length()
	values.isSoldOut := isSoldOut
	return values
}

;// 스마트 스토어 수정 화면까지 이동
ManageAndModifyProducts(xlWorksheet, row)
{
	SleepTime(1)
	Run, chrome.exe "https://sell.smartstore.naver.com/#/products/origin-list"
	SleepTime(0.5)
	Send ^{Tab}
	SleepTime(0.5)
	Send ^w
	SleepTime(2)

	if(ClickAtWhileFoundImage("스마트 스토어\로그인하기", 5, 5, 1))
	{
		SleepTime(2)
	}
	if(ClickAtWhileFoundImage("스마트 스토어\로그인", 5, 5, 1))
	{
		SleepTime(1)
		Run, chrome.exe "https://sell.smartstore.naver.com/#/products/origin-list"
		SleepTime(0.5)
		Send ^{Tab}
		SleepTime(0.5)
		Send ^w
		SleepTime(2)
	}

	SleepTime(2)
	Send, {Esc}
	SleepTime(1)

	; if(WhileFoundImage("스마트 스토어\상품 조회-수정 선택된 상태", 2))
	; {
	; 	ClickAtWhileFoundImage("스마트 스토어\상품 등록", 10, 10)
	; 	SleepTime(2)
	; 	ClickAtWhileFoundImage("스마트 스토어\상품 조회-수정", 10, 10)
	; 	SleepTime(2)
	; }
	; else
	; {
	; 	ClickAtWhileFoundImage("스마트 스토어\상품 관리", 10, 10)
	; 	SleepTime(1)
	; 	ClickAtWhileFoundImage("스마트 스토어\상품 조회-수정", 10, 10)
	; 	SleepTime(2)
	; }

	;// 상품 조회해서 상품 수정 화면으로 이동
	if(true)
	{
		ClickAtWhileFoundImage("스마트 스토어\상품 조회\상품번호", 150, 10)
		SleepTime(0.5)
		Clipboard := Round(xlWorksheet.Range(xl_A(row)).value)
		SleepTime(0.5)
		;// 상품번호 붙여넣기
		Send, ^v
		SleepTime(0.5)
		ClickAtWhileFoundImage("스마트 스토어\상품 조회\검색", 0, 0)
		SleepTime(1)
		ClickAtWhileFoundImage("스마트 스토어\상품 조회\수정", 0, 0)
		SleepTime(2)
	}

	; ClickAtWhileFoundImage("팀 뷰어\라이선스 구매", 150, 5, 1)

	;// "스마트 스토어\네트워크 불안정 느낌표"
	if(ClickAtWhileFoundImage("스마트 스토어\네트워크 불안정", 0, 0, 1))
	{
		return false
	}

	if(ClickAtWhileFoundImage("스마트 스토어\상품 수정\KC인증", 0, 0, 2))
	{
		ClickAtWhileFoundImage("스마트 스토어\상품 수정\KC인증 닫기", 10, 10, 2)
	}

	return true
}

;// 품절
SoldOut(xlWorkbook, xlWorksheet, row)
{
	;// 품절
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\옵션", 0, 0, -500)
	SleepTime(1)
	WheelAndMoveAtWhileFoundImage("스마트 스토어\상품 수정\옵션에서 선택형", 0, 0, -500)
	SleepTime(1)
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, optionEndX, optionEndY ; 현재 마우스 위치 얻기
	SleepTime(0.5)
	MoveAtWhileFoundImage("스마트 스토어\상품 수정\옵션", 0, 0, 2, 1, optionEndX - 50, optionEndY - 150)
	SleepTime(1)
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, optionStartX, optionStartY ; 현재 마우스 위치 얻기
	SleepTime(0.5)
	ScreenMouseMove(optionEndX + 1000, optionEndY + 100)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\옵션에서 선택형에서 설정함 상태", 0, 0, 2,1, optionStartX, optionStartY, optionEndX + 1000, optionEndY + 100)
	WheelAndMoveAtWhileFoundImage("스마트 스토어\상품 수정\재고수량에 개", 0, 0, 500)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\재고수량에 개", -80, 0)
	SleepTime(1)
	Send ^a
	SleepTime(0.5)
	Send 0
	SleepTime(1.5)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\저장하기", 5, 5)
	; SleepTime(5)
	; ClickAtWhileFoundImage("스마트 스토어\상품 수정\상품관리", 5, 5)
	; SleepTime(1)

	xl_J_(xlWorkbook, xlWorksheet, row, "품절 상태로 변경 완료", true)

	TelegramSend("품절 상태로 변경 완료   row(" . row . ") ")
}

UpdateAndReturnSalePrice(korMony)
{
	;// 판매가 입력
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\판매가", 250, 85)
	SleepTime(0.5)
	NowMouseClick()
	SleepTime(0.5)
	Send, ^a
	SleepTime(0.5)
	Send, {Delete}
	SleepTime(0.5)
	Clipboard := korMony
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)
}
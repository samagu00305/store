#Include EnvData.ahk
#Include GlobalData.ahk

;// 현재 미국 환율 정보 출력
KRWUSD()
{
	http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	http.Open("GET", "https://quotation-api-cdn.dunamu.com/v1/forex/recent?codes=FRX.KRWUSD", false)
	http.Send()
	response := http.ResponseText ; 응답 받기
	RegExMatch(response, "basePrice"":(\d+\.\d+)", output)
	RegExMatch(output, "(\d+\.\d+)", value)
	if(value = 0)
	{
		Debug("******** Error  ---- KRWUSD() ==== 0  ----  Error")
		TelegramSend("******** Error  ---- KRWUSD() ==== 0  ----  Error")
	}
	return value
}

;// 현재 유료 환율 정보 출력
KRWEUR()
{
	http := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	http.Open("GET", "https://quotation-api-cdn.dunamu.com/v1/forex/recent?codes=FRX.KRWEUR", false)
	http.Send()
	response := http.ResponseText ; 응답 받기
	RegExMatch(response, "basePrice"":(\d+\.\d+)", output)
	RegExMatch(output, "(\d+\.\d+)", value)
	if(value = 0)
	{
		Debug("******** Error  ---- KRWEUR() ==== 0  ----  Error")
		TelegramSend("******** Error  ---- KRWEUR() ==== 0  ----  Error")
	}
	return value
}

;// 현재 위치에서 마우스 PixelColor
GetNowMousePixelColor()
{
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, vx, vy
	CoordMode, Pixel, Screen
	PixelGetColor, OutputPixelColor, %vx%, %vy%, RGB
	return OutputPixelColor
}

ScreenMouseMove(x, y)
{
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	mousemove, %x%, %y%
}

;// 현재 위치에서 마우스 이동
NowMouseMove(addX, addY)
{
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, currentX, currentY ; 현재 마우스 위치 얻기
	currentX += addX
	currentY += addY
	mousemove, %currentX%, %currentY%
}

;// 현재 위치에서 마우스 클릭
NowMouseClick()
{
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, currentX, currentY ; 현재 마우스 위치 얻기
	MouseClick, Left, %currentX%, %currentY%, 1 ; 현재 마우스 위치에서 왼쪽 클릭
}

;// 현재 위치에서 마우스 오른쪽 클릭
NowMouseClickRight()
{
	CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
	MouseGetPos, currentX, currentY ; 현재 마우스 위치 얻기
	MouseClick, Right, %currentX%, %currentY%, 1 ; 현재 마우스 위치에서 왼쪽 클릭
}

FoundImage(imageName)
{
	FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY)
	if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
	{
		return true
	}
	return false
}

;// 이미지 찾을 때 까지 대기
WhileFoundImage(imageName, findMaxCount := 10, delayTime := 1, searchStart_x := 0, searchStart_y := 0)
{
	findCount := 0
	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY, searchStart_x, searchStart_y)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			SleepTime(0.5)
			return true
		}
		SleepTime(delayTime)
		++findCount
		if(findMaxCount <= findCount)
		{
			Debug("Error WhileFoundImage - " . findMaxCount . "번을 찾았으나 실패했습니다. imageName : " . imageName)
			return false
		}
	}
}

;// 이미지 찾아서 마우스 이동
MoveAtWhileFoundImage(imageName, addX := 0, addY := 0, findMaxCount := 10, delayTime := 1, searchStart_x := 0, searchStart_y := 0)
{
	findCount := 0
	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY, searchStart_x, searchStart_y)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX
			Byref_foundY += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			return true
		}
		SleepTime(delayTime)
		++findCount
		if(findMaxCount <= findCount)
		{
			Debug("Error ClickAtWhileFoundImage - " . findMaxCount . "번을 찾았으나 실패했습니다. imageName : " . imageName)
			return false
		}
	}
}

;// 이미지 찾아서 클릭
ClickAtWhileFoundImage(imageName, addX := 0, addY := 0, findMaxCount := 10, delayTime := 1, searchStart_x := 0, searchStart_y := 0, searchEnd_x := -1, searchEnd_y := -1)
{
	findCount := 0
	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY, searchStart_x, searchStart_y, searchEnd_x, searchEnd_y)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX
			Byref_foundY += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			MouseClick, Left, %Byref_foundX%, %Byref_foundY%, 1
			return true
		}
		SleepTime(delayTime)
		++findCount
		if(findMaxCount <= findCount)
		{
			Debug("Error ClickAtWhileFoundImage - " . findMaxCount . "번을 찾았으나 실패했습니다. imageName : " . imageName)
			return false
		}
	}
}

;// 이미지 찾아서 더블클릭
DoubleClickAtWhileFoundImage(imageName, addX := 0, addY := 0)
{
	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX
			Byref_foundY += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			MouseClick, Left, %Byref_foundX%, %Byref_foundY%, 2
			return
		}
		SleepTime(1)
	}
}

;// 이미지 찾아서 Drag
DragAtFoundImage(imageName1, addX1, addY1, imageName2, addX2, addY2)
{
	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName1, Byref_errlevel, Byref_foundX, Byref_foundY)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX1
			Byref_foundY += addY1
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			Click down
			break
		}
		SleepTime(1)
	}

	while (true) ; 무한 루프
	{
		FindImage_Byref(imageName2, Byref_errlevel, Byref_foundX, Byref_foundY)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX2
			Byref_foundY += addY2
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			Click up
			break
		}
		SleepTime(1)
	}
}

;// 아래로 스크롤 하면서 이미지 찾아서 클릭
WheelAndClickAtWhileFoundImage(imageName, addX := 0, addY := 0, wheelMove := -1000, inputCount := -1)
{
	count := 0
	while (true)
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX
			Byref_foundY += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			MouseClick, Left, %Byref_foundX%, %Byref_foundY%, 1
			return
		}

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, wheelMove, uint, 0) ; wheelMove 틱 스크롤 다운

		SleepTime(1)

		if(inputCount != -1)
		{
			++count
			if(count >= inputCount)
			{
				return
			}

			if(count == 10)
			{
				TelegramSend("******** Error  WheelAndClickAtWhileFoundImage()   imageName : " . imageName)
			}
		}
	}
}

;// 아래로 스크롤 하면서 이미지 찾아서 클릭
WheelAndClickAtWhileFoundImage_v2(imageName_1, imageName_2, addX := 0, addY := 0, wheelMove := -1000)
{
	count := 0
	while (true)
	{
		FindImage_Byref(imageName_1, Byref_errlevel_1, Byref_foundX_1, Byref_foundY_1)
		if(Byref_errlevel_1 = g_ErrorType_Success() && Byref_foundX_1 != "" && Byref_foundY_1 != "")
		{
			Byref_foundX_1 += addX
			Byref_foundY_1 += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX_1%, %Byref_foundY_1%
			SleepTime(0.5)
			MouseClick, Left, %Byref_foundX_1%, %Byref_foundY_1%, 1
			return 1
		}

		FindImage_Byref(imageName_2, Byref_errlevel_2, Byref_foundX_2, Byref_foundY_2)
		if(Byref_errlevel_2 = g_ErrorType_Success() && Byref_foundX_2 != "" && Byref_foundY_2 != "")
		{
			Byref_foundX_2 += addX
			Byref_foundY_2 += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX_2%, %Byref_foundY_2%
			SleepTime(0.5)
			MouseClick, Left, %Byref_foundX_2%, %Byref_foundY_2%, 1
			return 2
		}

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, wheelMove, uint, 0) ; wheelMove 틱 스크롤 다운

		SleepTime(1)

		++count
		if(count == 10)
		{
			TelegramSend("******** Error  WheelAndClickAtWhileFoundImage_v2()   imageName_1 : " . imageName_1 . "    imageName_2 : " . imageName_2)
		}
	}

	return 0
}

;// 아래로 스크롤 하면서 이미지 찾아서 이동
WheelAndMoveAtWhileFoundImage(imageName, addX := 0, addY := 0, wheelMove := -1000)
{
	count := 0
	while (true)
	{
		FindImage_Byref(imageName, Byref_errlevel, Byref_foundX, Byref_foundY)
		if(Byref_errlevel = g_ErrorType_Success() && Byref_foundX != "" && Byref_foundY != "")
		{
			Byref_foundX += addX
			Byref_foundY += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX%, %Byref_foundY%
			SleepTime(0.5)
			return
		}

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, wheelMove, uint, 0) ; wheelMove 틱 스크롤 다운

		SleepTime(1)

		++count
		if(count == 10)
		{
			TelegramSend("******** Error  WheelAndMoveAtWhileFoundImage()   imageName : " . imageName)
		}
	}
}

;// 아래로 스크롤 하면서 이미지 찾아서 이동
WheelAndMoveAtWhileFoundImage_v2(imageName_1, imageName_2, addX := 0, addY := 0, wheelMove := -1000)
{
	count := 0
	while (true)
	{
		FindImage_Byref(imageName_1, Byref_errlevel_1, Byref_foundX_1, Byref_foundY_1)
		if(Byref_errlevel_1 = g_ErrorType_Success() && Byref_foundX_1 != "" && Byref_foundY_1 != "")
		{
			Byref_foundX_1 += addX
			Byref_foundY_1 += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX_1%, %Byref_foundY_1%
			SleepTime(0.5)
			return 1
		}

		FindImage_Byref(imageName_2, Byref_errlevel_2, Byref_foundX_2, Byref_foundY_2)
		if(Byref_errlevel_2 = g_ErrorType_Success() && Byref_foundX_2 != "" && Byref_foundY_2 != "")
		{
			Byref_foundX_2 += addX
			Byref_foundY_2 += addY
			CoordMode, Mouse, Screen
			mousemove, %Byref_foundX_2%, %Byref_foundY_2%
			SleepTime(0.5)
			return 2
		}

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, wheelMove, uint, 0) ; wheelMove 틱 스크롤 다운

		SleepTime(1)

		++count
		if(count == 10)
		{
			TelegramSend("******** Error  WheelAndClickAtWhileFoundImage_v2()   imageName_1 : " . imageName_1 . "    imageName_2 : " . imageName_2)
		}
	}

	return 0
}

;// 이미지를 화면 전체에서 검색해서 검색 정보를 Byref로 값을 넘김
FindImage_Byref(imageName, Byref errLevel, Byref foundX, Byref foundY, searchStart_x := 0, searchStart_y := 0, searchEnd_x := -1, searchEnd_y := -1)
{
	GuiControl,,B, 이미지 찾기 시작(%imageName%)
	CoordMode, Pixel, Screen
	Debug("FindImage_Byref의 이미지 검색 시작 imageName : " . imageName)
	if(searchEnd_x = -1)
	{
		searchEnd_x := A_ScreenWidth
	}
	if(searchEnd_y = -1)
	{
		searchEnd_y := A_ScreenHeight
	}
	Debug(imageName . " searchStart_x:" . searchStart_x . " searchStart_y:" . searchStart_y . " searchEnd_x:" . searchEnd_x . " searchEnd_y:" . searchEnd_y)
	ImageSearch, FoundX, FoundY, searchStart_x, searchStart_y, searchEnd_x, searchEnd_y,*Trans091A36 *40 %A_ScriptDir%\Image\%imageName%.png
	GuiControl,,B, 이미지 찾기 끝(%imageName%)
	Debug(imageName . " ErrorLevel:" . ErrorLevel . " FoundX:" . FoundX . " FoundY:" . FoundX)
	errLevel := ErrorLevel
	foundX := FoundX
	foundY := FoundY
}

;// 이미지 서치해서 있는지 알려주는 함수
IsImageSearch(imageName, searchStart_x := 0, searchStart_y := 0)
{
	GuiControl,,B,IsImageSearch imageName(%imageName%) %searchStart_x%, %searchStart_y%
	CoordMode, pixel, screen
	Debug("IsImageSearch의 이미지 검색 시작 imageName : " . imageName)
	ImageSearch, FoundX, FoundY, searchStart_x, searchStart_y, A_ScreenWidth, A_ScreenHeight,*Trans091A36 *40 %A_ScriptDir%\Image\%imageName%.png
	Debug(imageName . " ErrorLevel:" . ErrorLevel . " FoundX:" . FoundX . " FoundY:" . FoundY)
	return ErrorLevel = g_ErrorType_Success() && FoundX != "" && FoundY != ""
}

;// 이미지 서치해서 있는지 알려주는 함수
;// imageSearchStartPoint := [0, 0]
IsImageSearchV2(imageName, searchStart_x := 0, searchStart_y := 0)
{
	GuiControl,,B,IsImageSearch imageName(%imageName%) %searchStart_x%,*Trans091A36 *40 %searchStart_y%
	CoordMode, pixel, screen
	ImageSearch, FoundX, FoundY, searchStart_x, searchStart_y, A_ScreenWidth, A_ScreenHeight,*Trans091A36 *40 %A_ScriptDir%\Image\%imageName%.png
	Debug(imageName . " ErrorLevel:" . ErrorLevel . " FoundX:" . FoundX . " FoundY:" . FoundY)
	return ErrorLevel = g_ErrorType_Success() && FoundX != "" && FoundY != ""
}

;// 클립보드에서_중복되는_것중에_최대_값_추출
GetClipboardOverlapMaxValue(clipboardData)
{
	maxValue := 0

	resultArray := [] ; 값을 저장할 배열

	pos := 1
	while pos := RegExMatch(clipboardData, "/w_(\d+)(?=\/)/", match, pos) {
		resultArray.push(match1) ; 매치된 값을 배열에 추가
		pos += StrLen(match) ; 다음 검색 위치 업데이트
	}

	uniqueElements := {} ; 중복 인지 여부 요소를 담을 객체 생성

	; 중복되지 않는 요소를 객체에 기록
	for _, value in resultArray
		uniqueElements[value] := uniqueElements.HasKey(value) ? 2 : 1 ; 중복된 요소는 2, 그 외는 1로 설정

	; 중복되지 않는 요소 제외하고 배열 재구성
	overlappedValueArray := []
	for _, value in resultArray
	{
		if (uniqueElements[value] = 2) ; 중복된 요소만 선택하여 새 배열에 추가
			overlappedValueArray.push(value)
	}

	;// 배열에서 최대값 찾기
	for _, value in resultArray
	{
		if (value > maxValue)
			maxValue := value
	}

	return maxValue
}

;// 16진수 -> 10진수로 변경
HEX2DEC(rgb)
{
	Debug("HEX2DEC rgb : " . rgb)
	formattedNumber := Format("{:d}", rgb) ; 숫자를 10진수로 형식화 로 나중에 테스트 하기기
	Debug("HEX2DEC Format formattedNumber : " . formattedNumber)
	formattedNumber := "0x" + formattedNumber
	formattedNumber += 0
	Debug("HEX2DEC Format formattedNumber end : " . formattedNumber)

	SetFormat, IntegerFast, d ; 10진수로 정수 형식으로 설정
	decimalCode := "0x" + rgb
	decimalCode += 0
	Debug("HEX2DEC SetFormat decimalCode : " . decimalCode)
	return decimalCode
}

;// 10진수 -> 16진수로 변경
DEC2HEX(rgb)
{
	return Format("{:X}", rgb)
}

PixelGetCoIor(x, y)
{
	coordmode, pixel, screen
	PixelGetColor, OutputVar, %x%, %y%, RGB
	return OutputVar
}

;// 영역 내 컬러가 있는지 체크
;// color = 0x000000
IsCoIorByPosionRange(color, findPosionType, range)
{
	CoordMode, Mouse, Screen
	MouseGetPos, x, y
	count1 := 1
	while(range >= count1)
	{
		count2 := 1
		while(range >= count2)
		{
			positon := GetPosion(findPosionType, x , y, count1, count2)
			Debug("x : " . positon.x . " y : " . positon.y)
			if(PixelGetCoIor(positon.x, positon.y) = color)
			{
				return true
			}
			count2++
		}
		count1++
	}

	return false
}

;// 방향에 따라서 x, y 좌표 구하는 함수
GetPosion(findPosionType, x, y, count1, count2)
{
	Switch findPosionType
	{
	Case g_FindPosionType_upR():
		return { x: x + count1, y: y - count2}
	Case g_FindPosionType_downR():
		return { x: x + count1, y: y + count2}
	Case g_FindPosionType_upL():
		return { x: x - count1, y: y - count2}
	Case g_FindPosionType_downL():
		return { x: x - count1, y: y + count2}
	}
	return { x: x, y: y}
}

;// 배열에 중복값 제거
RemoveDuplicatesFromArray(arr)
{
	uniqueArray := []
	for _, value in arr {
		if !IsValueInArray(value, uniqueArray) {
			uniqueArray.Push(value)
		}
	}
	return uniqueArray
}

; 값이 배열에 있는지 확인하는 함수
IsValueInArray(value, arr)
{
	for _, val in arr
	{
		if(val = value)
		{
			return true
		}
	}
	return false ; 함수 내에서 사용되는 return 문
}

Debug(value)
{
	debugMessage =
	(

=>
	)

	OutputDebug, %debugMessage%
	nowTime := GetFormattedCurrentDateTime()
	OutputDebug, %nowTime% %value%
	ShowPopup(value, 1)
}

; 팝업을 보여주는 함수
ShowPopup(text, duration) {
    ToolTip, %text%, 0, A_ScreenHeight - 10  ; 화면의 중앙 하단에 팝업 표시
	durationValue := duration * 1000
    SetTimer, HidePopup, % -durationValue  ; duration 시간이 지난 후에 HidePopup 함수 호출
    return
}

; 팝업을 숨기는 함수
HidePopup() {
    ToolTip
}

; 크롬 현재 탭 닫기
NowChromeTabExit(){
	Send, ^w ; Ctrl + W로 현재 탭 닫기
}

; Clipboard에 복사 됨
ChromeTranslateGoogle(textToTranslate)
{
	; 번역할 텍스트 복사
	Clipboard := textToTranslate
	; 웹 브라우저 열기 및 Google 번역 페이지로 이동
	Run, chrome.exe https://translate.google.com/
	;// Google Chrome이 실행되고 그 창이 활성화될 때까지 대기
	WinWaitActive, ahk_class Chrome_WidgetWin_1
	SleepTime(1)
	WhileFoundImage("구글 번역\구글 번역 화면 로딩 끝")
	Send, ^v ; 번역한 텍스트 붙여넣기
	ClickAtWhileFoundImage("구글 번역\번역 된 것 복사 버튼", 10, 10)
	NowChromeTabExit()
	SleepTime(0.5)
	return Clipboard
}

;// 크롬에서 특정 탭으로 이동하는 기능(제목 포함 하면 멈춤)
WhileFindChromeTab(tabTitle)
{
	while (true) ; 무한 루프
	{
		SetTitleMatchMode, 2 ; 탭 제목을 기준으로 매칭
		WinGetTitle, nowTabTitle, A
		if(InStr(nowTabTitle, tabTitle) > 0)
		{
			SleepTime(1)
			return
		}
		; Ctrl + Tab 조합을 사용하여 다음 탭으로 이동
		Send, ^{Tab}
		SleepTime(1)
	}
}

;// 통관 고유부호 검증
;// name := "나철환"
;// ecm := "P180002538686"
;// phoneNumber := "01091700607"
IsCheckPcc(name, ecm, phoneNumber)
{
	WhileFindChromeTab("통관고유부호 검증") ;// https://gsiexpress.com/pcc_chk.php
	ClickAtWhileFoundImage("통관고유부호\통관 고유부호 검증", 0, 150)
	Clipboard := name . "/" . ecm . "/" . phoneNumber
	Send, ^v ; 번역한 텍스트 붙여넣기
	ClickAtWhileFoundImage("통관고유부호\통관부호 검증 확인")
	SleepTime(1)
	return IsImageSearchV2("통관고유부호\검증 결과 정상")
}

;// 크롬 창에서 주소창으로 이동
GoToTheAddressWindow()
{
	SendInput, !d ; Alt + D 키 전송
}

GetYesshipAddress()
{
	FormatTime, currentDate,, yyyyMMdd ; 현재 날짜를 년월일 형식(yyyyMMdd)으로 얻기

	; 현재 날짜에서 2달을 빼고 2달 전의 날짜 구하기
	currentYear := SubStr(currentDate, 1, 4)
	currentMonth := SubStr(currentDate, 5, 2)
	currentDay := SubStr(currentDate, 7, 2)

	; 2달 전의 월 계산
	if (currentMonth <= 2) {
		currentYear := currentYear - 1
		previousMonth := currentMonth + 10 ; 12월인 경우 10을 더함
	} else {
		previousMonth := currentMonth - 2
	}

	previousDate := Format("{:04d}{:02d}{:02d}", currentYear, previousMonth, currentDay)

	;// 예스쉽 입고완료에서 검색옵션을 주문번호 조회
	GoToTheAddressWindow()
	return "https://www.yesship.kr:8440/page/deliveryorder_list.asp?txtstartdate=" . previousDate . "&txtenddate=" . currentDate . "&gubun2="
}

;// 예스쉽 입고완료에서 검색옵션을 주문번호 조회
IsFindYesshipByOrderNumber(gubun, orderNumber)
{
	GoToTheAddressWindow()
	Clipboard := GetYesshipAddress() . gubun . "&gubun=pnumber&keyword=" . orderNumber . "&page=1"
	Send, ^v ; 번역한 텍스트 붙여넣기
	Send, {Enter}
	SleepTime(3)
	return !IsImageSearch("예스쉽\데이터가 없습니다")
}

;// 통관(세관) 시작 여부 체크
IsStartCustomsClearance(hblNo)
{
	oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	oHTTP.Open("GET", "https://unipass.customs.go.kr:38010/ext/rest/cargCsclPrgsInfoQry/retrieveCargCsclPrgsInfo?crkyCn=h280n293k172p003i070b030e0&hblNo=" . hblNo . "&blYy=2023", false)
	oHTTP.Send()
	responseBody := oHTTP.ResponseText
	return InStr(responseBody, "<hblNo>" . hblNo . "</hblNo>")
}

;// CJ 운송장 받았는지 여부 체크 만드는 중
IsStartCJ(hblNo)
{
	GoToTheAddressWindow()
	Clipboard := "https://trace.cjlogistics.com/next/tracking.html?wblNo=" . hblNo
	Send, ^v ; 번역한 텍스트 붙여넣기
	Send, {Enter}
	SleepTime(3)
	return !IsImageSearch("유효하지 않는 운송장번호 입니다")

	; oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	; oHTTP.Open("GET", "https://trace.cjlogistics.com/next/tracking.html?wblNo=433032447342", false)
	; oHTTP.Send()
	; responseBody := oHTTP.ResponseText
}

;// 배열을 , 구분해서 string로 반환
JoinArrayToString(array)
{
	outputString := ""
	Loop, % array.MaxIndex()
	{
		outputString .= (A_Index > 1 ? "," : "") . array[A_Index]
	}
	return outputString
}

;// 입력 예 -"aaa(205,260,270),bbb(260,261)"
StringToDoubleArray(input)
{
	input := StrReplace(input, " ", "_")
	result := []
	pos := 1
	while RegExMatch(input, "(\w+)\(([^)]+)\)", match, pos)
	{
		key := match1
		values := StrSplit(match2, ",")

		keyValue := []
		keyValue.push(key)
		keyValue.push(values)

		result.push(keyValue)

		input := SubStr(input, pos + StrLen(match))
	}
	return result
}

DoubleArrayToString(array)
{
	formattedString := ""
	Loop % array.MaxIndex()
	{
		key := array[A_Index][1]
		values := array[A_Index][2]

		valuesString := ""
		Loop % values.MaxIndex()
		{
			valuesString .= values[A_Index]
			if (A_Index < values.MaxIndex())
				valuesString .= ","
		}

		formattedString .= key . "(" . valuesString . ")"
		if (A_Index < array.MaxIndex())
			formattedString .= ","
	}
	return formattedString
}

SleepTime(second)
{
	Sleep, g_OneSecond() * second
}

;// 현재날짜및시간반환
GetFormattedCurrentDateTime()
{
	FormatTime, 현재날짜시간, , yyyy/MM/dd/HH:mm:ss
	return 현재날짜시간
}

;// 옵션 엑셀 세팅
SetExcelOptionByString(stringDoubleArray)
{
	result := StringToDoubleArray(stringDoubleArray)
	SetExcelOption(result, false)
}

;// 옵션 엑셀 세팅
SetExcelOption(doubleArray, customsDuty)
{
	; ;// 과세 부가 되는 기준을 정해서 옵션에 추가 할지 여부 정하기
	; ;// 31만원 이상이면 무조건 과세 부가
	; ;// 26 ~ 30만원 사이는 과세가 부가 될 수도 있습니다. 정확한 것을 판매자에게 문의 주세요 라고 옵션에 적기
	; ;// 25만원 이하는 무조건 과세 아님

	;// 사이즈 것을 옵션에 넣을까 고민 중
	;// 사이즈 신중하게 결정해주세요(사이즈로 교환 및 환불 불가 합니다.)

	xlFile := g_DefaultPath() . "\엑셀\OptionCombinationTemplate" . (customsDuty != 0 ? "_" : "2_") . ".xlsx"
	xl := ComObjCreate("Excel.Application") ; 엑셀 파일 열기
	xl.Visible := false ; Excel을 보이게 설정 할지 여부
	workbook := xl.Workbooks.Open(xlFile, 0, false, , "") ; 비밀번호로 보호된 파일 열기
	xl.DisplayAlerts := false
	worksheet := workbook.Sheets(1)
	rowCount := worksheet.UsedRange.Rows.Count
	allCount := 0
	for index, item in doubleArray {
		for index2, item2 in item[2] {
			allCount++
			;Debug("(" . index . "," . index2 . ")" . " : " . item[1] . " - " . item2)
			worksheet.Cells((allCount + 1) , 1).Value := item[1]
			worksheet.Cells((allCount + 1) , 2).Value := item2
			if(customsDuty != 0)
			{
				;// 25자 이내로 적어야 됨
				worksheet.Cells((allCount + 1) , 3).Value := "관부가세(23%) 수취인 부담(통관시 납부))"
			}
			worksheet.Cells((allCount + 1) , customsDuty != 0 ? 4 : 3).Value := 0
			worksheet.Cells((allCount + 1) , customsDuty != 0 ? 5 : 4).Value := 300
			worksheet.Cells((allCount + 1) , customsDuty != 0 ? 6 : 5).Value := index . index2
			worksheet.Cells((allCount + 1) , customsDuty != 0 ? 7 : 6).Value := "Y"
		}
	}

	; 이전 것 초과 된 행 삭제
	Loop % (rowCount - (allCount + 1)) {
		worksheet.Rows(allCount + 2).Delete()
	}

	; true를 전달하여 저장 여부 설정
	workbook.Close(true)
	xl.Quit()
}

CopyToClipboardAndGet()
{
	Send, ^c
	SleepTime(0.5)
	return Clipboard
}

IsArray(array, checkValue)
{
	Loop % array.MaxIndex()
	{
		if (array[A_Index] = checkValue)
		{
			return true
		}
	}
	return false
}

;// Matches
{
	GetRegExMatcheList(value, regexPattern, startPos := 1)
	{
		matcheList := []

		pos := startPos
		while (pos := RegExMatch(value, regexPattern, match, pos))
		{
			matcheList.Push(match)
			pos += StrLen(match)
		}

		Loop, % matcheList.Length()
		{
			Debug("matcheList[" . A_Index . "] : " . matcheList[A_Index])
		}

		return matcheList
	}

	GetRegExMatche1List(value, regexPattern, startPos := 1)
	{
		matche1List := []

		pos := startPos
		while (pos := RegExMatch(value, regexPattern, match, pos))
		{
			matche1List.Push(match1)
			pos += StrLen(match)
		}

		Loop, % matche1List.Length()
		{
			Debug("matche1List[" . A_Index . "] : " . matche1List[A_Index])
		}

		return matche1List
	}
}

DownloadImageUrl(url, saveName)
{
	savePath := g_DefaultPath() . "\DownloadImage\" . saveName . ".png" ; 저장할 파일의 경로 및 사용자가 원하는 파일 이름

	; 이미지 다운로드
	UrlDownloadToFile, %url%, %savePath%

	if (ErrorLevel == 0)
	{
		Debug("이미지 다운로드 url : " . url)
	}
	else
	{
		Debug("이미지 다운로드 실패 ErrorLevel : " . ErrorLevel)
	}
	return ErrorLevel == 0
}

FolderToDelete(folderPath)
{
	; 폴더 안의 모든 파일 삭제
	Loop, Files, %folderPath%\*.*
	{
		FileDelete, %A_LoopFileFullPath%
	}
	Debug("폴더 안의 모든 파일 삭제  folderPath : " . folderPath)
}

TelegramSend(Message) { 
	Debug(Message)
    ChatID := TelegramSend_ChatID()
    Token := TelegramSend_Token()
    Param := "chat_id=" ChatID "&text=" Message
    URL := "https://api.telegram.org/bot" Token "/sendmessage?" 
    a := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    a.Open("POST", URL)
    a.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
    a.Send(Param)
}

GetKorMony(mony, exchangeRate)
{
	;// 물건 원가
	korCostPrice := mony * exchangeRate
	if(korCostPrice = 0 && mony != 0)
	{
		Debug("******** Error  ---- GetKorMony() ==== 0   mony : " . mony . "    exchangeRate : " . exchangeRate . " ----  Error  ")
		TelegramSend("******** Error  ---- GetKorMony() ==== 0   mony : " . mony . "    exchangeRate : " . exchangeRate . " ----  Error  ")
		return 0
	}
	else
	{
		;// 수익율
		marginPrice := korCostPrice * g_MarginRate()
		;// 배송비
		courierPrice := g_CourierPrice()
		;// -3은 천원 단위로 반올림 하기 위함
		outValue := Round(korCostPrice + marginPrice + courierPrice, -3)
		if(outValue = 0 && mony != 0)
		{
			Debug("******** Error  ---- GetKorMony() ==== " . outValue . "   mony : " . mony . "    exchangeRate : " . exchangeRate . " ----  Error  ")
			TelegramSend("******** Error  ---- GetKorMony() ==== " . outValue . "   mony : " . mony . "    exchangeRate : " . exchangeRate . " ----  Error  ")
		}
		return outValue
	}
}

GetUggKorSize(usSize)
{
	Switch usSize {
	Case 4: 	return 210
	Case 5: 	return 220
	Case 5.5: 	return 225
	Case 6: 	return 230
	Case 6.5:	return 235
	Case 7:		return 240
	Case 7.5:	return 245
	Case 8:		return 250
	Case 8.5:	return 255
	Case 9:		return 260
	Case 9.5:	return 265
	Case 10:	return 270
	Case 10.5:	return 275
	Case 11:	return 280
	Case 11.5:	return 285
	Case 12:	return 290
	Case 13:	return 300
	Case 14:	return 310
	}
	return 0
}

GetMytheresaKorSize(size)
{
	Switch size {
	Case "EU 33": 	return "210"
	Case "EU 33.5": return "215-"
	Case "EU 34": 	return "215+"
	Case "EU 34.5":	return "220"
	Case "EU 35":	return "225-"
	Case "EU 35.5":	return "225+"
	Case "EU 36":	return "230"
	Case "EU 36.5":	return "235-"
	Case "EU 37":	return "235+"
	Case "EU 37.5":	return "240"
	Case "EU 38":	return "245-"
	Case "EU 38.5":	return "245+"
	Case "EU 39":	return "250"
	Case "EU 39.5":	return "255-"
	Case "EU 40":	return "255+"
	Case "EU 40.5":	return "260"
	Case "EU 41":	return "265-"
	Case "EU 41.5":	return "265+"
	Case "EU 42":	return "270"
	Case "EU 42.5":	return "275-"
	Case "EU 43":	return "275+"
	Case "EU 43.5":	return "280"
	Case "EU 44":	return "285-"
	Case "EU 44.5":	return "285+"
	}
	return ""
}
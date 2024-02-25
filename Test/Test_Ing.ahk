#Include EnvData.ahk
#Include GlobalData.ahk
#Include Util.ahk
#Include ProductRegistration.ahk

;// 나중에 <img src="{LINK}" />로 찾는 것 다시 시도해 보기
;// 실해 이유 source 받는 값에서 .png 정보에 가장 큰 이미지 정보가 없어서 (화면에 보이는 모든 url를 다 읽어서 해야 되는 것인 것 같음)
CrawlingTest()
{
	wh := ComObjCreate("WinHTTP.WinHTTPRequest.5.1")
	wh.Open("GET", "https://www.ugg.com/women-snow-boots/adirondack-boot-iii/1143530.html")
	;// 406 에러 나지 않도록 Header 설정
	; Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36
	wh.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
	wh.Send()
	wh.WaitForResponse()
	source := wh.ResponseText() 
	MsgBox, %source%

	values := []
	; 정규식 패턴 정의
	pattern := "(\w+):\s*([^,]+)"
	pos := 1  ; 문자열 내 위치 초기화
	; while 루프를 사용하여 모든 값을 가져옴
	while RegExMatch(source, pattern, output, pos) {
		;key := output1  ; 키(항목 이름)
		value := output2  ; 값(항목 값)

		values.push(value)
		;MsgBox, %key%: %value%  ; 각 항목의 키와 값을 메시지 박스로 출력

		pos := output.Pos + StrLen(output)  ; 다음 항목을 찾기 위해 위치 갱신
	}

	MsgBox, 끝났음
}

; 번역 함수 정의
;// 예) translatedText := DeepLTranslate(textToTranslate, "EN")  ; 번역할 언어 코드 입력
DeepLTranslate(textToTranslate, targetLanguage){
    apiKey := "7ae40e64-3bef-73d2-c357-e5565374bf69:fx"  ; 본인의 DeepL API 키로 변경
    apiUrl := "https://api.deepl.com/v2/translate"
    
    ; 번역할 텍스트와 대상 언어 설정
    body := "{""text"": """ . textToTranslate . """, ""target_lang"": """ . targetLanguage . """}"
    
    ; HTTP 요청 전송
    httpRequest := ComObjCreate("WinHttp.WinHttpRequest.5.1")
    httpRequest.Open("POST", apiUrl, false)
    httpRequest.SetRequestHeader("Content-Type", "application/json")
    httpRequest.SetRequestHeader("Authorization", "DeepL-Auth-Key " . apiKey)
    httpRequest.Send(body)
    
    ; JSON 응답 해석
    jsonResponse := httpRequest.ResponseText
    translatedText := RegExReplace(jsonResponse, ".*""text"":""(.*?)"".*", "$1")
    
    ; 번역된 텍스트 반환
    return translatedText
}

aaTest()
{
		;url := "https://unipass.customs.go.kr:38010/ext/rest/ecmQry/retrieveEcm?crkyCn=[k230h223c112s072z070e070h0]&ecm=P180002538686&rppnNm=나철환&rprsTelno=01091700607"
		; url := "https://unipass.customs.go.kr:38010/ext/rest/persEcmQry/retrievePersEcm?crkyCn=[인증키]
		; &persEcm=P160000000207&pltxNm=나철환&cralTelno=01091700607"
	
		; oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
		; ;oHTTP.Open("POST", url, false)
		; oHTTP.Open("GET", url, false)
		; oHTTP.Send()
		; responseBody := oHTTP.ResponseText
}

SendNateOnMessage(phonenNmber, message)
{
	WhileFindChromeTab("네이트온")
	ClickAtWhileFoundImage("네이트온\네이트온 문자 새로쓰기", 0, 100)
	Clipboard := message
	Send, ^v  ; 번역한 텍스트 붙여넣기

	ClickAtWhileFoundImage("네이트온\네이트온 문자 1번", 150, 20)
	Clipboard := phonenNmber
	Send, ^v  ; 번역한 텍스트 붙여넣기

	ClickAtWhileFoundImage("네이트온\네이트온 문자 보내기 버튼", 10, 10)
	ClickAtWhileFoundImage("네이트온\네이트온 문자 확인 버튼", 10, 10)
	ClickAtWhileFoundImage("네이트온\네이트온 문자 확인 버튼", 10, 10)
	ClickAtWhileFoundImage("네이트온\네이트온 문자 번호 삭제 버튼", 10, 10)
}


로그인테스트해본것()
{
	wh := ComObjCreate("WinHTTP.WinHTTPRequest.5.1")
	wh.Open("POST", "https://www.yesship.kr:8440/login/login_prc.asp")
	;// 406 에러 나지 않도록 Header 설정
	wh.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	; Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36
	
	wh.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
	; 로그인 요청을 위한 데이터 준비 (POST 요청)
	postData := "txtid=" . ID() . "&txtpwd=" . Password()
	wh.Send(postData)
	; wh.Send()
	wh.WaitForResponse()
	source := wh.ResponseText

	GoToTheAddressWindow()
	Clipboard := GetYesshipAddress() . "OC" . "&gubun=pnumber&keyword=" . "2311261670" . "&page=1"

	wh.Open("POST", "https://www.yesship.kr:8440/page/deliveryorder_detail.asp?pnumber=2311261670")
	;// 406 에러 나지 않도록 Header 설정
	; Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36
	; wh.SetRequestHeader("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36")
	wh.Send()
	wh.WaitForResponse()
	source1 := wh.ResponseText
	source2 := wh.GetAllResponseHeaders()

	;// source1에 일부만 받아 오는 이유는 동적으로 되어져 있어서 나중에 더 테스트 해 보기
}

엑셀_읽기_테스트()
{
	; 열고자 하는 엑셀 파일 경로 설정
	xlFile := g_DefaultPath() . "\엑셀\OptionCombinationTemplate (11).xlsx"
	password := "skcjf158!" ; 여기에 Excel 파일의 비밀번호를 입력하세요
	xl := ComObjCreate("Excel.Application") ; 엑셀 파일 열기
	xl.Visible := false ; Excel을 보이게 설정 할지 여부
	workbook := xl.Workbooks.Open(xlFile, 0, false, , password) ; 비밀번호로 보호된 파일 열기
	worksheet := workbook.Sheets(1)

	; 데이터 읽기
	rowCount := worksheet.UsedRange.Rows.Count
	colCount := worksheet.UsedRange.Columns.Count

	A_Index1 := 1
	Loop, % rowCount {
		A_Index2 := 1
		Loop, % colCount {
			cellValue := worksheet.Cells(A_Index1, A_Index2).Value
			if (cellValue != "") {
				Debug("Row: " A_Index1 ", Col: " A_Index2 ", Value: " cellValue)
			} else {
				Debug("Row: " A_Index1 ", Col: " A_Index2 ", Value: Empty Cell")
			}
			A_Index2++
		}
		A_Index1++
	}

	; 엑셀 파일 닫기
	workbook.Close()
	xl.Quit()
}

구글_시트_특정_계정_접근으로_하는_것_엑셀_읽기_테스트()
{
	; ;// 파이썬이나 nodejs로 해서 결과 받는 걸로 해야 됨
	; oHTTP.SetRequestHeader("X-API-KEY", apiKey) ; API 키를 요청 헤더에 추가
	; serviceJson := {"type": "service_account","project_id": "magusaa-5","private_key_id": "e44ad6d92d3fa13531facf450f73024c648a2103","private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC/xvDAQBQybHRb\nO5RKNtqfo/qs85EEfuqZy0sca0QsvnsIo8E7X0ZrV5SdCHELHlYqqn/CO8NKCHNJ\nhIQ0ZZFMHjolT9aU7DKSPMUNg4LbJ1gZKBQ2IjC14gMqG07zY7xZowZcMc6L5np8\nGk80rv3hxyofAr/oaQ7RIqIUGs+qKZdW7k5p8X1OrI0X4rwZdEkup1ULmE5nmhVt\nzMJ3dIHr7F9ITm7l4dVvhXxnAwp0+L0Z/YcdSgmXIZwXMRzmaIX2LGYe53BmvkvU\na+Her50BSIo75B39c3yc7QiGpyQoQ71gdP59yitekOLnP+HF/1MzG9/hytazAgjP\n3Gt38TlbAgMBAAECggEAHPYFG3NR2X+JXtGzhjWcdGlGDOJHbx9ffFQ4BpWoWP13\nBQn9v2KW9tTmC4Tf1WcCajUPUBzIVCDMkDij4mnING+IJmRVKm41AOKOe4j1tPTP\nGBV2X+pv4re79JrlJFpWck7tQfT/lR78Nkue1HzsuDDkioRWyNk8tJZ/VXvrCm4z\n5PpCE3jx3gqvaZJAV+S7VWtBJ+fFdsgmALRpk/OdAkDAN7EVzEaf5LrMKnfmnZlk\ns08yWg2LpB3qnBICjaKBE4lHIXtpwDQQlyNcPMcZXs7maaQVGMBix7g5S2exBOCo\nT/QKybrA5R0jXqTFqxSx7dVLrb+WzQb6PRUGL5F1CQKBgQD/UwQlC97uZs+on0xf\nVtkDKzR3rqSMRdl4fdUgAvfV4kuRiQCGnB52Pro9iWp4qRjdL501ucCXVpoNyQJA\nBCC+oggKkL7jzBKjjCdluERosgBlVK8go8ivnE9iSEF5ElwP7NRFA+RiZm2t0ZNM\ncqYlGREGItoGRY0PtobhEhfyzQKBgQDASN7h58+Hj1wKQVhqGn+rCzgiD2gR8Pp0\nuAiLPlFYxuzC3zc1uQ4y8ymICOumegHNvTdUEo1BrwQ2l7ssC1rS49IjCxriv9HK\nMBwLyGK3YJRXRuZjbGhp7mffawr+2ilzcX9jruJG4/X69+yRmsgy0lW4ewIupJwg\nwCpwsnlsxwKBgQD29bdGlgrVkYA+W4almP1jAUFImhXy0AUfdKbWxcguiyoI5Pkr\nOoqEWPwPVYE0oGq6Vrm7I6ZTO6Lavph8jwGVImigv4zEDbnhk0jwLKGOms2jNZwG\n+CS/J3PpXnZlwwplJO/UqYUYYHap79KH2UU3EN3Uj5VPB6r/jc88mCGt6QKBgF0Q\n6A+fCyspj/rGtexk9vXqcDjMDCri39YuXLRIbUbywRVwxGAUOXMfjjJxXt0soELc\nGjNu5z+rXfauacFfnY4FBmg/r7uf7AJYVrq9OkpXTHURs+DbT441/cB9Js1C+l0N\nygKNWqfFHgFijfXLXKp8c1De+KdqtMaFPAVf3LVxAoGANgEGPimBHzYZTj+0abWf\nJUcC+xh1l+aDSVGkhx3W5Zpj0rdSnJdfP+7CYhz3NOkVk/3csLfeb+ajHm3xsUBH\nk2f5Np5kaz05xDF4IWeHtLzswa4t12/uvUy0lo+CHeh4B3Mbv8iLJj5N4KQwvxxg\n4HLgw1/x7Uv1LbzU4MgBP9s=\n-----END PRIVATE KEY-----\n","client_email": "maguass-5@magusaa-5.iam.gserviceaccount.com","client_id": "104864907741831138717","auth_uri": "https://accounts.google.com/o/oauth2/auth","token_uri": "https://oauth2.googleapis.com/token","auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs","client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/maguass-5%40magusaa-5.iam.gserviceaccount.com","universe_domain": "googleapis.com"}

	; ; Google Sheets API 요청에 필요한 정보 설정
	; serviceAccountEmail := "maguass-5@magusaa-5.iam.gserviceaccount.com" ; 서비스 계정 이메일
	; serviceAccountKeyFile := g_DefaultPath() . "\Test\magusaa-5-e44ad6d92d3f.json" ; 서비스 계정 JSON 키 파일 경로
	; spreadsheetId := "1eoVRmnRYw-y0Rayp32zcWDPBPRcwolEveWXItTheCSk"
	; range := "시트1!G30" ; 예: "Sheet1!A1"


	; objHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	; url := "https://www.googleapis.com/oauth2/v4/token"
	; client_id := "Y800061949795-ef3q3g3v2abna89rburue93o0o0ga84k.apps.googleusercontent.com" ; 여기에 실제 client ID 값 입력
	; client_secret := "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" ; 여기에 실제 client secret 값 입력

	; objHTTP.Open("POST", url, false)
	; objHTTP.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded;charset=UTF-8")

	; payload := "client_id=" . client_id . "&client_secret=" . client_secret . "&grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer"
	; objHTTP.Send(payload)

	; response := objHTTP.ResponseText

	; Debug(objHTTP.ResponseText)
	
	; if (objHTTP.Status = 200) {
	; 	Debug(objHTTP.ResponseText)
	; } else {
	; 	Debug("에러 발생: " objHTTP.Status " - " objHTTP.StatusText "`n응답 내용: " objHTTP.ResponseText)
	; }

	; Access Token 얻기
	; access_token := RegExReplace(response, ".*access_token""\s*:\s*""([^""]+).*", "$1")

	; Google Sheets API 호출
	; objHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	; objHTTP.Open("GET", "https://sheets.googleapis.com/v4/spreadsheets/" spreadsheetId "/values/" range, false)
	; objHTTP.SetRequestHeader("Content-Type", g_DefaultPath() . "\Test\magusaa-5-e44ad6d92d3f.json")
	; ; objHTTP.SetRequestHeader("Authorization", "Bearer " access_token)
	; objHTTP.SetRequestHeader("X-API-KEY", apiKey) ; API 키를 요청 헤더에 추가
	; objHTTP.Send()

	; if (objHTTP.Status = 200) {
	; 	Debug(objHTTP.ResponseText)
	; } else {
	; 	Debug("에러 발생: " objHTTP.Status " - " objHTTP.StatusText "`n응답 내용: " objHTTP.ResponseText)
	; }
}

구글_시트_값_넣는_것_테스트_중이였음()
{
	; range := "시트1!G30"
	; url := "https://sheets.googleapis.com/v4/spreadsheets/" . g_SheetId_magussa5_Root() . "/values/" . range . "?key=" . g_SheetApiKey()
	; oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	; oHTTP.Open("GET", url, false)
	; oHTTP.Send()

	; RunWait, python g_DefaultPath() . \Test\Test1.py,,Hide

	; if(ErrorLevel = 0)
	; {
	; 	Debug("외부 프로그램 실행 성공")
	; }
	; else
	; {
	; 	Debug("외부 프로그램 실행 실패  : " . ErrorLevel)
	; }

	; ; aaa := {"token": "ya29.a0AfB_byCw8qxvv8801XmqS00pA59qvePUZ8qBfKDj8YBVQ9WvGOf2Go7ncFwOEL4m5mbaoXDNjCWPS3cD4zTwiAsHzpi2q4Ic9isqGAEJ8u40JhYMRnnwzia9_x4xDZMDHw7VOfpprNGYiLKBr7CkPvl_cfFD85503FxBaCgYKATYSARISFQHGX2MiHEdZIyblqc2apiTffCZlaA0171", "refresh_token": "1//0edOBwOZXf5ZaCgYIARAAGA4SNwF-L9Ira4Ru2yBTfQUxqgskAqDxqYkGfSGqBwVlHcHFSutf1Xdma1psW2bc_S8ionkqradaUS0", "token_uri": "https://oauth2.googleapis.com/token", "client_id": "800061949795-ef3q3g3v2abna89rburue93o0o0ga84k.apps.googleusercontent.com", "client_secret": "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", "scopes": ["https://www.googleapis.com/auth/spreadsheets"], "universe_domain": "googleapis.com", "expiry": "2023-12-24T10:33:23.872564Z"}
	; aa := Clipboard

	; range := "시트1!H30"
	; new_value := "New Value" ; 업데이트할 새로운 값
	; TestText:="{""values"":[[""" . new_value . """]]}"
	; update_data := { "range": range, "values": [[new_value]] }
	; valueInputOption := "USER_ENTERED"
	; ;url := "https://sheets.googleapis.com/v4/spreadsheets/" . g_SheetId_magussa5_Root() . "/values/" . range . "?valueInputOption=" . valueInputOption . "&key=" . g_SheetApiKey()
	; url := "https://sheets.googleapis.com/v4/spreadsheets/" . g_SheetId_magussa5_Root() . "/values/" . range . "?valueInputOption=" . valueInputOption . "&access_token=" . g_SheetApiKey()
	; oHTTP := ComObjCreate("WinHttp.WinHttpRequest.5.1")
	; oHTTP.Open("PUT", url, false)
	; oHTTP.SetRequestHeader("Content-Type", "application/json")
	; ;oHTTP.Send(ObjBindMethod(update_data, "toJSON"))
	; oHTTP.send(TestText)
	; oHTTP.WaitForResponse()

	; if (oHTTP.Status = 200) {
	; 	Debug(oHTTP.ResponseText)
	; } else {
	; 	Debug("에러 발생: " oHTTP.Status " - " oHTTP.StatusText "`응답 내용: " oHTTP.ResponseText)
	; }
}

;// 이미지 한글 이름 반환(문제점 - 깊이 관련된 색으로 하니깐 같은 색이 나오는 경우가 있음)
GetKoreanColorName_before(colorHax)
{
	while(true)
	{
		url := "https://encycolorpedia.kr/" . colorHax
		Run, chrome.exe %url%
		WinWaitActive, ahk_class Chrome_WidgetWin_1
		SleepTime(1)

		WhileFoundImage("색 이름 찾기\Before\이미지 코드 색상 검색 로딩 끝")

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, -1000, uint, 0) ; -1000 틱 스크롤 다운
		SleepTime(1)
		if(!FoundImage("색 이름 찾기\Before\광고"))
		{
			ClickAtWhileFoundImage("색 이름 찾기\Before\깊이 관련된", -100, 35)

			if(!FoundImage("색 이름 찾기\\Before광고"))
			{
				Click down                   ; 마우스 왼쪽 버튼을 누른 상태로 클릭합니다.
				CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
				MouseGetPos, currentX, currentY ; 현재 마우스 위치 얻기
				MouseMove currentX + 300, currentY
				Click up       
				
				Send, ^c  ; 복사

				;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
				Send ^w

				return Clipboard
			}
		}
	}
}

;// 신규 등록 할 상품 url 배열로 추출
GetUGGNewProductURLs_Test()
{
	urlArray := []
	while (true)
	{
		xArray := [300, 750, 1200, 1600]
		yArray := [300, 600, 900]

		for index, y in yArray
		{
			for _, x in xArray
			{
				CoordMode, Mouse, Screen
				mousemove, %x%, %y%
				SleepTime(0.5)
				NowMouseClickRight()
				SleepTime(1)

				searchStart_x := x
				searchStart_y := index = 1 ? (y + 150): (y - 400)

				if(!ClickAtWhileFoundImage("마우스 오른쪽 누른 후 나오는 목록\링크 주소 복사", 5, 5, 1, 1, searchStart_x, searchStart_y))
				{
					SleepTime(0.5)
					;// 링크 주소 복사가 없는 곳은 클릭해서 마우스 오른쪽 누른 후 나오는 목록 없애기(링크 주소 없는 곳은 클릭해도 무방함)
					NowMouseClick()
					SleepTime(0.5)
					;// 링크 주소가 없으면 무조건 다음 줄로 가도록 처리
					Break
				}
				else
				{
					SleepTime(0.5)
					;// 링크 주소 복사 된 것을 배열에 넣기
					if(!IsArray(urlArray, Clipboard) && 0 != InStr(Clipboard, ".html"))
					{
						urlArray.Push(Clipboard)
					}
					else
					{ ;// 이미 배열에 존재 하면 다른 줄로 가도록 처리
						SleepTime(0.5)
						ScreenMouseMove(5, 600)
						SleepTime(0.5)
						NowMouseClick()
						SleepTime(0.5)
						Break
					}
				}
			}
		}

		SleepTime(0.5)
		ScreenMouseMove(5, 600)
		SleepTime(0.5)
		NowMouseClick()
		SleepTime(0.5)

		; 스크롤 시작 위치에서 아래로 이동하여 스크롤링
		DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, -1000, uint, 0) ; -1000 틱 스크롤 다운
		SleepTime(1)

		if(ClickAtWhileFoundImage("크롬\오른쪽 스트롤바가 제일 아래인 이미지", 0, 0, 2, 1, 1800, 900))
		{
			;// 상품 더 보기가 있는지 체크
			if(ClickAtWhileFoundImage("UGG\상품 리스트\상품 더 보기 버튼", 0, 0, 2))
			{
				SleepTime(5)
			}
			else
			{ ;// 상품이 더이상 없음
				Break
			}
		}
	}
	return urlArray
}

;// UGG_상품_컬러_Hex 정보
GetUggProductColorHexList()
{
	saveColorHexList := []
	ClickAtWhileFoundImage("UGG\Color_v3", 0, 0)
	;// 바닥 컬러 저장해서 바닥 컬러 색이 나오면 중지 하기
	pixelColor := DEC2HEX(GetNowMousePixelColor())
	NowMouseMove(18, 43)
	saveColor := DEC2HEX(GetNowMousePixelColor())
	while (pixelColor != saveColor) {
		saveColorHexList.Push(saveColor)
		SleepTime(1)
		NowMouseMove(38, 0)
		saveColor := DEC2HEX(GetNowMousePixelColor())
		Debug("saveColorHexList.Length : " . saveColorHexList.Length())
	}
	return saveColorHexList
}

;// 이미지 한글 이름 반환(문제점 - 깊이 관련된 색으로 하니깐 같은 색이 나오는 경우가 있음)
GetKoreanColorName(colorHax)
{
	while(true)
	{
		url := "https://www.htmlcsscolor.com/hex/" . colorHax
		Run, chrome.exe %url%
		WinWaitActive, ahk_class Chrome_WidgetWin_1
		SleepTime(1)

		WhileFoundImage("색 이름 찾기\이미지 코드 색상 검색 로딩 끝")

		SleepTime(1)

		ClickAtWhileFoundImage("색 이름 찾기\Donate", -300, 70)
		SleepTime(1)

		Click down ; 마우스 왼쪽 버튼을 누른 상태로 클릭합니다.
		CoordMode, Mouse, Screen ; 마우스 좌표 모드 설정 (화면 기준 좌표)
		MouseGetPos, currentX, currentY ; 현재 마우스 위치 얻기
		MouseMove currentX + 375, currentY
		Click up

		Send, ^c ; 복사

		;// Ctrl + W를 눌러 현재 Chrome 탭 닫기
		Send ^w

		SleepTime(0.5)

		return Clipboard
	}
}
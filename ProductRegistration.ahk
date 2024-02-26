#Include EnvData.ahk
#Include GlobalData.ahk
#Include Util.ahk

;// 상품 등록시 기본 세팅 들
ProductRegistrationDefaultSettings()
{
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\상품 주요정보", 10, 10)
	SleepTime(0.5)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\구매대행 라디오 버튼", 13, 13)
	SleepTime(0.5)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\국산", 10, 10)
	SleepTime(0.5)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\기타", 10, 10)
	SleepTime(0.5)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\에누리", 10, 10)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\다나와", 10, 10, 100, 10)
	SleepTime(0.5)
	WheelAndMoveAtWhileFoundImage("스마트 스토어\제일 위 상품등록", 0, 0, 10000)
}

;// 에디터_상품_이미지들_등록
SetEditorProductImageList(count)
{
	Clipboard := count . "번 색상"
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)
	Send, {Enter}
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\사진", 15, 15)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\내 사진", 0, 0)
	SleepTime(1)
	ClickAtWhileFoundImage("파일 탐색기\다운로드", 5, 5)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\이름", 20, 75)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\열기", 5, 5)
	SleepTime(0.5)
	DragAtFoundImage("파일 탐색기\열기 탐색기\이름", -20, 75, "파일 탐색기\열기 탐색기\열기", 5, 5)
	NowMouseClick()
}

;// 에디터_상단_이미지_등록
SetEditorTopHtmlImage()
{
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\왼쪽 정렬 상태", 10, 10)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\가운데 정렬", 10, 10)
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\HTML이미지", 15, 15)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\내용을 입력해주세요", 15, 15)
	SleepTime(0.5)
	Clipboard := "<img src=""https://nacharhan.github.io/photo/2.png""/>"
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\변환", 15, 15)
	SleepTime(1)
}

;// 에디터_하단_이미지_등록
SetEditorBottomHtmlImage()
{
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\HTML이미지", 15, 15)
	SleepTime(1)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\내용을 입력해주세요", 15, 15)
	SleepTime(0.5)
	Clipboard := "<img src=""https://nacharhan.github.io/photo/11.png""/>"
	SleepTime(0.5)
	Send, ^v
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\스마트 에디터\변환", 15, 15)
	SleepTime(1)
}

;// 다운로드에_있는_것_다_삭제
DeleteInTheDownload()
{
	ClickAtWhileFoundImage("파일 탐색기", 15, 15)
	SleepTime(1)
	ClickAtWhileFoundImage("파일 탐색기\내 피씨", 15, 15)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\다운로드", 5, 5)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\여러개 선택 이름", 10, 10)
	SleepTime(1)
	ClickAtWhileFoundImage("파일 탐색기\삭제 버튼", 10, 10, 2)
	SleepTime(1)
}

;// 이미지 등록(대표, 추가)
IamgeRegistration()
{
	;// 대표 이미지 등록
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\대표이미지", 250, 20)
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\내 사진", 0, 0)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\다운로드", 5, 5)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\이름", 20, 75)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\열기", 5, 5)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\이름", 20, 75)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\열기 탐색기\열기", 5, 5)
	SleepTime(0.5)

	;// 추가 이미지 등록
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\추가이미지", 250, 5)
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\내 사진", 0, 0)
	SleepTime(0.5)
	DragAtFoundImage("파일 탐색기\열기 탐색기\이름", -20, 105, "파일 탐색기\열기 탐색기\열기", 5, 5)
	SleepTime(0.5)
	NowMouseClick()
}

;// 이미지 등록(대표, 추가)
IamgeRegistration_v2()
{
	;// 대표 이미지 등록
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\대표이미지", 250, 20)
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\내 사진", 0, 0)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\DownloadImage", 5, 5)
	SleepTime(0.5)
	Send, {Tab}
	SleepTime(0.5)
	Send, {Right}
	SleepTime(0.5)
	Send, {Left}
	SleepTime(0.5)
	Send, {Enter}
	SleepTime(2)

	Send, {Esc}
	SleepTime(1)

	if(MoveAtWhileFoundImage("스마트 스토어\내 사진", 0, 0, 2))
	{
		;// 이미지 넣을 수 없는 것임(너무 큼)
		return false
	}

	;// 추가 이미지 등록
	ClickAtWhileFoundImage("스마트 스토어\상품 수정\추가이미지", 0, 0)
	SleepTime(1)
	DllCall("mouse_event", uint, 0x800, int, 0, int, 0, uint, -50, uint, 0) ; wheelMove 틱 스크롤 다운
	SleepTime(1)
	WheelAndClickAtWhileFoundImage("스마트 스토어\상품 수정\추가이미지", 250, 20, -500)
	SleepTime(0.5)
	ClickAtWhileFoundImage("스마트 스토어\내 사진", 0, 0)
	SleepTime(0.5)
	ClickAtWhileFoundImage("파일 탐색기\DownloadImage", 5, 5)
	SleepTime(0.5)
	Send, {Tab}
	SleepTime(0.5)
	Send, {Right}
	SleepTime(0.5)
	Send, {Left}
	SleepTime(0.5)
	Send, {Delete}
	SleepTime(0.5)
	Send, ^a
	SleepTime(0.5)
	Send, {Enter}

	return true
}

;// 발 사이즈 어떤 것이 있는지와 한국 사이즈로 변환
;// findSizeList : [{korSize:220, sizeImage:"us_5"}]
GetArrImageSearch(imageName, findSizeList)
{
	;// 이미지 서치 시작하는 곳을 찾아서 서치하는 영역 줄이도록 처리
	imageSearchStartPoint := [0, 0]
	GuiControl,,B,%imageName%
	coordmode, pixel, screen
	FindImage_Byref(imageName, error := it, x := it, y := it)
	if(error = g_ErrorType_Success() && x != "" && y != "")
	{
		imageSearchStartPoint := [x, y]
	}
	imageSearchStartPoint[1] -= 10
	imageSearchStartPoint[2] -= 10

	korSizeList := []
	count := 1
	while(findSizeList.Length() >= count)
	{
		if(IsImageSearch(findSizeList[count].sizeImage, imageSearchStartPoint[1], imageSearchStartPoint[2]))
		{
			korSizeList.Push(findSizeList[count].korSize)
			Debug("Add korSize :" . findSizeList[count].korSize)
		}else if(IsImageSearch(findSizeList[count].sizeImage . "_select", imageSearchStartPoint[1], imageSearchStartPoint[2]))
		{
			korSizeList.Push(findSizeList[count].korSize)
			Debug("Add korSize :" . findSizeList[count].korSize)
		}
		count += 1
	}

	return korSizeList
}
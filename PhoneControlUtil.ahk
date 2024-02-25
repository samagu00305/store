#Include EnvData.ahk
#Include Util.ahk

PhoneControlUtil_SendMessage(phoneNumber, message)
{
	ClickAtWhileFoundImage("폰 제어\전화", 20, 20)
	SleepTime(2)
	Loop, 20
	{
		Send, {Backspace}
		SleepTime(0.1)
	}
	PhoneControlUtil_InputWordByWord(phoneNumber, true)
	SleepTime(0.5)
	ClickAtWhileFoundImage("폰 제어\전화에서 메시지 보내는 메뉴", 10, 10)
	SleepTime(2)
	ClickAtWhileFoundImage("폰 제어\메시지 보내기", 10, 10)
	SleepTime(2)
	PhoneControlUtil_InputWordByWord(message)
	SleepTime(0.5)
	ClickAtWhileFoundImage("폰 제어\문자 전송 버튼", 20, 20)
	SleepTime(0.5)
	ClickAtWhileFoundImage("폰 제어\메인 화면으로 이동 버튼", 10, 10)
	SleepTime(0.5)
}


;// 문장을 각각 하나씨 차례대로 입력
;// 사용 할 곳 - 폰 화면에서 입력 할 때 사용
PhoneControlUtil_InputWordByWord(text, isRight := false)
{
	index := 1
	Loop {
		if (index > StrLen(text))
			break
		
		char := SubStr(text, index, 1)
		Send, %char%
		if(isRight)
		{
			Send, {Right}
		}
		index++
		SleepTime(0.05) ; 각 글자 입력 사이에 잠시 지연을 주기 위해 50밀리초 대기합니다.
	}
}

#Include EnvData.ahk
#Include Util.ahk

;// 엑셀 폴더 git 동기화
GitSyncExcelFolder()
{
	folderPath := g_DefaultPath() . "\엑셀"
	RunWait, git pull, %folderPath%
}

;// 특정 파일 업로그 
GitSyncPushExcelFile()
{
	filePath := g_DefaultPath() . "\엑셀\마구싸5_구매루트.xlsx"
	RunWait git add %filePath%
    RunWait git commit -m "구매루트 최신화"
    RunWait git push origin main
	TelegramSend("Git 에 마구싸5_구매루트.xlsx 최신화")
}
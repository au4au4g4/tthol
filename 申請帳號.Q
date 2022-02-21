[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=062dd7c2-b2af-495f-807f-eb515d1d4c25
Description=申請帳號
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
// 申請遊戲帳號
accounts = array("fgcjjhhhjb","gdbiiadhig","ibgggghdai","idhbiejfjj","ejgfdhaagc","iebaddgcfj","adfdebjacf","bjihaaddec","biccbbfgij","daaiihdjfe")
For Each account In accounts
	MoveTo 560, 408
	LeftClick 2
	SayString account
	MoveTo 561, 443
	LeftClick 2
	SayString "Ugj83dj4"
	MoveTo 573, 496
	LeftClick 1
	MoveTo 472, 613
	LeftClick 1
	Delay 1500
	
	MoveTo 571, 590
	LeftClick 1
	Delay 1500
	
	MoveTo 551, 374
	LeftClick 2
	SayString account	
	MoveTo 554, 425
	LeftClick 2
	SayString "gj83dj4"
	MoveTo 556, 479
	LeftClick 2
	SayString "gj83dj4"
	MoveTo 658, 551
	LeftClick 1
	MoveTo 569, 639
	LeftClick 1
	Delay 1500
	
	MoveTo 418, 426
	LeftClick 1
	Delay 1500
	TracePrint 1
Next

// 註冊
//accounts = array("daaiihdjfe")
//For Each account In accounts
//	MouseWheel -5
//	MoveTo 672, 93
//	LeftClick 1
//	Delay 300
//	MoveTo 511, 142
//	LeftClick 2
//	SayString account
//	MoveTo 471, 505
//	LeftClick 1
//	WaitKey
//	MoveTo 657, 888
//	LeftClick 1
//	Delay 2000
//	KeyDown 18, 1
//	KeyPress 37, 1
//	KeyUp 18, 1
//	Delay 1000
//Next

// 生成帳號
//ss = ""
//For i = 0 To 5
//	ss = ss & RandomString(10) & ","
//Next
//TracePrint ss
//
//Function RandomString( ByVal strLen )
//    Dim str, LETTERS, i
//    LETTERS = "abcdefghijklmnopqrstuvwxyz0123456789"
//    Randomize
//    For i = 1 To strLen
//        str = str & Mid( LETTERS, Int(strLen*Rnd+1), 1 )
//    Next
//    RandomString = str
//End Function

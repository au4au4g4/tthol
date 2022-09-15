[General]
SyntaxVersion=2
BeginHotkey=57
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=3781a671-1216-431f-be63-fdde4c1be28d
Description=開小方
Enable=1
AutoRun=1
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set dm = createobject("dm.dmsoft")

w = dm.getscreenwidth()
h = dm.getscreenheight() - 40
c = 5
r = 2

teams = t.addr("teams")

windowCnt = UBound(t.getAllHwnds())
For each team in teams
	team = t.addr(team)
	For Each id In team
		setting = t.addr(id)
		dm_ret = dm.UnBindWindow()
		dm.moveto 120, 1060
		dm.leftclick 
		While windowCnt >= UBound(t.getAllHwnds())	
			Delay 500
		Wend
		windowCnt = windowCnt + 1

		// 調整位置
		hwnd = dm.FindWindow("", "絕代方程式")
		dm_ret = dm.SetClientSize(hwnd, w / c - 2, h / r - 32)
		x = (setting(0) mod 5) * w / c - 7
		y = (setting(0) \ 5) * h / r
		dm.MoveWindow hwnd, x, y
		
		// 輸入帳密
		login = dm.FindWindow("", "登錄")
		edits = split(dm.EnumWindow(login, "", "Edit", 2 + 4), ",")
		dm_ret = dm.BindWindow(edits(0), "normal", "windows", "windows", 0)
		call del(15)
		dm.SendString edits(0), id
		dm.SendString edits(1), "gj83dj4"
		comboBoxs = split(dm.EnumWindow(login, "", "ComboBox", 2 + 4), ",")
		dm_ret = dm.BindWindow(login, "normal", "windows", "windows", 0)
		dm.SendString comboBoxs(0), "飛雁山莊(花)"
		dm.SendString comboBoxs(1), "1"
		Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+E4A4", 0, 0)
		
		// 破解
		t.init hwnd
		t.crackAll setting(1), setting(2), setting(3)
		t.login
	Next
	
	// 換頁
	KeyDown "Ctrl", 1
	KeyDown "Win", 1
	KeyPress "Right", 1
	KeyUp "Win", 1
	KeyUp "Ctrl", 1
Next
KeyDown "Ctrl", 1
KeyPress "2", 1
KeyUp "Ctrl", 1

Function del(cnt)
	Dim i
	For i = 0 To cnt
		dm.KeyPress 8
		dm.KeyPress 46
	Next
End Function

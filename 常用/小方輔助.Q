[General]
SyntaxVersion=2
BeginHotkey=97
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=30f7364a-18a2-47f5-9cae-b716b9751a8a
Description=小方輔助
Enable=1
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
Set dm = createobject("dm.dmsoft")

dw = 250 : dh = 130

While true
	key = WaitKey()
	hwnds = getPartHwnds()
	// 點擊
	If key = 112 Then 
		xy = getXY()
		For each hwnd in hwnds
			Call Bind(hwnd)
			Call lClick(xy)
		Next
	// 輸入文字
	ElseIf key = 113 Then
		For each hwnd in hwnds
			Call Bind(hwnd)
			Call del(10)
			Delay 100
			dm.SendPaste hwnd
		Next
	// 輸入按鍵
	ElseIf key = 114 Then
		key = WaitKey()
		For each hwnd in hwnds
			Call Bind(hwnd)
			dm.KeyPress key
		Next
	// 改地點
	ElseIf key = 115 Then
		hwnds = split(dm.EnumWindow(0, "跑動路徑選擇", "", 1 + 4 + 16), ",")
		GetCursorPos x, y
		dm_ret = dm.ScreenToClient(hwnds(0), x, y)
		For Each hwnd In hwnds
			btn = split(dm.EnumWindow(hwnd, "", "Button", 2 + 8 + 16), ",")
			Call bindLClick(btn(0))
			Call bind(hwnd)
			Call lClick(array(x, y))
			Call bindLClick(btn(1))
		Next
	// 調視窗大小
	ElseIf key = 116 Then
		hwnds = split(dm.EnumWindow(0, "Tab2", "SysTabControl32", 1 + 2 + 16), ",")
		For Each hwnd In hwnds
			dm.GetClientRect dm.GetWindow(hwnd, 7),x,y,x0,y
			dm.GetClientRect hwnd, x1, y, x, y
			d = (x0 - x1 > 250) * - 2  - 1
			
			hwnd = dm.GetWindow(hwnd, 4)
			dm_ret = dm.BindWindow(hwnd, "normal", "windows", "windows", 0)
			dm.moveto 0, 0
			dm.LeftDown 
			dm.moveto 260 * d, 0
			dm.LeftUp 
				
			hwnd = dm.GetWindow(hwnd, 4)
			dm_ret = dm.BindWindow(hwnd, "normal", "windows", "windows", 0)
			dm.moveto 0, 0
			dm.LeftDown 
			dm.moveto 0, - 130  * d
			dm.LeftUp 
		Next
	End If
 	dm_ret = dm.UnBindWindow()
Wend

Function getPartHwnds()
	Dim w, h, col, row, x, y, nx, ny, hwnd,result
	w = dm.getscreenwidth()
	h = dm.getscreenheight() - 40
	col = 5 : row = 2
	result = ""
	Call GetCursorPos(x, y)
	For i = 0 To col-1
		For j = 0 To row - 1
			nx = (x + w * i / col) mod w
			ny = (y + h * j / row) mod h
			hwnd = dm.GetPointWindow(nx, ny)
			If instr(dm.GetWindowProcessPath(hwnd), "ttha") <> 0 Then 
				result = result & " " & hwnd
			End If
		Next
	Next
	getPartHwnds = split(trim(result)," ")
End Function

Function getXY()
	Call GetCursorPos(x, y)
	hwnd = dm.GetPointWindow(x, y)
	Call dm.ScreenToClient(hwnd, x, y)
	getXY = array(x,y)
End Function

Function lClick(xy)
	dm.moveto xy(0), xy(1)
	dm.leftclick 
End Function

Function bindLClick(hwnd)
	Call bind(hwnd)
	dm.moveto 1, 1
	dm.leftdoubleclick
End Function

Function del(cnt)
	Dim i
	For i = 0 To cnt
		dm.KeyPress 8
		dm.KeyPress 46
	Next
End Function

Function hotKey(a,b)
	dm.KeyDown a
	dm.KeyPress b
	dm.KeyUp a
End Function

Function bind(hwnd)
	dm_ret = dm.BindWindow(hwnd, "normal", "windows", "windows", 0)
End Function

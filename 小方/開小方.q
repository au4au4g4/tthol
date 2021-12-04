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
MacroID=3781a671-1216-431f-be63-fdde4c1be28d
Description=開小方
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
Set dm = createobject("dm.dmsoft")

w = dm.getscreenwidth()
h = dm.getscreenheight() - 40
c = 5
r = 2

three = array("SnowFlakes", "FrozenBubbles", "SnowRollers")
sky = array("tongxia", "xingzhi", "cangcui", "fenxue", "qiuyun", "yaolan", "yuanxiang", "yunmiao", "guiyan", "daohua")
fox  = array("ef7a82","eaff56","ca6924","a4e2c6","b0a4e3","e4c6d0","e3f9fd","afdd22","a2cc8c","f1c143")
kylin  = array("tzuyeh","rihdan","posiao","jihhsi","riiyuu","jhongrih","yangrih","chenrih","puurih","renjing")
bad = array("Hydrogenn", "Nitrogenn", "Oxygenn", "Fluorine", "Helium", "Neonnn", "Argonn", "Krypton", "Xenonn", "Radonn")
account = array(bad,sky,fox)

hwnd = 0
windowCnt = UBound(split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4), ","))
For j = 0 To UBound(account)
	For i = 0 To UBound(account(j))
		// 開程式
		MoveTo 121,1058
		LeftClick 1
		While windowCnt >= UBound(split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4), ","))
			Delay 500
		Wend
		windowCnt = windowCnt + 1

		// 調整位置
		hwnd = dm.FindWindow("", "絕代方程式")
		dm_ret = dm.SetClientSize(hwnd, w / c - 2, h / r - 32)
		x = (i mod 5) * w / c - 7
		y = (i \ 5) * h / r
		dm.MoveWindow hwnd, x, y
		
		// 輸入帳密
		login = dm.FindWindow("","登錄")
		edits = split(dm.EnumWindow(login, "", "Edit", 2 + 4), ",")
		dm_ret = dm.BindWindow(edits(0), "normal", "windows", "windows", 0)
		call del(10)
		dm.SendString edits(0), account(j)(i)
		dm.SendString edits(1),"gj83dj4"
	Next
	
	// 換頁
	KeyDown "Ctrl", 1
	KeyDown "Win", 1
	KeyPress "Right", 1
	KeyUp "Win", 1
	KeyUp "Ctrl", 1
Next

Function del(cnt)
	Dim i
	For i = 0 To cnt
		dm.KeyPress 8
		dm.KeyPress 46
	Next
End Function

Function findWindowByTitle(title)
	hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4), ",")
	If UBound(hwnds) = - 1  Then 
	TracePrint 999
		findWindowByTitle = 0
	Else 
		findWindowByTitle = hwnds(0)
	End If
End Function

[General]
SyntaxVersion=2
BeginHotkey=98
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=98
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=cd61f9b4-695f-44f7-99e5-a7cc69cd6b1b
Description=ºû«ù³s½u
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Import "QMScript/Util.vbs" : Set u = New Util
Set dm = createobject("dm.dmsoft")
Set cash = CreateObject("Scripting.Dictionary")

hwnds = t.getAllHwnds()
While True
	min = Minute(Now) 
	If (min mod 2) = 0 Then 
		call keep()
	End If
	If (min mod 30) = 0 Then 
		call record()
	End If
	Delay 60 * 1000
Wend

Function keep()
	For Each hwnd In hwnds
		t.init(hwnd)
		If t.isOffLine() Then 
			dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
			Call lClick(array(37, 14))
		ElseIf (t.x < 10) * (t.locationNo = 229) * t.isStart Then
			dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
			Call lClick(array(92, 14))
		End If
	Next	
End Function

Function lClick(xy)
	dm.moveto xy(0), xy(1)
	Delay 10
	dm.leftclick 
	Delay 100
End Function

Function record()
	str = ""
	For Each hwnd In hwnds
		t.init (hwnd)
		earn = t.cash - cash.item(hwnd)
		If earn < 10000 Then 
			str = str + t.id + + "/" + cstr(earn) + "-"
		End If
		cash(hwnd) = cash(hwnd) + earn
	Next
	u.pushMsg str
End Function

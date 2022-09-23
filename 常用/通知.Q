[General]
SyntaxVersion=2
BeginHotkey=55
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=4b30791e-e27d-4b7a-a6fd-7267e28b0005
Description=³qª¾
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
Set http = createObject("winhttp.winhttprequest.5.1")
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Import "QMScript/Util.vbs" : Set u = New Util

Set re = New RegExp 
re.Pattern= "±K§A¡G|¹ô"

hwnd = t.getHwnd()
t.init(hwnd)
While true
	str = t.getMsg()
	If (last <> str) * re.Test(str) Then 
		TracePrint str
		u.pushMsg str
		last = str
	End If
	Delay 1000
Wend


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
MacroID=4b30791e-e27d-4b7a-a6fd-7267e28b0005
Description=小方通知
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
Set http = createObject("winhttp.winhttprequest.5.1")
Import "Tthbn.vbs" : Set t = New Tthbn
Import "Util.vbs" : Set u = New Util

hwnd = dm.GetForegroundWindow()
t.init(hwnd)
While true
	str = t.getMsg()
	TracePrint str
	If (last<>str) * instr(str," 密你：") Then 
		u.pushMsg str
		last = str
	End If
	Delay 1000
Wend


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
MacroID=feadbf3e-88a7-4d9f-9194-1b44adb98e71
Description=小方學技能
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
Import "Tthbn.vbs" : Set t = New Tthbn
 
hwnds = t.getArrHwnds("子夜 日旦 日始 日夕 日禺 日中 日央 日沉 夕食 人靜")
For Each hwnd In hwnds
	t.init(hwnd)
	For Each code In split(Replace(dm.GetClipboard(), """", ""), chr(10))
		t.learn(code)
		Delay 100
	Next
Next

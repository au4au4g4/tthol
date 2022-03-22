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
MacroID=b759c16d-2047-467f-ac85-c1111be4e962
Description=мїкс
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn

hwnds = t.getPartHwnds
While True
	For Each hwnd In hwnds
		t.init(hwnd)
		t.stops
		t.talkOption &H18E0, array(1, 2, 4, 2, 2, 4, 2, 0, 8, 2)
		t.talkOption &H18C2, array(1, 2, 2)
		t.talkOption &H18E0, array(1, 2, 2)
	Next
	Delay 10*60*1000
Wend


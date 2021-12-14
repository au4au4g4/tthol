[General]
SyntaxVersion=2
BeginHotkey=101
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=101
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=7c2f4e18-8e42-4baa-b568-9f44c3d8568d
Description=¥È¤Ñ
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
Import "Tthbn.vbs" : Set t = New Tthbn

hwnds = t.getAllHwnds()
While True
	For Each hwnd In hwnds
		t.init (hwnd)
		If t.location > 1703 and t.location < 1710 Then 
			t.talkNPC
		End If
	Next
	Delay 1000
Wend


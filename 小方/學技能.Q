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
Description=¾Ç§Þ¯à
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
 
hwnds = t.getPartHwnds()
skills = split(Replace(dm.GetClipboard(), """", ""), chr(10))
For Each hwnd In hwnds
	t.init (hwnd)
	For Each skill In skills
		t.learn(skill)
		Delay 300
	Next
Next

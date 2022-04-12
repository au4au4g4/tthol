[General]
SyntaxVersion=2
BeginHotkey=100
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=100
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=c922930b-efd2-41af-b9e1-2a441a1c76d0
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
hwnd = dm.FindWindow("_UJONLINE_", "Tthol")

Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init

skillCodes = split(Replace(dm.GetClipboard(),"""",""), chr(10))
For Each code In skillCodes
	code = split(code, chr(9))
	t.learnSkill (HEX(code(0)*100+code(1)))
	Delay 300
Next
Beep
//
//For i=53 to 99
//	t.learnSkill (HEX(270*100+i))
//	Delay 300
//Next
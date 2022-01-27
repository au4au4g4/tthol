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
MacroID=86fffc39-de60-4463-a9b3-993ece6e42b9
Description=§R§Þ¯à
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

Import "QMScript/Tthol.vbs" : Set t = New Tthol
t.init (hwnd)

//arr = array(array(586,1),array(587,1),array(588,1))
//For Each code In arr

skillCodes = split(Replace(dm.GetClipboard(),"""",""), chr(10))
For Each code In skillCodes
	code = split(code, chr(9))
	t.reomveSkill (HEX(code(0)*100+code(1)-1))
	Delay 300
Next
Beep


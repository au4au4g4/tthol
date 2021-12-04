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
MacroID=2abf892c-d04e-449a-9352-ebdb596a64cc
Description=CE
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
RunApp "..\CE\Cheat Engine.exe"

Delay 1700
hwnd = dm.FindWindow("", "Cheat")

For ubound (split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4), ",")) + 1
	dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
	dm.moveto 10, 10
	dm.leftclick 
	Delay 700
	dm_ret = dm.UnBindWindow()
	dm.KeyPress 38
	Delay 700
	dm.KeyPress 13
	Delay 700
	dm.KeyPress 13
	Delay 700
	dm.KeyPress 13
	Delay 700
	dm.KeyDown 17
	dm.KeyPress 105
	dm.KeyUp 17
	Delay 700
Next

[General]
SyntaxVersion=2
BeginHotkey=48
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=570baa1f-73ed-40f0-ace8-203ad2744822
Description=pull
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
KeyDown 18, 1
KeyPress 50, 1
KeyUp 18, 1

RunApp "cmd.exe /c cd QMScript && git reset --hard && git pull"
Delay 2000
While dm.findwindow("ConsoleWindowClass", "") <> 0
	Delay 1000
Wend
hwnd = split(dm.EnumWindow(0,"","Afx:ToolBar",2),",")(3)
dm_ret = dm.BindWindow(hwnd, "normal", "windows", "windows", 0)
dm.moveto 80,15
dm.leftclick
Delay 2000

KeyDown "Ctrl", 1
KeyPress "2", 1
KeyUp "Ctrl", 1
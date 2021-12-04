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
MacroID=91e9612d-20d6-4caf-a126-936830484abe
Description=破解
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

hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4 + 8 + 16), ",")
For Each hwnd In hwnds
	call dm.WriteData(hwnd,"<tthbn.bin>+ACD08","E1 E7 F6 B7 84 AC ED B4 E1 E0 A0 9C A9 FB ED E6 B6 8B B1")
	call dm.WriteData(hwnd,"<ttha.bin>+3B162","EB")
	call dm.WriteInt(hwnd,"<ttha.bin>+3FD53",0,&H6E)
Next

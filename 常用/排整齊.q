[General]
SyntaxVersion=2
BeginHotkey=99
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=d70db4da-7ce9-48a0-984d-657fc72d9df8
Description=排整齊
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

w = dm.getscreenwidth()
h = dm.getscreenheight() - 40
c = 5
r = 2
ew = w / c
eh = h / r

hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1), ",")
For each hwnd in hwnds
	dm_ret = dm.SetWindowSize(hwnd, ew+14, eh+7)
	dm_ret = dm_ret = dm.GetWindowRect(hwnd, x1, y1, x2, y2)
	x = ((x1 + 7) \ ew) * ew - 7
	y = (y1 \ eh) * eh
	dm.MoveWindow hwnd, x, y
Next

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
w = dm.getscreenwidth() / 5
h = (dm.getscreenheight() - 40) / 2

hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1+4), ",")
For each hwnd in hwnds
    dm_ret = dm.SetWindowSize(hwnd, w + 14, h + 7)
    dm_ret = dm.GetWindowRect(hwnd, x1, y1, x2, y2)
    cx = (x1 + x2) / 2
    cy = (y1 + y2) / 2
    x = ((cx + 7) \ w) * w - 7
    y = (cy \ h) * h
    dm.MoveWindow hwnd, x, y
Next

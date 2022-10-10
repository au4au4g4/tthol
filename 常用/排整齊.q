[General]
SyntaxVersion=2
BeginHotkey=51
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set dm = createobject("dm.dmsoft")
w = dm.getscreenwidth() / 5
h = (dm.getscreenheight() - 40) / 2

hwnds = t.getAllHwnds()
For each hwnd in hwnds
    dm_ret = dm.SetWindowSize(hwnd, w + 14, h + 7)
    t.init hwnd
	i = t.memberST(t.account)(0)
    x = (i mod 5) * w - 7
    y = (i \ 5) * h
    dm.MoveWindow hwnd, x, y
Next

hwnds = split( dm.EnumWindow(0,"虛擬機器連線","",1+4),",")
for each hwnd in hwnds
	dm.MoveWindow hwnd, -7, 0
next
[General]
SyntaxVersion=2
BeginHotkey=100
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=ee88f05a-5cdc-4843-97ce-fcf748188f26
Description=打幣統計
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
Import "Tthbn.vbs" : Set t = New Tthbn
Import "Util.vbs" : Set u = New Util
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

hwnds = t.getAllHwnds()
For Each hwnd In hwnds
	t.init(hwnd)
	period = (timeGetTime() - t.start) / 60 / 60 / 1000
	Redim Preserve datas(ubound(datas) + 1)
	datas(ubound(datas)) = array(t.id, t.level, t.place, period, t.monster, t.expp, t.money, t.cash)
Next
u.post "掛機", datas



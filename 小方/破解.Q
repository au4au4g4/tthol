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
MacroID=21bf90ac-ccf2-4fb6-90c9-23cb0b5d4566
Description=¯}¸Ñ
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
Import "Tthbn.vbs" : Set t = New Tthbn
Set re = New RegExp
Set dm = createobject("dm.dmsoft")

hwnds = array(t.getHwnd())
//hwnds = t.getAllHwnds()
For each hwnd in hwnds
	t.init hwnd
	t.crack
	t.moveSpeed(110)
	t.atkRange(4)
	t.atkSpeed(100)
	t.addHP
	t.addMP
	t.defBuff(1)
	t.reLoginMin(2)
Next
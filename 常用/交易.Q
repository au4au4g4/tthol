[General]
SyntaxVersion=2
BeginHotkey=103
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=a4173351-bfb2-4d11-927b-01aca90a14bb
Description=¥æ©ö
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
Import "Tthbn.vbs" : Set t = New Tthbn

UserVar keywords "ÃöÁä¦r" : keywords = split(keywords," ")
buyer = t.getHwnd()
sellers = t.getAllHwnds()
For Each seller In sellers
	t.init seller
	t.trade buyer, keywords
Next

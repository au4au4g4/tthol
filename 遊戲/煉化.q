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
MacroID=846b6079-ddeb-4e1c-ac66-3f38b1a7651f
Description=ทาคฦ
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
Import "Tthol.vbs" : Set t = New Tthol
hwnd = dm.FindWindow("", "Tthol")
t.init(hwnd)

Set bag = t.getItems("bag")
For i = 0 To bag.Count - 1
  	Set item = bag.GetByIndex(i)
  	iID = item.Item("id")
	cID = dm.ReadIni("compound", HEX(iID), ".\shop.ini")
	TracePrint cID
	If cID <> "" Then 
		t.compound cID, item.Item("addr")
	End If
Next

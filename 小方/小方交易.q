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
MacroID=5928a3b8-0c7b-4557-acd1-f97a87294aa5
Description=小方交易
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
Import "Tthbn.vbs" : Set t = New Tthbn

UserVar keywords ""

keywords = split(keywords," ")
buyer = dm.GetForegroundWindow()
sellers = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4), ",")
For Each seller In sellers
	t.init seller
	bag = t.getBag("")
	For i = 0 To UBound(bag)
		Set item = bag(i)
		For Each keyword In keywords
			If instr(item.item("name"), keyword) <> 0 Then 
				t.trade buyer, item
			End If
		Next
	Next
Next

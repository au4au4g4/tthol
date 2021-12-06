[General]
SyntaxVersion=2
BeginHotkey=103
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=103
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=b4c4da39-8b39-4ea2-992f-c0ce1f8d8a09
Description=Â\Åu
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
hwnd = t.getHwnd()
t.init hwnd

While true
	t.closeShop 	
	count = 0
	bag = t.getBag(array(""))
	For i = 0 To UBound(bag) - 1
  		Set item = bag(i)
  		name = item.Item("name")
  		price = dm.ReadIni("sale", name, ".\shop.ini")
  		If price = "" Then
  			dm.WriteIni "not_for_sale", name, "99", ".\shop.ini"
  		ElseIf count < 15 Then
			Call dm.WriteInt(hwnd, "<tthbn.bin>+" & HEX(&HA8AE8 + count * 20), 0, 12)
			Call dm.WriteInt(hwnd, "<tthbn.bin>+" & HEX(&HA8AE8 + count * 20 + 4), 0, item.Item("id"))
			Call dm.WriteInt(hwnd, "<tthbn.bin>+" & HEX(&HA8AE8 + count * 20 + 8), 0, item.Item("sn"))	
			Call dm.WriteInt(hwnd, "<tthbn.bin>+" & HEX(&HA8AE8 + count * 20 + 12), 0, item.Item("cnt"))
			Call dm.WriteInt(hwnd, "<tthbn.bin>+" & HEX(&HA8AE8 + count * 20 + 16), 0, price * 1000000)
			count = count + 1
		End If
	Next
	Call dm.WriteInt(hwnd, "<tthbn.bin>+111674", 0, count)
	t.openShop
	Delay 10*60000
Wend

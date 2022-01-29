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
MacroID=b3e9e719-083d-46fd-b8e5-ce2b1f5a2bd9
Description=開店
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
Import "QMScript/Tthol.vbs" : Set t = New Tthol
hwnd = dm.FindWindow("", "Tthol")
t.init(hwnd)

Call order("bank")
Call order("bag")
Call addItem()
Call openShop()

// ------------------------------開店------------------------------

Function addItem()
	Dim bag
	set bag = t.getItems("bag")
	shopName = dm.ReadIni("shop", "店名", ".\shop.ini")
	If shopName = "" Then 
		dm.WriteIni "shop", "店名", "99", ".\shop.ini"
	end if
	dm.WriteIni "sale", "商品名", "99", ".\shop.ini"
	dm.DeleteIni "not_for_sale","" ,".\shop.ini"
	For i = 0 To bag.Count - 1
  		Set item = bag.GetByIndex(i)
  		id = item.Item("id")
  		name = item.Item("name")
  		addr = item.Item("addr")
  		price = dm.ReadIni("sale", name, ".\shop.ini")
  		If (price = "")*(name<>last) Then
  			dm.WriteIni "not_for_sale",name,"99",".\shop.ini"
  		End If
	Next
End Function

Function openShop()
	Dim addr,count,str,bag
	count = 0
	
	// 把商品裝入RAM
	set bag = t.getItems("bag")//裝滿bag
	For i = 0 To bag.Count - 1
  		Set item = bag.GetByIndex(i)
  		id = item.Item("id")
  		amount = item.Item("amount")
  		addr = item.Item("addr")
  		name = item.Item("name")
  		price = dm.ReadIni("sale", name, ".\shop.ini")
  		If price <> "" and count < 15 Then 
			str = str + dm.ReadData(hwnd,HEX(addr+3),8) + "00000000" + toRevHex(price*1000000) + toRevHex(amount)
			count = count + 1
		End If
	Next
	Call dm.WriteData(hwnd, "[[<tthola.dat>+003ED978]+10]+B08", str)
	
	// 店名轉HEX
	shopName = dm.ReadIni("shop", "店名", ".\shop.ini")
	shopNameHex = ""
	For i = 1 To len(shopName)
		shopNameHex = shopNameHex + HEX(asc(mid(shopName, i, 1)))
	Next
	
	dm.AsmClear 
	For i = 0 To 13
		dm.AsmAdd "mov cl,00" + mid(shopNameHex, 1 + 2 * i, 2)
		dm.AsmAdd "mov [esp+" & HEX(&H14+i) & "],cl"
	Next
	dm.AsmAdd "mov esi,"+HEX(read("[<tthola.dat>+003ED97C]+10"))
	dm.AsmAdd "mov edi,0"+HEX(count)
	dm.AsmAdd "push " + HEX(read("[<tthola.dat>+003ED978]+10") + &HB08)
	dm.AsmAdd "mov ecx," + HEX(read("[[<tthola.dat>+003ED978]+10]+10"))
	dm.AsmAdd "lea edx,[esp+18]"
	dm.AsmAdd "push edx"
	dm.AsmAdd "call 00441670"
	dm.AsmCall hwnd,1
End Function
Function toRevHex(num)
	dim str
	str = hex(num)
	str = String(8 - len(str), "0") & str
	str = Mid(str, 7, 2) & Mid(str, 5, 2) & Mid(str, 3, 2) & Mid(str, 1, 2)
	toRevHex = str
End Function

Function order(place)
	Dim i, item, addr,bag
	set bag = t.getItems(place)
	For i = 0 To bag.Count - 1
  		Set item = bag.GetByIndex(i)
  		addr = item.Item("addr")
  		Call dm.WriteInt(hwnd, HEX(read(t.getBagInfo(place)) + 4 * i), 0, addr)
	Next
End Function

//------------------------------印ITEM------------------------------

Function printItems(place)
	Dim str, i, item, id, name,bag
	set bag = t.getItems(place)
	
	For i = 0 To bag.Count - 1
  		Set item = bag.GetByIndex(i)
  		id = item.Item("id")
  		name = item.Item("name")
  		addr = item.Item("addr")
  		price = dm.ReadIni("sale", name, ".\shop.ini")
  		nprice = dm.ReadIni("not_for_sale", name, ".\shop.ini")
  		If (price = "")*(nprice = "")*(name<>last) Then 
  			str = str & name & " = 999" & chr(10)
  			last = name
  		End If
	Next
	printItems = str
End Function


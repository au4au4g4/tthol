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
MacroID=71ec21f7-a2a9-47e4-b537-275ceefe2d34
Description=逛街
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
Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
Import "QMScript/Util.vbs" : Set u = New Util

hwnd = dm.FindWindow("", "Tthol")
bossList = ""
xyList = array(array(28, 16), array(28,30), array(47,28), array(47, 15))
Dim datas : datas = array()
Dim j : j = 0

For Each xy In xyList
	t.go xy(0), xy(1)
	bossList = readShop(bossList)	
Next
u.post "市集", datas

Function readShop(list)
	btnAddr = dm.FindData(hwnd, "12000000-30000000", "BC 44 5E 00")
	For Each addr In split(btnAddr, "|")
		If dm.readint(hwnd, HEX(CLNG("&H" & addr) + 8), 0) = &H778 Then // 是商店按鈕
		TracePrint 1
			t.shopping(addr)
			Delay 1000
			
			// 將商品存入文檔
			boss = dm.ReadString(hwnd, "[" & addr & "+D4]+1E4", 0, 16)
			TracePrint 2
			If instr(list, boss) = 0 Then 
				list = list & " " & boss
				cnt = dm.ReadInt(hwnd, HEX(t.getMainAddr() + &H3D0), 0)
				TracePrint 3
				For i = 0 To cnt - 1
					base = "[[" & HEX(t.getMainAddr() + &H3D4) & "]+" & HEX(i * 4) & "]"
					id = dm.ReadInt(hwnd, base & "+4", 0) / &H100
					'product = dm.ReadString(hwnd, base & "+14", 0, 16)
					price = dm.ReadInt(hwnd, base & "+114", 0)
					'amount = dm.ReadInt(hwnd, base & "+10", 0)
					ReDim Preserve datas(j) : datas(j) = array(id,boss,price)
					j = j+ 1
				Next
			End If
		End If	
	Next
	readShop = list
End Function


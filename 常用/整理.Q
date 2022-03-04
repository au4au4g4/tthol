[General]
SyntaxVersion=2
BeginHotkey=102
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=102
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=0e8977bb-a552-405d-8d8d-608e4cf1529b
Description=整理
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn

hwnd = t.getHwnd()
tasks = split(dm.getclipboard(), chr(13) & chr(10))
For Each task In tasks
	order = split(task, "-")
	param = order(1)
	paramArr = split(order(1), " ")
	
	If order(0) = "o" or order(0) = "u" Then 
		sellers = array(t.getHwnd)
	Elseif order(0) = "i" then
		sellers = t.getAllHwnds()
	Else 
		sellers = t.getPartHwnds()
	End If

	For Each seller In sellers
		t.init seller
	 	Select Case order(0)
    		Case "o"
        		t.pop paramArr
    		Case "u"
        		t.push paramArr
    		Case "t"
				t.trades hwnd, paramArr, -1
			Case "g"
				t.init hwnd
				t.trades seller, array(paramArr(0)), paramArr(1)
			Case "a"
        		t.apply param
			Case "l"
        		t.learn paramArr
			Case "i"
    			t.updateItem paramArr
    		Case "m"
    			t.map param
    		Case "f"
    			t.frame paramArr
    		Case "b"
				Call dm.WriteData(seller, "<tthbn.bin>+24EA8", "EB0C")
				t.openbank 
				Delay 200
//				While (t.deposit + t.cash) > 1000000
//					t.withdrawal (99999999 - t.cash)
//					t.withdrawal (t.deposit)
					Set coin = CreateObject("Scripting.Dictionary")
					coin.add "id", &H5e16
					coin.add "cnt", int(t.cash/10^6)
					Set npc = t.findNPC("歡喜")
					t.talkOption npc.item("id"), array(1,5)
					t.buy npc, coin	
//				Wend
				Call dm.WriteData(seller, "<tthbn.bin>+24EA8", "8B45")
		End Select
	Next  
Next

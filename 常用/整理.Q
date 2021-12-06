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
sellers = t.getAllHwnds()
tasks = split(dm.getclipboard(), chr(13) & chr(10))
For Each task In tasks
	keywords = split(task," ")
	Select Case keywords(0)
    	Case "o"
    		t.init hwnd
        	t.pop keywords
    	Case "u"
    		t.init hwnd
        	t.push keywords
    	Case "t"
			For Each seller In sellers
				t.init seller
				t.trades hwnd, keywords
			Next
    	Case "b"
			For Each hwnd In sellers
				Call dm.WriteData(hwnd, "<tthbn.bin>+24EA8", "EB0C")
				t.init hwnd
				While (t.deposit + t.cash) > 1000000
					t.withdrawal (99999999 - t.cash)
					t.withdrawal (t.deposit)
					Set coin = CreateObject("Scripting.Dictionary")
					coin.add "id", &H5e16
					coin.add "cnt", int(t.cash / 10 ^ 6)
					Set npc = t.findNPC("歡喜")
					t.talkOption npc.item("id"), array(1,5)
					t.buy npc, coin	
				Wend
				call dm.WriteData(hwnd,"<tthbn.bin+24EA8","8B45")
			Next
	End Select
Next

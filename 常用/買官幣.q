[General]
SyntaxVersion=2
BeginHotkey=105
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=ac989b52-6698-448e-99ac-58cb5e104251
Description=¶R©x¹ô
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

hwnds = t.getAllHwnds()
For Each hwnd In hwnds
	call dm.WriteData(hwnd,"<tthbn.bin>+24EA8","EB0C")
	t.init(hwnd)
	While (t.deposit + t.cash) > 1000000
		t.withdrawal (99999999 - t.cash)
		t.withdrawal (t.deposit)
		Set coin = CreateObject("Scripting.Dictionary")
		coin.add "id", &H5e16
		coin.add "cnt", int(t.cash / 10 ^ 6)
		Set npc = t.findNPC("Åw³ß")
		t.talkOption npc.item("id"), array(1,5)
		t.buy npc, coin	
	Wend
	call dm.WriteData(hwnd,"<tthbn.bin+24EA8","8B45")
Next

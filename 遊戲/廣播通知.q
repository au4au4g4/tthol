[General]
SyntaxVersion=2
BeginHotkey=104
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=101
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=c3097257-35cb-475a-9f3d-5455c56bb6b7
Description=廣播通知
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
Import "QMScript/VbsJson.vbs" : Set json = New vbsJson
Set http = createObject("winhttp.winhttprequest.5.1")

hwnd = dm.GetForegroundWindow()
dm_ret = dm.BindWindow(hwnd, "normal", "windows", "windows", 0)
check = array(Minute(now),Minute(now),Minute(now))

While true
	str = dm.ReadString(hwnd, "[[[[[[<tthola.dat>+003ED978]+10]+1538]+174]+4]+C]", 0, 115)
	If (last<>str) * instr(str,")告訴你") Then 
		Call pushMsg(mid(str, 10))
		//dm.WriteFile "..\msg.txt",mid(str, 10)+chr(10)
		last = str
	End If
	
	i = 0
	While dm.ReadIni("broadcast", "word" & i, ".\shop.ini") <> ""
		min = dm.ReadIni("broadcast", "min" & i, ".\shop.ini")
		word = dm.ReadIni("broadcast", "word" & i, ".\shop.ini")
		If check(i) = Minute(now) Then 
			Delay 400
			dm.SendString hwnd, word
			Delay 400
			dm.KeyPressChar "enter"
			check(i) = (check(i) + min) mod 60
		End If
		i = i + 1
	Wend
	Delay 500
Wend

// line notify
Function pushMsg(str)
	http.open "POST", "https://notify-api.line.me/api/notify", False
	http.setrequestheader "Content-Type", "application/x-www-form-urlencoded"
	http.setrequestheader "Authorization","Bearer MSakOPLvpgbeYtuiU8K2cRUjgiHruOQiQ37Jd7CNoKz"
	http.send "message="&str
End Function

// line bot
Function pushMsg2(str)
	http.open "POST", "https://api.line.me/v2/bot/message/push", False
	http.setrequestheader "Content-Type", "application/json"
	http.setrequestheader "Authorization","Bearer zRMgWml4Wa7NT7BB26mCCE3N2dO3eVeRE/d+fm0YgQikEbbfXZubZF2lRW1JUcEbbA3CcV6twklA/7x1mERZllnjHw4r3tWAqgaORQ9sbJ9s2HCykn661IzDcnavSfIh2cERR4Tr7TrO2WHLus1+2QdB04t89/1O/w1cDnyilFU="
	Set message = CreateObject("Scripting.Dictionary")
	message.add "type", "text"
	message.add "text", str
	Set bag = CreateObject("Scripting.Dictionary")
	bag.add "to", "U86cb01451b4bf90220e57bf5ec133348"
	bag.add "messages", array(message)
	TracePrint json.encode(bag)
	http.send json.encode(bag)
End Function


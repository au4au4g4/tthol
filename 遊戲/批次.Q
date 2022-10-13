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
MacroID=00ac2641-0e21-4154-bda7-61229b91eab1
Description=批次
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
Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init 
Set dm = createobject("dm.dmsoft")

Set dic = CreateObject("Scripting.Dictionary")
dic.Add "莫愁", 4
dic.Add "天外", 7
dic.Add "無名", 4
dic.Add "飛雁", 5
dic.Add "天師", 4
dic.Add "少林", 6

accounts = t.IDs("t1")
For i = 0 To UBound(accounts)
	TracePrint i & "/" & accounts(i)
	t.login accounts(i), "gj83dj4"
	t.choose
//	call joinSky()
//	Call joinGod()
//	Call make()
	t.wear2
	t.logout 
//WaitKey
Next

Function joinSky()
	If t.location <> "天外天試煉場" Then 
		t.learnSkills (array(array(4, 2), array(4, 3), array(5, 2), array(5, 3), array(1, 1), array(9, 1), array(7, 1)))
		t.talkNPC &H1936, array(1, 2, 2, 4, 2, 2, 2, 2, 2, 2)'魏布衣 
		t.transWait "天外天試煉場"	
	End If
	
	While t.getBag(array(&H6D81)).Count < 1
		Set monster = t.getMonster
		key = monster.item("key")
		If key <> 0 Then 
			t.atk (monster)
			While key = t.getMonster.item("key")
				Delay 500
			wend
		End If
		set poisons = t.getBag(array(&H6D82))
		If poisons.count > 0 Then 
			t.use poisons.getbyindex(0)
		End If
	Wend
	t.go 30,5
	t.trans
	t.talkNPC &HEA64&, array(4)
	t.transWait "流星村"
	
	t.go 44, 47
	t.use t.getBag(array(&H6D81)).getbyindex(0)
	t.talkNPC &H1936, array(1,2)'魏布衣
End Function

Function joinGod()
	t.talkNPC &H1dd4, array(1, array(2, 15)) 
	t.talkNPC &H1dd5, array(1, array(2, 6))
	TracePrint t.chat
	t.talkNPC &H1dd5, array(dic.item(mid(t.chat, 4, 2)), 2, 2)
	TracePrint t.chat
	t.talkNPC &H1dd5, array(dic.item(mid(t.chat, 4, 2)), 2, 2)
	TracePrint t.chat
	t.talkNPC &H1dd5, array(dic.item(mid(t.chat, 4, 2)), 2, 2, 2)
	While instr(t.chat, "從不起眼") < 1
		t.trans	
	Wend
	t.talkNPC &HEA64&, array(2, 2, 2, 2)
	t.transWait "成都少城"
	t.go 112, 136
	t.talkNPC &H1dd4, array(1, array(2, 5))	
End Function

Function make()
	t.compound0 &H4E33,t.getBag(array("大印記")).getByIndex(0)
End Function

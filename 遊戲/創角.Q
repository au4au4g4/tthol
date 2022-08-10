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
MacroID=48d9c45c-e3ac-43b2-b544-dfbc0e08e16b
Description=創角
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

accounts = array("idddidffak","chkchjacai","eekhdfahic","ikacidbdjf","jbkjhbdfck","cjfieiekde","abcdjkecai","bkhbdkchdc","fcafceaach","edfjjfddef","akefcaaedk","djdbjhfake","kjbbdbfjhj","ccaabjcdbb","cbjacicdbb","kbcfkadbcf","ejeaafeeia","jkheebddae","dajbhhiecj","bakhfiffji")
ids = array("尖葉槭","土田七","野牽牛","尖瓣花","萬年麻","漿果莧","蔓竹杞","仙丹花","枯里珍","彩虹菊")
t.speed (150)
For k = 10 To UBound(accounts)
	TracePrint k
	TracePrint accounts(k)
	TracePrint ids(k)
	t.login accounts(k), "gj83dj4"
	t.del accounts(k), ids(k)
	t.create ids(k)
	t.choose
	
	t.trans	'進入傳點
	t.talkNPC &HEA64&, array(2, 2, 2, 4, 2)'進入對話
	t.go 26, 9
	t.talkNPC &H1901, array(1, array(2, 5))'明若蘭1
	t.go 44, 7
	t.trans '*入弓手傳點
	Delay 3000
	
	Set monster = t.getMonster
	key = monster.item("key")
	While key = monster.item("key")
		If monster.count = 1 Then 
			t.go 34, 37
			Set monster = t.getMonster
			key = monster.item("key")
		Else 
			t.atk (monster)
			Delay 1000
			Set monster = t.getMonster		
		End If
	Wend
	
	t.go 6, 40
	t.trans	'*出弓手傳點
	t.talkNPC &HEA64&, array(array(2,20))'出弓手對話
	t.talkNPC &H1901, array(1,array(2,40))'明若蘭3
	Delay 500
	t.wear t.getBag(array("陸天灝衣")).GetByIndex(0)
	t.talkNPC &H1901, array(1,array(2,30))'明若蘭3
	t.learnSkills(array(array(8,1),array(4,1),array(5,1),array(2,1),array(3,1),array(6,1),array(522,1),array(523,1),array(524,1),array(525,1),array(550,1),array(586,1),array(588,1),array(587,1)))
	t.trans'*惡乞丐傳點
	
	Set monster = t.getMonster
	key = monster.item("key")
	While key = monster.item("key")
		If monster.count = 1 Then 
			t.go 99,36
			Set monster = t.getMonster
			key = monster.item("key")
		Else 
			t.atk (monster)
			Delay 1000
			Set monster = t.getMonster		
		End If
	Wend
	
	t.go 20, 69
	t.talkNPC &H1902, array(1,array(2,30))'惡乞丐	 
	t.trans	'*入山洞傳點
	Delay 1000
	t.go 17, 12
	
	t.trans	'*泉水
	t.talkNPC &HEA64&, array(array(2,10))
	t.compound0 &H4E5F, t.getBag(array("泉水")).GetByIndex(0)
	t.go 8, 10
	
	t.transBySN 360 '*出山洞傳點
	t.talkNPC &HEA64&, array(array(2,10))	'出山洞對話
	t.go 61, 68
	
	t.talkNPC &H1901, array(1,array(2,90))	'明若蘭2
	t.go 109, 64
	t.trans	'*結束傳點
	Delay 6000
	
	t.logout 
Next

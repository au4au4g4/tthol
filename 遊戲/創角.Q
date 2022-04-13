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

accounts = array("hjeahefdec","kkhekbecih","bakfdkfibd","dfbcjiiebh","cdjfajekab","kfhiebkaif","fajbhfdebh","jcddhchhaa","abheaaadde","aekfhabfej","jjjkjccddf","ebkkafbhhk","kbikfabfdd","aibiakcbfe","cjifbjaaej","kjbajbbjfc","ikfjfcijhh","ikjhkcdbke","zfbhjfcaji","fhbkhfkkcb","eekckhhebh","jjbebbfihi","eihhkhacic","hdbccddeaj","bifhcjbjkc","jfhecbjdeh","kkkiiaejfa","eeehhhidfb","kfdhbaiccf","cjbekiehik","bffaejijae","akcjikddhi","faiafehiif","beaiejkehj","ecabbhkjdd","edaekhhiji","djekfhcbaj","cdeffibkdc","jckidbifeh","beejdjaida","cjheahddai","cakehiehba","zakfhjhdhf","idaejcfikj","dafhhhdeka","dkccjhhbje","zajbkeebij","cachhifkic","kfkbakdafk","zfbfdhbkid","jiccfbfaai","jjcfebcafb","kabbcaecaf","zeecefdfeh","cibjjidjba","dffhchccef","ijkaccahdj","hfbfbcbjjf","ffbhaechdk","keachaicad","hjaficbjdi")
ids = array("乾為天","天地否","天澤履","天山遯","天火同人","天水訟","天雷無妄","天風姤","地天泰","坤為地","地澤臨","地山謙","地火明夷","地水師","地雷復","澤天夬","澤地萃","兌為澤","澤山咸","澤火革","澤水困","澤雷隨","澤風大過","山天大畜","山地剝","山澤損","艮為山","山火賁","山水蒙","山雷頤","山風蠱","火天大有","火地晉","火澤睽","火山旅","離為火","火水未濟","火雷噬嗑","火風鼎","水天需","水地比","水澤節","水山蹇","水火既濟","坎為水","水雷屯","水風井","雷天大壯","雷地豫","雷澤歸妹","雷山小過","雷火豐","雷水解","震為雷","雷風恒","風天小畜","風地觀","風澤中孚","風山漸","風火家人","風水渙","風雷益","巽為風")
accounts = array("fdjtykutfk")
ids = array("風雷益阿")
For k = 0 To UBound(accounts)
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
	t.talkNPC &H1901, array(1, array(2,5))'明若蘭1
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

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
MacroID=0d3a083f-71b8-4bef-ba03-6881197d0cbd
Description=創角
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
Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init 

accounts = array("gdhcfccjbb","iidaacjggb","hhecdcebah","chbddbbhge","bbdjhcibfb","igbidcecjf","icigfcifaf","chedbceiaf","ajfcedfabd","eadejicbaj","giddddgcgf","hgghgbjdci","jbgibcfgii","ebfgcfghff","gdhcfccjbb","aiaaefehfi","egbehifhbb","dbefhbhggf","gijhdigieg","hbdeeeaihi","iidaacjggb","hfeijjgehh","gagdeijbbc","hhecdcebah","idijbjgeia","ifcfcjecib","deijejjaig","dgheffdbdb","gdgdbbgcce","jhgfefccjg","cdaecjhbgb","hhddcjjcab","dcdgdabjii","jhafbahhda","cbebbgcdff","ihahehighf","gbeeahdecf","ibcggdjjdi","cbajgfheii","hdbbbigiic","dhhgfifdaj","aghieagehb","cdjbbadfjh","dhigagfbdj","bbjdeghgcj","cfjhjaafhe","gcbeaebcfh","gfiggdegih","hcicdbceha","bbgbhcegdi","fgcjjhhhjb","gdbiiadhig","ibgggghdai","idhbiejfjj","ejgfdhaagc","iebaddgcfj","adfdebjacf","bjihaaddec","biccbbfgij","daaiihdjfe")
ids = array("三角榕","洋茴香","粗勒草","一葉蘭","洋甘菊","龜背芋","雞冠花","白芒草","橡皮樹","銅錢草","鶴望蘭","姑婆芋","五色梅","三色堇","三角榕","鬼針草","白荳蔻","散尾葵","洋香菜","巴西里","洋茴香","蝦夷蔥","香蘭葉","粗勒草","馬鞭草","萬年青","鏡面草","鳳凰木","夏枯草","飛羽木","虎尾蘭","魚腥草","青剛櫟","山茱萸","梵天花","風車草","鈍葉桂","懸星花","庭菖蒲","車桑子","番茉莉","羊角藤","咬人貓","紅菽草","糙葉樹","菟絲子","鐵角蕨","小蝦花","啤酒花","蓮子草","松葉菊","芝麻菜","大星蕨","落羽杉","稀子蕨","洋玉蘭","樹葡萄","山蘇花","紅棕櫚","七葉膽","桔梗蘭","玉蜂蘭","星點木","水線草","金花茶","人心果","烏柑仔","水茄苳","貓腥菊","山白蘭","木芙蓉","彩虹菊","青棉花","星蘋果","水毛花","倒地鈴","貝殼杉")

For k = 2 To 2
	TracePrint k
	TracePrint accounts(k)
	TracePrint ids(k)
	t.login accounts(k), "gj83dj4"
	t.create ids(k)
	
	t.trans	'進入傳點
	t.talkNPC &HEA64&, array(2, 2, 2, 4, 2)'進入對話
	t.go 26, 9
	Delay 1000
	t.talkNPC &H1901, array(1, array(2,5))'明若蘭1
	t.go 44, 7
	Delay 3000
	t.trans '*入弓手傳點
	Delay 3000
	
	Set monster = t.getMonster
	key = monster.item("key")
	While key = monster.item("key")
		If monster.count = 1 Then 
			t.go 34, 37
			Delay 1000
			Set monster = t.getMonster
			key = monster.item("key")
		Else 
			t.atk (monster)
			Delay 1000
			Set monster = t.getMonster		
		End If
	Wend
	
	t.go 6, 40
	Delay 3000
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
			t.go 109,36
			Delay 1000
			Set monster = t.getMonster
			key = monster.item("key")
		Else 
			t.atk (monster)
			Delay 1000
			Set monster = t.getMonster		
		End If
	Wend
	
	t.go 20, 69
	Delay 2000
	t.talkNPC &H1902, array(1,array(2,30))'惡乞丐	 
	t.trans	'*入山洞傳點
	Delay 1000
	t.go 17, 12
	Delay 3000
	
	t.trans	'*泉水
	t.talkNPC &HEA64&, array(array(2,10))
	t.compound0 &H4E5F, t.getBag(array("泉水")).GetByIndex(0)
	t.go 8, 10
	Delay 2000
	
	t.transBySN 360 '*出山洞傳點
	t.talkNPC &HEA64&, array(array(2,10))	'出山洞對話
	t.go 61, 68
	Delay 2000
	
	t.talkNPC &H1901, array(1,array(2,90))	'明若蘭2
	t.go 109, 64
	Delay 1000
	t.trans	'*結束傳點
	Delay 6000
	
	t.logout 
	TracePrint k
Next
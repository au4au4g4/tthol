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
MacroID=922b1d6f-345d-4b4f-8e50-f32efabf1549
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

Set dic = CreateObject("Scripting.Dictionary")
dic.Add "莫愁", 4
dic.Add "天外", 7
dic.Add "無名", 4
dic.Add "飛雁", 5
dic.Add "天師", 4
dic.Add "少林", 6

aa = array("ggfecdaehg", "jgihagbffh", "jiedihbejd", "chbddbbhge", "bbdjhcibfb", "igbidcecjf", "icigfcifaf", "chedbceiaf", "ajfcedfabd", "eadejicbaj")
cc = array("iidaacjggb", "hfeijjgehh", "gagdeijbbc", "hhecdcebah", "idijbjgeia", "ifcfcjecib", "deijejjaig", "dgheffdbdb", "gdgdbbgcce", "jhgfefccjg")
bb = array("giddddgcgf", "hgghgbjdci", "jbgibcfgii", "ebfgcfghff", "gdhcfccjbb", "aiaaefehfi", "egbehifhbb", "dbefhbhggf", "gijhdigieg", "hbdeeeaihi")
dd = array("cdaecjhbgb", "hhddcjjcab", "dcdgdabjii", "jhafbahhda", "cbebbgcdff", "ihahehighf", "gbeeahdecf", "ibcggdjjdi", "cbajgfheii", "hdbbbigiic")
ee = array("gfiggdegih", "hcicdbceha", "bbgbhcegdi")
ff = array("fgcjjhhhjb", "gdbiiadhig", "ibgggghdai", "idhbiejfjj", "ejgfdhaagc", "iebaddgcfj", "adfdebjacf", "bjihaaddec", "biccbbfgij", "daaiihdjfe")

gg = array("hjeahefdec", "kkhekbecih", "bakfdkfibd", "dfbcjiiebh", "cdjfajekab", "kfhiebkaif", "fajbhfdebh", "jcddhchhaa", "abheaaadde", "aekfhabfej")
hh = array("jjjkjccddf", "ebkkafbhhk", "kbikfabfdd", "aibiakcbfe", "cjifbjaaej", "kjbajbbjfc", "ikfjfcijhh", "ikjhkcdbke", "zfbhjfcaji", "fhbkhfkkcb")
ii = array("eekckhhebh", "jjbebbfihi", "eihhkhacic", "hdbccddeaj", "bifhcjbjkc", "jfhecbjdeh", "kkkiiaejfa", "eeehhhidfb", "kfdhbaiccf", "cjbekiehik")
jj = array("bffaejijae", "akcjikddhi", "faiafehiif", "beaiejkehj", "ecabbhkjdd", "edaekhhiji", "djekfhcbaj", "cdeffibkdc", "jckidbifeh", "beejdjaida")
kk = array("cjheahddai", "cakehiehba", "zakfhjhdhf", "idaejcfikj", "dafhhhdeka", "dkccjhhbje", "zajbkeebij", "cachhifkic", "kfkbakdafk", "zfbfdhbkid")
ll = array("jiccfbfaai", "jjcfebcafb", "kabbcaecaf", "zeecefdfeh", "cibjjidjba", "dffhchccef", "ijkaccahdj", "hfbfbcbjjf", "ffbhaechdk", "keachaicad")

accounts = hh
'ids = array("乾為天","天地否","天澤履","天山遯","天火同人","天水訟","天雷無妄","天風姤","地天泰","坤為地","地澤臨","地山謙","地火明夷","地水師","地雷復","澤天夬","澤地萃","兌為澤","澤山咸","澤火革","澤水困","澤雷隨","澤風大過","山天大畜","山地剝","山澤損","艮為山","山火賁","山水蒙","山雷頤","山風蠱","火天大有","火地晉","火澤睽","火山旅","離為火","火水未濟","火雷噬嗑","火風鼎","水天需","水地比","水澤節","水山蹇","水火既濟","坎為水","水雷屯","水風井","雷天大壯","雷地豫","雷澤歸妹","雷山小過","雷火豐","雷水解","震為雷","雷風恒","風天小畜","風地觀","風澤中孚","風山漸","風火家人","風水渙","風雷益","巽為風")
For i = 0 To UBound(accounts)
	TracePrint i & "/" & accounts(i) & "/" & ids(i)
	t.login accounts(i), "gj83dj4"
	t.choose
	'call joinSky()
	'call joinGod()
	'Call make()
//	t.wear2
//	t.logout 
WaitKey
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
			while key = t.getMonster.item("key")  
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
	t.talkNPC &H1dd5, array(dic.item(mid(t.chat, 4, 2)), 2, 2)
	t.talkNPC &H1dd5, array(dic.item(mid(t.chat, 4, 2)), 2, 2)
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

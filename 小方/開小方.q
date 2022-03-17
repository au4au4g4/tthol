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
MacroID=3781a671-1216-431f-be63-fdde4c1be28d
Description=開小方
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set dm = createobject("dm.dmsoft")

w = dm.getscreenwidth()
h = dm.getscreenheight() - 40
c = 5
r = 2

aa =  array(210, 1, array("ggfecdaehg","jgihagbffh","jiedihbejd","chbddbbhge","bbdjhcibfb","igbidcecjf","icigfcifaf","chedbceiaf","ajfcedfabd","eadejicbaj"))
bb =  array(210, 1, array("giddddgcgf","hgghgbjdci","jbgibcfgii","ebfgcfghff","gdhcfccjbb","aiaaefehfi","egbehifhbb","dbefhbhggf","gijhdigieg","hbdeeeaihi"))
cc =  array(210, 1, array("iidaacjggb","hfeijjgehh","gagdeijbbc","hhecdcebah","idijbjgeia","ifcfcjecib","deijejjaig","dgheffdbdb","gdgdbbgcce","jhgfefccjg"))
dd =  array(210, 1, array("cdaecjhbgb","hhddcjjcab","dcdgdabjii","jhafbahhda","cbebbgcdff","ihahehighf","gbeeahdecf","ibcggdjjdi","cbajgfheii","hdbbbigiic"))
ee =  array(210, 1, array("dhhgfifdaj","aghieagehb","cdjbbadfjh","dhigagfbdj","bbjdeghgcj","cfjhjaafhe","gcbeaebcfh","gfiggdegih","hcicdbceha","bbgbhcegdi"))
ff =  array(210, 1, array("fgcjjhhhjb","gdbiiadhig","ibgggghdai","idhbiejfjj","ejgfdhaagc","iebaddgcfj","adfdebjacf","bjihaaddec","biccbbfgij","daaiihdjfe"))

gg =  array(230, 2, array("hjeahefdec","kkhekbecih","bakfdkfibd","dfbcjiiebh","cdjfajekab","kfhiebkaif","fajbhfdebh","jcddhchhaa","abheaaadde","aekfhabfej"))
hh =  array(230, 2, array("jjjkjccddf","ebkkafbhhk","kbikfabfdd","aibiakcbfe","cjifbjaaej","kjbajbbjfc","ikfjfcijhh","ikjhkcdbke","zfbhjfcaji","fhbkhfkkcb"))
ii =  array(230, 2, array("eekckhhebh","jjbebbfihi","eihhkhacic","hdbccddeaj","bifhcjbjkc","jfhecbjdeh","kkkiiaejfa","eeehhhidfb","kfdhbaiccf","cjbekiehik"))
jj =  array(230, 2, array("bffaejijae","akcjikddhi","faiafehiif","beaiejkehj","ecabbhkjdd","edaekhhiji","djekfhcbaj","cdeffibkdc","jckidbifeh","beejdjaida"))
kk =  array(230, 2, array("cjheahddai","cakehiehba","zakfhjhdhf","idaejcfikj","dafhhhdeka","dkccjhhbje","zajbkeebij","cachhifkic","kfkbakdafk","zfbfdhbkid"))
ll =  array(230, 2, array("jiccfbfaai","jjcfebcafb","kabbcaecaf","zeecefdfeh","cibjjidjba","dffhchccef","ijkaccahdj","hfbfbcbjjf","ffbhaechdk","keachaicad"))
mm =  array(230, 2, array("hjaficbjdi"))

teams = array(gg,hh,ii)

windowCnt = UBound(t.getAllHwnds())
For j = 0 To UBound(teams)
	team = teams(j)
	For i = 0 To UBound(team(2))
		dm_ret = dm.UnBindWindow()
		dm.moveto 120, 1060
		dm.leftclick 
		While windowCnt >= UBound(t.getAllHwnds())	
			Delay 500
		Wend
		windowCnt = windowCnt + 1

		// 調整位置
		hwnd = dm.FindWindow("", "絕代方程式")
		dm_ret = dm.SetClientSize(hwnd, w / c - 2, h / r - 32)
		x = (i mod 5) * w / c - 7
		y = (i \ 5) * h / r
		dm.MoveWindow hwnd, x, y
		
		// 輸入帳密
		login = dm.FindWindow("","登錄")
		edits = split(dm.EnumWindow(login, "", "Edit", 2 + 4), ",")
		dm_ret = dm.BindWindow(edits(0), "normal", "windows", "windows", 0)
		call del(15)
		dm.SendString edits(0), team(2)(i)
		dm.SendString edits(1), "gj83dj4"
		comboBoxs = split(dm.EnumWindow(login, "", "ComboBox", 2 + 4), ",")
		dm_ret = dm.BindWindow(login, "normal", "windows", "windows", 0)
		dm.SendString comboBoxs(0), "飛雁山莊(花)"
		dm.SendString comboBoxs(1), "1"
		Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+E4A4", 0, 0)
		
		// 破解
		t.init hwnd
		t.crackAll team(0),team(1)
		t.login
	Next
	
	// 換頁
	KeyDown "Ctrl", 1
	KeyDown "Win", 1
	KeyPress "Right", 1
	KeyUp "Win", 1
	KeyUp "Ctrl", 1
Next

Function del(cnt)
	Dim i
	For i = 0 To cnt
		dm.KeyPress 8
		dm.KeyPress 46
	Next
End Function
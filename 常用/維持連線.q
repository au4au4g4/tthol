[General]
SyntaxVersion=2
BeginHotkey=98
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=98
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=cd61f9b4-695f-44f7-99e5-a7cc69cd6b1b
Description=ºû«ù³s½u
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Import "QMScript/Util.vbs" : Set u = New Util
Set dm = createobject("dm.dmsoft")
Set cash = CreateObject("Scripting.Dictionary")

skills = array(array(9,190,"±j¤O¨ëÀ»",8),array(25,511,"¥©¹Ü¤o»Ø",3),array(30,750,"ÁëÅé­×·Ò",10),array(30,756,"·Ò¤ß­×·Ò",1),array(34,751,"¼CÀú¤d¬î",3),array(40,753,"¸U¼C«ß¥O",15),array(36,761,"¦BÁ÷¯Pª¢¬Þ",5),array(45,762,"¤K¨ö¯«ªZ°}",5),array(80,11,"´c·N±þ»ù",10),array(80,12,"«¡©ï»ù®æ",10),array(80,18,"ÃÄ¤ý¸g",10),array(80,193,"ÃÃÅÚ¦å«´",5),array(80,194,"µéµ·ÆF«´",5),array(80,359,"¯ð´Æ¦å«´",5),array(80,360,"°­µ·ÆF«´",20))
boxes = array(array(1,"»î¦^"),array(1,"·s¤â"),array(10,"ªì¶¥"),array(20,"¤¤¶¥"),array(30,"°ª¶¥"),array(45,"ªì¶¥"),array(55,"¤¤¶¥"),array(65,"°ª¶¥"),array(70,"ªì¶¥"),array(80,"¤¤¶¥"))
maps = array(array(1,62),array(15,22),array(19,99),array(23,21),array(27,76),array(31,77),array(35,188),array(38,215),array(43,214),array(47,37),array(51,86),array(55,87),array(62,89),array(65,49),array(69,90),array(69,90),array(74,91),array(77,39),array(80,83),array(84,211),array(90,210))

hwnds = t.getAllHwnds()
While True
	min = Minute(Now)
	If (min mod 1) = 0 Then 
		call disconnect()
	End If
	If (min mod 5) = 0 Then 
		call connect()
	End If
	If (min mod 60) = 0 Then 
		Call record()
		'Call train()
	End If
	Delay 60 * 1000
Wend

Function disconnect()
	For Each hwnd In hwnds
		t.init (hwnd)
		TracePrint t.isDisConnect
		If (t.x < 10) * (t.locationNo = 229) * t.isStart + t.isDisConnect Then
			dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
			Call lClick(array(92, 14))
		End If
	Next	
End Function

Function connect()
	For Each hwnd In hwnds
		t.init(hwnd)
		If t.isOffLine() Then 
			dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
			Call lClick(array(37, 14))
		End If
	Next	
End Function

Function lClick(xy)
	dm.moveto xy(0), xy(1)
	Delay 10
	dm.leftclick 
	Delay 100
End Function

Function record()
	str = ""
	For Each hwnd In hwnds
		t.init (hwnd)
		earn = t.cash - cash.item(hwnd)
		If earn < 10000 Then 
			str = str + t.id + + "/" + cstr(earn) + "-"
		End If
		cash(hwnd) = cash(hwnd) + earn
	Next
	u.pushMsg str
End Function

Function train()
	For Each hwnd In hwnds
		t.init (hwnd)
		lv = t.level
		'°tÂI
		For 2
			t.addpoint
		Next
		'§Þ¯à
		For Each skill In skills
			slv = t.skillLv(skill(1),skill(2))
			If (lv >= skill(0)) * (slv < skill(3)) Then 
				t.learn (array(skill(1), slv + 1, slv + 1))
			End If
		Next
		'Ä_½c
		For Each box In boxes
			If (lv >= box(0)) * (lv < box(0) + 2) Then
				t.apply box(1)
				Delay 2000
				t.apply box(1)
				Delay 2000
				t.wear("¤M")
			End If
		Next
		'¦a¹Ï
		For Each map In maps
			If (lv >= map(0)) * (lv < map(0) + 2) Then
				t.map(map)
			End If
		Next
	Next	
End Function
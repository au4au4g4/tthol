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
		For 5
			t.addpoint
		Next
		'§Þ¯à
		For Each skill In skills
			slv = t.skillLv(skill(1),skill(2))
			If (lv >= skill(0)) * (slv < skill(3)) Then 
				t.learn (array(skill(1), slv + 1))
			End If
		Next
		'Ä_½c
		For Each box In boxes
			If (lv >= box(0)) * (lv < box(0) + 2) Then
				t.apply box(1)
				t.apply box(1)
				t.wear("¤M")
			End If
		Next
	Next	
End Function
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
Description=維持連線
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
Set expp = CreateObject("Scripting.Dictionary")

skills = array(array(9,190,"強力刺擊",8),array(25,511,"巧奪寸隙",3),array(30,750,"鍛體修煉",10),array(30,756,"煉心修煉",1),array(34,751,"劍盪千秋",3),array(40,753,"萬劍律令",15),array(36,761,"冰霜烈炎盾",5),array(45,762,"八卦神武陣",5),array(80,11,"惡意殺價",10),array(80,12,"哄抬價格",10),array(80,18,"藥王經",10),array(80,193,"藤蘿血契",5),array(80,194,"菟絲靈契",5),array(80,359,"荊棘血契",5),array(80,360,"鬼絲靈契",20))
boxes = array(array(1,"魂回"),array(1,"新手"),array(10,"初階"),array(20,"中階"),array(30,"高階"),array(45,"初階"),array(55,"中階"),array(65,"高階"),array(70,"初階"),array(80,"中階"))
maps = array(array(1,62),array(15,22),array(19,99),array(23,21),array(27,76),array(31,77),array(35,188),array(38,215),array(43,214),array(47,37),array(51,86),array(55,87),array(62,89),array(65,49),array(69,90),array(74,91),array(77,39),array(80,83),array(84,211),array(90,210))

hwnds = t.getAllHwnds()
While True
	min = Minute(Now)
	If (min mod 1) = 0 Then 
		call disconnect()
	End If
	If (min mod 4) = 0 Then 
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
		earn = t.cash - expp.item(hwnd)
		If earn = 0 Then 
			str = str + t.id + + "/" + cstr(earn) + "-"
		End If
		expp(hwnd) = expp(hwnd) + earn
	Next
	u.pushMsg str
End Function

Function train()
	Randomize
	For Each hwnd In hwnds
		t.init (hwnd)
		LV = t.level
		
		'配點
		all = addr("point", "all")
		For Each p In all
			parm = addr("point", p)
			If (parm(0) - t.point(p) > 0) * (parm(1) - rnd() > 0) Then 
				t.addpoint (p)
				t.addpoint (p)
			End If
		Next
		
		'學技能
		learnLVs = addr("learn", "all")
		For Each learnLV In learnLVs
			If LV - learnLV >= 0 Then 
				lLV = learnLV
			End If
		Next
		learns = addr("learn", lLV)
		For Each learn In learns
			learn = split(learn, "|")
			lLV=-1
			While lLV < t.skillLv(learn(0), learn(1))
				lLV = t.skillLv(learn(0),learn(1))
				If lLV - learn(2) < 0 Then 
					t.learn (array(learn(0), lLV + 1, lLV + 1))
				End If
			Wend
		Next
		
		'用技能
		skillLVs = addr("skill", "all")
		For Each skillLV In skillLVs
			If LV - cint(skillLV) >= 0 Then 
				sLV = skillLV
			End If
		Next
		skill = addr("skill", sLV)
		t.useSkill(skill(0))
		
		'道具
		items = addr("item", LV)
		For Each item In items
			While t.getItemCnt(item) <> 0
				t.apply item
			Wend
		Next
		
		'裝備
		weapons = addr("weapon", LV)
		For Each weapon In weapons
			t.wear array(weapon)
		Next
		
		'地圖
		map = addr("map", LV)
		t.map(map(0))
	Next	
End Function

Function addr(section, key)
	addr = split(dm.readini(section, key, ".\QMScript\train.ini"), ",")
End Function
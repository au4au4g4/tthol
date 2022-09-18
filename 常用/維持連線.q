[General]
SyntaxVersion=2
BeginHotkey=50
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=50
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

hwnds = t.getAllHwnds()
While True
	min = hour(now) * 60 + Minute(Now)
	If (min mod 1) = 0 Then 
		call disconnect()
	End If
	If (min mod 4) = 0 Then 
		call connect()
	End If
	If (min mod 60) = 0 Then 
		Call alarm()
		'Call train()
	End If
	If (min mod 12 * 60) = 0 Then 
		Call record()
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

Function alarm()
	str = ""
	For Each hwnd In hwnds
		t.init (hwnd)
		earn = t.monster - expp.item(hwnd)
		If earn = 0 Then 
			str = str + t.id + join(t.addr("teams")) + "/"
		End If
		expp(hwnd) = expp(hwnd) + earn
	Next
	u.pushMsg str
End Function

Function record()
	datas = array()
	For Each hwnd In hwnds
		t.init(hwnd)
		Redim Preserve datas(ubound(datas) + 1)
		datas(ubound(datas)) = array(t.id, t.level, t.place, t.period, t.monster, t.expp, t.money, t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6, t.getItemCnt("特貢令"))
	Next
	u.post "掛機", datas
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
			point = - 1 
			While (parm(0) - t.point(p) > 0) * (parm(1) - rnd() > 0) * (point <> t.point(p))
				point = t.point(p)
				t.addpoint (p)
			Wend
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
		t.useSkill(skill)
		
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
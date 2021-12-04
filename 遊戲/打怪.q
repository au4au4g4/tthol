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
MacroID=7da42200-306e-4671-830e-7215779ee589
Description=打怪
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
DimEnv mainID,disconnectedID,keepAliveID
//recordID = BeginThread(record)
mainID = BeginThread(main)
//disconnectedID = BeginThread(disconnected)
//keepAliveID = BeginThread(keepAlive)

Function main()
	Call initDM()
	POINTS = Array(Array(7,48),Array(19,41))
	pNum = 0
	MP_POT = &H5DE200
	BAG_ADDR = getBagInfo()
	While True
		If noTarget() Then 
			Call keepBuff(array(6,7))
			Call killMonster2()
			If noTarget() Then 
				Call patrol(POINTS)
				If need(MP_POT) Then 
					Call getWater()
				End If
			End If
			Delay 500
		End If
	Wend
End Function

Function need(item)
	Dim need
	need = false
	If not contain(MP_POT, getBag(&H360))  Then 
		If contain(MP_POT, getBag(&H368)) Then 
			need = true
		End If
	End If
End Function

Function kill()
	Call dm.WriteData(hwnd, "<tthola.dat>+4228B", "EB35")
	Call press(116)
	dm.rightclick 
	Call dm.WriteData(hwnd, "<tthola.dat>+4228B", "740A")
End Function

Function killMonster2()
	If hasTarget() Then 
		Exit Function
	End If
	Call dm.WriteData(hwnd, "<tthola.dat>+422BE", "8B4424048B4014C20C00")
	Call press(116)	// 攻擊技能鍵
	dm.rightclick 
	Call dm.WriteData(hwnd, "<tthola.dat>+422BE", "C20C0090909090909090")
End Function

Function patrol(points)
	If arrive(points(pNum)) Then 
		pNum = (pNum + 1) mod (UBound(points) + 1)
	End If
	If not walking() Then 
		call go(points(pNum))
	End If
End Function

Function go(xy)
	xy(0) = toRevHex(xy(0) * 40)
	xy(1) = toRevHex(xy(1) * 40)
	Call dm.WriteData(hwnd, "<tthola.dat>+5FF1D", "BF" & xy(1) & " BB" & xy(0) & " 90909090")
	Call dm.WriteData(hwnd, "<tthola.dat>+5FC88", "39")
	Call dm.WriteData(hwnd, "<tthola.dat>+5FB8D", "84")
	Delay 5
	Call dm.WriteData(hwnd, "<tthola.dat>+5FB8D", "85")
	Call dm.WriteData(hwnd, "<tthola.dat>+5FC88", "85")
	Call dm.WriteData(hwnd, "<tthola.dat>+5FF1D", "0F84ED0100008B7C244C8B5C2448")
End Function

Function arrive(point)
	Dim mx, my
	mx = read("[[[[<tthola.dat>+3EB99C]+124]+8]+AE4]+2C04")
	my = read("[[[[<tthola.dat>+3EB99C]+124]+8]+AE4]+2C08")
	point = array(point(0) * 40, point(1) * 40)
	arrive = distance(point, array(mx, my)) < 200
End Function

Function keepAlive()
	Call initDM()
	HP_LIMIT = 0.6
	MP_LIMIT = 0.3
	HP_MAX = read("[[[<tthola.dat>+123F38]+98]+68]+130")
	MP_MAX = read("[[[<tthola.dat>+123F38]+98]+6C]+130")
	While True
		hp = read("[[[<tthola.dat>+123F38]+98]+68]+140")/HP_MAX
		mp = read("[[[<tthola.dat>+123F38]+98]+6C]+140")/MP_MAX
		// HP
		If hp < HP_LIMIT Then 
			TracePrint "沒血"
			Call run()
		ElseIf hp < HP_LIMIT + 0.2 Then
			Call press(112)
		End If
		// MP
		If mp < MP_LIMIT Then 
			TracePrint "沒魔"
			Call run()
		ElseIf mp < MP_LIMIT + 0.2 Then
			Call press(113)
		End If
		Delay 100
	Wend
End Function

Function keepBuff(codes)
	usingBuff = getUsingBuff()
	For Each code In codes
		If not contain(getSkill(code), usingBuff)  Then 
			While skilling()
				Delay 100
			Wend
			Call press(111+code)
			dm.rightclick 
		End If
	Next
End Function

Function getUsingBuff()
	Dim headAddr,lenAddr
	lenAddr = "[[<TTHOLA.dat>+3E77D4]+10]+20F0"
	headAddr = read("[<TTHOLA.dat>+3E77D4]+10") + &H20F8
	getUsingBuff = getAddrs(headAddr,lenAddr)
End Function

Function getBag(owner)
	Dim lenAddr, headAddr,i
	Redim bag(0)
	i=0
	lenAddr = BAG_ADDR+owner-4
	headAddr = read(BAG_ADDR+owner)
	itemOffsets = getAddrs(headAddr, lenAddr)
	For Each itemOffset In itemOffsets
		Redim Preserve bag(i)
		bag(i) = read(itemOffset + 4)
		i=i+1
	Next
	getBag = bag
End Function

Function getBagInfo()
	Dim eax,esi
	eax = read("[[<tthola.dat>+003E77D4]+10]+AFC") and "&HFFFF"
	esi = read("[[[<tthola.dat>+003E77D4]+10]+C]+8")
	getBagInfo = read(esi+ eax*4)
End Function
 
Function getWater()
	Dim PET_XY,pet,xy
	PET_XY = Array(60, 196, 187, 323)
	pet = Array("B5B2FF","3|11|5224FF,13|3|000000")
	Call hotKey(18,73)
	Call hotKey(18, 80)
	Call lClick(array(78, 172))		// 開寵
	Delay 100
	Call lClick(array(83, 219))		// 選寵
	Call lClick(array(129, 317))	// 點水
	Call lClick(array(683,265))		// 放包
	Call press(13)
	Call lClick(array(78, 172))		// 開寵
	Delay 100
	Call lClick(array(83, 219))		// 關寵
	Call press(27)
	Delay 50
	Call press(27)
End Function

Function run()
	call press(114)
	EndScript
End Function

Function disconnected()
	call initDM()
	COMFIRM = Array("424542", "13|0|737573,55|19|009ECE,-5|22|DE9E84")
	PASSWORD = Array("FFFFFF", "88|32|CE4121,216|46|D68A00")
	GAME_PAGE = Array("EFA2A5", "556|577|000000")
	LOGIN_PAGE = "424539-000000"
	SERVER_XY = Array(613,222)
	CHARACTER_XY = Array(282,295)
	START_XY = Array(106,442)
	While true
		xy = findIt(COMFIRM)
		If xy(0) > 0 Then 
		 	StopThread mainID
			// 連線中斷
			Call lClick(xy)
			xy = untilFindIt(PASSWORD)
			enter = Array(xy(0), xy(1) + 30)
			
			// 登入頁
			Call lClick(SERVER_XY) 	//選伺服器
			Call lClick(xy)				//密碼框
			dm.SendString hwnd,"gj83dj4"
			
			// 登入
			While dm.CmpColor(0, 0, LOGIN_PAGE, 1) = 0
				Call lClick(enter)
				Delay 3000
				xy = untilFindIt(COMFIRM)
				Call lClick(xy)
			Wend
			// 角色頁
			Call lClick(CHARACTER_XY)	//選角
			Call lClick(START_XY)	// 開始遊戲
			Call untilFindIt(GAME_PAGE)
			mainID = BeginThread(main)
		End If
	Wend
End Function

// ------------------------------GAME------------------------------
Function noTarget()
	edx = read("[<tthola.dat>+251EB0]+8")
	eax = read("[[<tthola.dat>+003E77D4]+10]+AFC") and "&HFFFF"
	noTarget = read(array(edx + eax * 4, "10C8"))<>3
End Function

Function hasTarget()
	edx = read("[<tthola.dat>+251EB0]+8")
	eax = read("[[<tthola.dat>+003E77D4]+10]+AFC") and "&HFFFF"
	hasTarget = read(array(edx + eax * 4, "10C8"))=3
End Function

Function skilling()
	free = read("[<tthola.dat>+2230EC]+44")<>0
End Function

Function walking()
	edx = read("[<tthola.dat>+251EB0]+8")
	eax = read("[[<tthola.dat>+003E77D4]+10]+AFC") and "&HFFFF"
	walking = read(array(edx+eax*4,"A0"))=1
End Function

Function getSkill(code)
	getSkill = read(array("[<tthola.dat>+3E77D4]+10",&H3044+(code-1)*8))
End Function

Function swPortal()
	edx = read("<tthola.dat>+438CD")
End Function

// ------------------------------抓色------------------------------
Function findIt(colorData)
	Call dm.FindMultiColor(0, 0, 800, 600, colorData(0), colorData(1), 1, 1, intX, intY)
	findIt = Array(intX,intY)
End Function

Function untilFindIt(colorData)
	Dim xy
	xy = Array(-1,-1)
	While xy(0)<0
		xy = findIt(colorData) 
	Wend
	untilFindIt = xy
End Function
// ------------------------------鍵鼠------------------------------
Function lClick(xy)
	dm.moveto xy(0), xy(1)
	Delay 10
	dm.leftclick 
	Delay 100
End Function

Function rClick(xy)
	dm.moveto xy(0), xy(1)
	Delay 10
	dm.rightclick 
	Delay 100
End Function

Function press(a)
	dm.KeyPress a
End Function

Function hotKey(a,b)
	dm.KeyDown a
	dm.KeyPress b
	dm.KeyUp a
End Function

// ------------------------------記憶體------------------------------
Function read(addrs)
	read = dm.ReadInt(hwnd, getScript(addrs), 0)
End Function

Function getScript(addrs)
	Dim script,addr
	If isArray(addrs) Then 
		For Each addr In addrs
			addr = toHEX(addr)
			If isEmpty(script) Then 
				script = addr
			Else 
				script = "["&script&"]+"&addr
			End If
		Next
	Else 
		script = toHEX(addrs)
	End If
	getScript = script
End Function

Function toHEX(data)
	If TypeName(data) <> "String" Then 
		data = HEX(data)
	End If
	toHEX = data
End Function

Function toRevHex(num)
	dim str
	str = hex(num)
	str = String(8 - len(str), "0") & str
	str = Mid(str, 7, 2) & Mid(str, 5, 2) & Mid(str, 3, 2) & Mid(str, 1, 2)
	toRevHex = str
End Function

Function getAddrs(headAddr,lenAddr)
	Dim i,len
	Redim result(0)
	len = read(lenAddr)
	For i = 0 To len - 1
		Redim Preserve result(i)
		 result(i) = read(headAddr+4*i)
	Next
	getAddrs = result
End Function

// ------------------------------其他------------------------------
Function initDM()
	Set dm = createobject("dm.dmsoft")
	hwnd = dm.FindWindow("", "Tthol")
	dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
	dm.moveto 800,0
	dm.SetKeypadDelay "windows", 40
	dm.SetMouseDelay "windows", 40
End Function

Function distance(sXY,dXY)
	distance = abs(sXY(0)-dXY(0))+abs(sXY(1)-dXY(1))
End Function

Function contain(item,arr)
	contain = UBound(filter(arr, item)) > -1
End Function

Function record()
	While true
		KeyDown 91, 1
		KeyDown 18, 1
		KeyPress 82, 1
		KeyUp 18, 1
		KeyUp 91, 1
		Delay (30*60+10)*1000
	Wend
End Function

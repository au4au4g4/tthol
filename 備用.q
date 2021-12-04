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
MacroID=bc5ed7ca-d1df-4e01-8df8-25ebbd25114f
Description=備用
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
// 打怪範圍
RANGES = Array(Array(50, 50,400,300), Array(400, 50,750,300),Array(50, 300,400,470),Array(400, 300,750,470))
RANGES1 = getRange(Array(Array(250, 250), Array(650, 350)))

 //怪物類型
ALL = Array(Array(20,50),"-1|1|FFB600,-1|2|FFB600,-2|3|FFB600,-2|4|FFB600,-3|5|FFB600,-3|6|FFB600,-3|7|FFB600,-2|8|FFB600,-1|8|FFB600,0|8|FFB600,1|8|FFB600,2|8|FFB600,3|7|FFB600,3|6|FFB600,3|5|FFB600,2|4|FFB600,2|3|FFB600,1|2|FFB600,1|1|FFB600")
TRIANGLE = Array(Array(30,50),"-1|2|FFB600,-2|4|FFB600,-3|6|FFB600,-4|8|FFB600,-3|8|FFB600,-1|8|FFB600,1|8|FFB600,4|8|FFB600,3|6|FFB600,2|4|FFB600,1|2|FFB600")
CIRCLE = Array(Array(30,50),"-3|1|FFB600,-4|4|FFB600,-3|7|FFB600,0|8|FFB600,3|7|FFB600,4|4|FFB600,3|1|FFB600")

Function findMonster(mDatas)
	Dim mData,xy,range,offset,count,i,size
	KeyPress "F5", 1
	i = 0
	size = UBound(RANGES)+1
	count = size
	While count > 0
		For j = 0 To UBound(mDatas)
			range = RANGES(i)
			mData = mDatas(j)
			offset = mData(0)
			xy = monsterXY(range, mData)
			If xy(0) > 0 Then 
				Call rClick(xy(0) + offset(0), xy(1) + offset(1))
				count = size - 1
				Delay 500
			Else 
				count = count - 1
			End If
		Next
		i=(i+1)mod (size)
	Wend
End Function

Function monsterXY(range,mData)
	dim offset
	offset = mData(0)
	monsterXY = findIt(range(0)-offset(0),range(1)-offset(1),range(2)-offset(0),range(3)-offset(1) , "FFB600", mData(1))
End Function

Function useBuff2(buffs)
	Dim buff
	If isempty(schedule)  Then 
		schedule = Array(0, 0, 0, 0, 0)	
	End If
	For i = 0 To UBound(buffs)
		buff = buffs(i)
		If  schedule(i)=0 or now - schedule(i)>0 Then 
			KeyPress buff(0), 1
			Call rClick(400, 300)
			schedule(i) = now + buff(1) / 24 / 60 / 60
			Call TracePrintArray(schedule)
			Delay 2500
		End If
	Next
End Function

Function TracePrintArray(array)
	Dim i
	For i = 0 To UBound(array)
		TracePrint "array"&i&"："&array(i)
	Next
End Function

//------------------------------------------------------------------
WATERFALL_M = Array("5A2429", "B5B2B5")
PEACE_M = Array("6BC300", "31B6DE", "733029", "94AA4A")
CLIFF_M = Array("7B5510")
Function killByColor(colors)
	Dim xy,count
	KeyPress "F5", 1
	count = 1
	While count > 0
		count = 0
		For i = 0 To UBound(colors)
			xy = findColorXY(100, 0, 700, 550, colors(i))
			If xy(0) > 0 Then 
				Call rClick(xy(0) - 1, xy(1))
				count = count + 1
				call atkDone()
			End If
		Next
	Wend
End Function

//------------------------------------------------------------------

WATERFALL1_C = Array("29717B","3|0|31496B,8|0|F7C763")
WATERFALL1_P = Array( Array(69,73), Array(-23,77 ), Array(-23,48 ),  Array(-20,18 ),Array(64,11),Array(52,51))
WATERFALL2_C = Array("31617B","-5|0|4A71A5,-11|0|39717B,6|0|EFBE39")
WATERFALL2_P = Array(Array(25, - 70 ), Array(63, - 70 ), Array(65, - 25 ), Array(- 45 , - 25 ), Array(- 36 , 5), Array(- 45 , - 25 ), Array(65, - 25 ), Array(63, - 70 ))
CLIFF_C= Array("E7D7BD","0|-3|E7DBCE,0|-5|394D73,-2|0|E7D7BD,-2|-1|E7DBCE,-2|-4|ADC7AD")
CLIFF_P = Array(Array(- 55 , - 83 ), Array(- 79 , - 80 ), Array(- 106 , - 61 ), Array(- 81 , - 55 ), Array(- 37 , - 46 ), Array(17, - 52 ), Array(- 7 , - 79 ))
PEACE_C = Array("FFEFE7","1|0|F7E3DE,2|0|FFEFEF,7|0|EFEFEF,0|0|FFEFE7,1|0|F7E3DE,2|0|FFEFEF")
PEACE_P = Array(Array(78,-78 ),Array(114,-82 ),Array(56,-60),Array(16,-64 ),Array(-16,-84 ),Array(34,-79 ))
WATERFALL1 = Array(WATERFALL1_C, WATERFALL1_P, WATERFALL_M)
WATERFALL2 = Array(WATERFALL2_C, WATERFALL2_P, WATERFALL_M)
CLIFF = Array(CLIFF_C, CLIFF_P)
PEACE = Array(PEACE_C,PEACE_P)

Function patrol2(center,points)
	Dim cXY, pXY,MAP_XY
	MAP_XY = Array(655, 493, 789, 594)
	MoveTo L, T
	Delay 100
	cXY = findIt(MAP_XY, center)
	pXY = getPXY(cXY, points)
	If cXY(0) < 0 Then 
		TracePrint "找不到中心"
		Call run()
	End If
	While outside(pXY, MAP_XY) or checkColor(pXY(0), pXY(1), "1875FF",2)
		Call nextPoint(points)
		pXY = getPXY(cXY, points)
	Wend
	call lClick(pXY(0), pXY(1))
End Function

pointNo= 0
Function getPXY(cXY,points)
	Dim x,y
	x = cXY(0) + points(pointNo)(0)
	y = cXY(1) + points(pointNo)(1)
	getPXY = Array(x,y)
End Function

Function outside(xy, range)
	outside = xy(0)< range(0) or xy(0)>range(2) or xy(1)<range(1) or xy(1)>range(3)
End Function

Function nextPoint(points)
	pointNo = (pointNo + 1) mod (UBound(points)+1)
End Function

//------------------------------------------------------------------

Function useBuff(buffs)
	Dim buff
	For i = 0 To UBound(buffs)
		buff = buffs(i)
		If no(buff) Then 
			KeyPress buff(0), 1
			call rClick(400,300)
			Delay 1000
		End If
	Next
End Function

Function no(buff)
	Dim BUFF_XY
	BUFF_XY = Array(45, 450, 300, 480)
	no = findIt(BUFF_XY,buff(2))(0)< 0
End Function

//------------------------------------------------------------------

Function canAtk(x,y)
	MoveTo L + x, T + y
	Delay 50
	canAtk = getRam(Array(tthola+&H45620,0))<>0
End Function

//------------------------------------------------------------------

Function aktDone()
	Dim val
	val = 1
	While val <> 0
		val = getRam4(Array(TTHOLA+&H3E77D4,&H10,&H20DC))
	Wend	
End Function

//------------------------------------------------------------------

Function record2()
	Dim count
	While true
		count=0
		While  count<1
			IfColor 2324, 160, "0000FF", 1 Then
				KeyDown 91, 1
				KeyDown 18, 1
				KeyPress 82, 1
				KeyUp 18, 1
				KeyUp 91, 1
				count = count + 1
				Delay 10000
			End If
		Wend
		Delay 4*60*60*1000
	Wend
End Function

//------------------------------------------------------------------

AKT_ZONE = Array(0, 0, 800, 440)
Function getMonsterXY()
	dim ms_XY, mos_XY,carriage,xy
	carriage = head
	Do while  carriage > 0 and carriage < 3722304989
		If getRam4(Array(carriage + &H10))=&H1010000 Then 
			head = carriage
		ElseIf isMonster(carriage) Then
			ms_XY = getRam4Pair(Array(XY_DATA + &H2C0C))
			mos_XY = getRam2Pair(array(carriage + &H26))
			xy = array(mos_XY(0) - ms_XY(0), H - ms_XY(1) - mos_XY(1)-20)	
			If inside(xy, AKT_ZONE) Then 
				getMonsterXY = xy
				Exit Function
			End If
		End If
		carriage = getRam4(Array(carriage + &H124))
	Loop
	getMonsterXY = Array(- 1 , - 1 )
End Function

Function isMonster(addr)
	dim mType,alive
	mType = getRam4(Array(addr + &H14))
	alive = getRam4(Array(addr + &H4C))
	isMonster = (mType = &H20000001) and alive = &H10
End Function

Function inside(xy, zone)
	inside = xy(0) > zone(0) and xy(0) < zone(2) and xy(1) > zone(1) and xy(1) < zone(3)
End Function

//------------------------------------------------------------------

H = read(XY_DATA&"+2AF4")
W = read(XY_DATA&"+2AF0")
Function go(xy)
	dim cx,cy,ox,oy,my_XY,MAP_W,MAP_H
	my_XY = getRam4Pair(Array(XY_DATA + &H2C04))
	MAP_W = 1340
	MAP_H = 990
	cx = limit(my_XY(0), W - MAP_W, MAP_W)	
	cy = limit(my_XY(1), H - MAP_H, MAP_H)	
	ox = limit(xy(0) * 40 - cx, MAP_W, - MAP_W ) / 20	
	oy = limit(cy - xy(1) * 40, MAP_H, - MAP_H ) / 20	
	Call lClick(Array(721 + int(ox), 544 + int(oy)))
End Function

Function arrive(point)
	Dim my_XY
	my_XY = getRam4Pair(Array(XY_DATA + &H2C04))
	point = array(point(0)*40,point(1)*40)
	arrive = distance(point,my_XY)<10000
End Function

Function limit(num, max, min)
	If num < min Then 
		limit = min
	ElseIf num > max Then
		limit = max
	Else 
		limit = num
	End If
End Function
//------------------------------------------------------------------
head = read("[<tthola.dat>+251EB0]+14")
Function getMonsterAddr()
	dim carriage
	carriage = head
	Do While carriage > 0
		If read(carriage+16)=&H1010000 Then 
			head = carriage
		ElseIf isMonster(carriage) Then
			getMonsterAddr = carriage
			Exit Function
		End If
		carriage = read(carriage+&H124)
	Loop
	getMonsterAddr = 0
End Function

Function isMonster(addr)
	dim mType,alive,mID
	mType = read(addr + &H14)
	mID = read(addr + &HC)
	alive = read(addr+&H4C)
	isMonster = (mType = &H20000001) and (alive = &H10)
End Function
//------------------------------------------------------------------
//------------------------------------------------------------------

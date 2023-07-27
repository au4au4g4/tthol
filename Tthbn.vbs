Class Tthbn

	Private dm,dw,re,hwnd,ttha,tthbn,tthsj,addrs,timegettime
	
	Public Sub Class_Initialize()
		Set dm = createobject("dm.dmsoft")
		Set dw = CreateObject("DynamicWrapper")
		Set re = New RegExp
		dw.Register "kernel32.dll", "OpenProcess", "i=uuu", "r=h" 
		dw.Register "kernel32.dll", "VirtualAllocEx", "i=lllll", "r=l"
		dw.Register "WINMM.DLL", "timeGetTime", "r=l", "f=s"
		timegettime = dw.timeGetTime()
    End Sub
	
   	Public default function Init(p_hwnd)
		hwnd = p_hwnd
		tthsj = dm.GetModuleBaseAddr(hwnd, "tthsj.bin")
		ttha = dm.GetModuleBaseAddr(hwnd, "ttha.bin")
		tthbn = dm.GetModuleBaseAddr(hwnd, "tthbn.bin")
		Set addrs = CreateObject("Scripting.Dictionary")
   	End function

	'flag
	'已斷開
	Public Function isDisConnect()
		isDisConnect = (dm.ReadInt(hwnd, addr("isDisConnect",0), 0) = 0)
	End Function
	Public Function isStart()
		isStart = (dm.ReadInt(hwnd, "[<tthbn.bin>+C520]", 0) = 1)
	End Function
	'已斷線
	Public Function isOffLine()
		isOffLine = (dm.ReadString(hwnd, addr("isOffLine",0), 0, 16)="")
	End Function
	
	'read
	Public Function id()
		id = memberST(account)(1)
	End Function
	Public Function account()
		account = dm.ReadString(hwnd, addr("account",0), 0, 16)
	End Function
	Public Function level()
		level = dm.ReadInt(hwnd, addr("level",0), 1)
	End Function
	Public Function money()
		money = dm.ReadInt(hwnd, "<tthbn.bin>+ACD24", 0)
	End Function	
	Public Function expp()
		expp = dm.ReadInt(hwnd, "<tthbn.bin>+109F10", 0)
	End Function
	Public Function x()
		x = dm.ReadInt(hwnd, "[[[<ttha.bin>+004EF82C]+98]+2F8]+48", 0)
	End Function
	Public Function locationNo()
		locationNo = dm.ReadInt(hwnd, "[[<ttha.bin>+004EF82C]+98]+164", 0)
	End Function
	Public Function location()
		location = dm.ReadString(hwnd, "<tthbn.bin>+1099E4", 0, 16)
	End Function
	Public Function place()
		place = memberST(account)(2)
	End Function
	Public Function period()
		dim start
		start = dm.ReadInt(hwnd, "<tthbn.bin>+ACD20", 0)
		period = dw.timeGetTime() - start
		if period <= 0 then
			period = period + 2^32
		end if
		period = period / 60 / 60 / 1000
	End Function
	Public Function monster()
		cnt = 1 : i = 0
		While cnt > 0
			cnt = dm.ReadInt(hwnd, addr("monster", 28 * i), 0)
			monster = monster + cnt
			i=i+1
		Wend
	End Function
	Public Function cash ()
		cash = dm.ReadInt(hwnd, addr("cash",0), 0) 
	End Function
	Public Function deposit()
		deposit = dm.ReadInt(hwnd, "[<tthbn.bin>+E5FC]", 0)
	End Function
	Public Function team()
		team = dm.ReadIni("team", id, ".\tthbn.ini")
	End Function
	Public Function nRange()
		nRange = dm.ReadIni(team, "nRange", ".\tthbn.ini")	
	End Function
	Public Function gRange()
		gRange = dm.ReadIni(team, "gRange", ".\tthbn.ini")	
	End Function
	Public Function skillLv(no,name)
		dim i,learnedSkill
		i=0
		learnedSkill = dm.readint(hwnd,HEX(tthbn+&H1082A0+36*i),0)
		while (learnedSkill<>0) * (no<>cstr(learnedSkill))
			i=i+1
			learnedSkill = dm.readint(hwnd,HEX(tthbn+&H1082A0+36*i),0)
		Wend
		skillLv = dm.readint(hwnd,HEX(tthbn+&H1082A0+36*i+20),0)
	End Function
	
	Public Function getMsg()
		getMsg = dm.ReadString(hwnd, addr("getMsg",0), 0, 115)
	End Function
	
	Public Function getMonster()
		getMonster = getRecords(&HAEBDC)
	End Function
	
	Public Function getReward()
		getReward = getRecords(&HACD3C)
	End Function
	
	private Function getRecords(offset)
		Dim keys
		keys = array(array("name", 8), array("id", 0), array("cnt", 4))
		getRecords = getObjs(HEX(tthbn + offset), null, 28, array(""), keys)
	End Function
	
	Public Function findNPC(name)
		Dim keys : keys = array(array("name", 12),array("id", 4), array("sn", 0))
		set findNPC = getObjs(addr("npc",0), addr("npcCnt",0), 32, array(name), keys)(0)
	End Function
	
	Public Function getItemCnt(name)
		dim item,items
		getItemCnt = 0
		items = getBag(array(name))
		for each item in items
			getItemCnt = getItemCnt + item.item("cnt")
		Next
	End Function
	
	Public Function getBag(names)
		getBag = getItems(addr("bag",0), addr("bagCnt",0), names)
	End Function
	
	Public Function getBank(names)
		getBank = getItems(addr("bank",0), addr("bankCnt",0), names)
	End Function
	
	private Function getItems(addr, cntAddr, names)
		Dim keys
		keys = array(array("name", 32),array("sn", 0), array("id", 4), array("cnt", 8))
		getItems = getObjs(addr, cntAddr, 280, names, keys)
	End Function
	
	'function
	Public Function talkNPC()
		call dm.WriteInt(hwnd,"[[<ttha.bin>+004EF82C]+98]+150",0,&HE7) 'resetTalk
		simpleCall hwnd, ttha + &H44D90, array()
	End Function
	 
	'1開始 2對話 45678選項
	Public Function talkOption(npc, arr)
		For Each a In arr
			If a = 0 Then 
				Delay 500
			Else 
				simpleCall hwnd, tthbn + &H24720, array(1, npc, a)
				Delay 50
			End If
		Next
	End Function
	
	Public Function start()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,[" + HEX(ttha + &H4EF82C) + "]"
		dm.AsmAdd "mov ecx,[ecx+98]"
		dm.AsmAdd "call "+ HEX(ttha+&H54EE0)
		dm.AsmCall hwnd,1
	End Function
	
	Public Function stops()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,[" + HEX(ttha + &H4EF82C) + "]"
		dm.AsmAdd "mov ecx,[ecx+98]"
		dm.AsmAdd "call "+ HEX(ttha+&H54E80)
		dm.AsmCall hwnd,1
	End Function
	
	Public Function learn(code)
		For i = code(1) to code(2)
			simpleCall hwnd, tthbn + &H2BA10, array(i,code(0))
		Next
		delay 2000
	End Function
	
	' 蒐集物品
	Public Function trade(bHwnd, keywords, cntt)	
		dim sID,bID,sn,iID,bTthbn,bag,tableCnt,bName,items,cnt
		cnt = cntt
		temp = hwnd
		Init(bHwnd)
		bag = getBag(array(".*"))
		Init(temp)
		items = getBag(keywords)

		if hwnd - bHwnd = 0 or UBound(bag) = 39 or UBound(items) = -1 then
			exit Function
		end if

		' 開啟交易
		bName = dm.ReadString(bHwnd, addrStr("name",0), 0, 20)
		bID = getIdByName(hwnd, bName)
		sID = getIdByName(bHwnd, id)
		bTthbn = dm.GetModuleBaseAddr(bHwnd, "tthbn.bin")
		simpleCall hwnd, addr("invite", 0), array(bID, &HEA64)	' 邀請	
		simpleCall bHwnd, addr("accept", bTthbn - tthbn), array(sID, &HEA64)' 接受

		'放物品
		tableCnt = 0
		For i = 0 To UBound(items)
			If cnt = 0  or tableCnt = 10 or tableCnt + UBound(bag) = 39 Then
				Exit For
			End If
			
			Set item = items(i)
			sn = item.item("sn")
			iID = item.item("id")
			itemCnt = item.item("cnt") 
			
			if itemCnt - cnt >= 0 then
				itemCnt = cnt
			end if

			simpleCall hwnd, addr("give", 0), array(itemCnt, sn, iID)	' 交付
			cnt = cnt - itemCnt
			tableCnt = tableCnt + 1			
		next

		' 完成交易
		simpleCall hwnd, addr("comfirm1", 0), array()			' 確定1
		simpleCall bHwnd, addr("comfirm1", bTthbn - tthbn), array()
		simpleCall hwnd, addr("comfirm2", 0), array()			' 確定2
		simpleCall bHwnd, addr("comfirm2", bTthbn - tthbn), array()
		simpleCall bHwnd, addr("comfirm2", bTthbn - tthbn), array()
	End Function
	
	Public Function getIdByName(hwndd, name)
		tthbnn = dm.GetModuleBaseAddr(hwndd, "tthbn.bin")
		dim addrs : addrs = dm.FindString(hwndd,addr("getIdByName",tthbnn-tthbn)&"-FFFFFFFF",name,0)
		getIdByName = dm.readInt(hwndd,HEX(clng("&H"&split(addrs,"|")(0))-16),0)
	End Function	
	
	Public Function openShop()
		simpleCall hwnd, tthbn + &H29C40, array()
	End Function

	Public Function closeShop()
		simpleCall hwnd, tthbn + &H2AFE0, array()
	End Function
	
	Public Function addItem(price, cnt, no)
		simpleCall hwnd, tthbn + &H2ABA0, array(1, price, cnt, no)
	End Function
	
	Public Function withdrawal(amount)	
		openBank()
		simpleCall hwnd, tthbn + &H26000, array(0,amount)		
	End Function
	
	Public Function save(amount)
		openBank()
		simpleCall hwnd, tthbn + &H26000, array(1,amount)
	End Function
	
	Public Function pop(names)
		dim items, item
		openBank()
		items = getBank(names)
		for each item in items
			simpleCall hwnd, addr("pop",-106), array(0,item.item("cnt"),item.item("sn"),item.item("id"))
		Next
	End Function
	
	Public Function push(names)
		dim items, item
		openBank()
		items = getBag(names)
		for each item in items
			simpleCall hwnd, addr("pop",-106), array(1,item.item("cnt"),item.item("sn"),item.item("id"))
		Next
	End Function
	
	Public Function openBank()
		dim npc : Set npc = findNPC("伙計")
		simpleCall hwnd, addr("openBank",0), array(npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function buy(npcName, itemName, cnt)
		dim npc : Set npc = findNPC(npcName)
		dim itemId : itemId = itemCode(itemName)(0)
		simpleCall hwnd, addr("buy",0), array(cnt, itemId, npc.item("sn"), npc.item("id"))
		delay 1000
	End Function
	
	Public Function sell(npcName, itemName)
		dim npc : Set npc = findNPC(npcName)
		dim items :items = getBag(array(itemName))
		for each item in items
			simpleCall hwnd, addr("sell",0), array(item.item("cnt"),item.item("sn"),item.item("id"),npc.item("sn"),npc.item("id"))
			delay 300
		Next
	End Function
	
	Public Function kill(x,y,sn,id,skill)
		simpleCall hwnd, tthbn + &H23900, array(x,y,sn,id,skill)
	End Function

	Public Function login()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,0" + HEX(ttha + &H4F27CC)
		dm.AsmAdd "mov ecx,[ecx]"
		dm.AsmAdd "call 0" + addr("login",-29)
		dm.AsmCall hwnd, 1	
	End Function	
	
	Public Function groupAtk(flag)
		call dm.WriteInt(hwnd, "<ttha.bin>+53ADB", 2, flag)
	End Function
	
	Public Function skill(code)
		dim base,x,y,sn,id
		base = dm.ReadInt(hwnd, "[[<ttha.bin>+43B90]+20]+98", 0)
		y = dm.ReadInt(hwnd, HEX(base+&H1B6), 0)
		x = dm.ReadInt(hwnd, HEX(base+&H1B4), 0)
		sn = dm.ReadInt(hwnd, HEX(base+&H184), 0)
		id = dm.ReadInt(hwnd, HEX(base+&H180), 0)
		simpleCall hwnd, tthbn + &H23900, array(y,x,sn,id,code)
	End Function
	
	Public Function useSkill(code)
		call dm.WriteInt(hwnd, "[[[<ttha.bin>+4EF82C]+98]+11E90]", 0, code(0))
		dim src,name,big5
		src = "../JD_FCS21.60/config/"
		name = t.id & "_飛雁山莊(花)"
		big5 = replace(dm.StringToData(name, 0), " ", "")
		if dm.IsFileExist(src & big5 & ".dat") then
			src = src & big5 & ".dat"
		else
			src = src & name & ".dat"
		end if
		dm.writeini "hit","loopskill1",code(1),src 
	End Function
	
	Public Function point(p)
		Select Case p
			Case "w"
				offset=0
			Case "n"
				offset=1
			Case "gg"
				offset=2
			Case "s"
				offset=3
			Case "g"
				offset=4
			Case "ss"
				offset=5
		End Select
		point = dm.ReadInt(hwnd, HEX(tthbn+&H109F1A+offset*2), 1)
	End Function
	
	Public Function addPoint(p)
		Select Case p
			Case "w"
				offset=0
			Case "n"
				offset=1
			Case "gg"
				offset=2
			Case "s"
				offset=3
			Case "g"
				offset=4
			Case "ss"
				offset=5
		End Select
		simpleCall hwnd, ttha + &H3EBB0 + offset*16, array()
		delay 2000
	End Function
	
	Public Function wear(names)
		dim weapon,id,sn
		set weapon = t.getBag(names)(0)
		id = weapon.item("id")
		sn = weapon.item("sn")
		simpleCall hwnd, tthbn + &H26AB0, array(sn,id)
	End Function
	
	Public function removePlayer()
		Call dm.WriteData(hwnd, "<ttha.bin>+66EAD", "909090909090")
		reDim codes(4)
		codes(0) = "push edx"
		codes(1) = "push 0000013C"
		codes(2) = "call 0" + HEX(ttha + &H7E764)
		codes(3) = "call 0" + HEX(tthbn + &H26180)
		codes(4) = "jmp 0" + HEX(ttha + &H67396)
		Call inAsm("ttha.bin+6738B", codes)
	End Function
	
	Public function updateItem(itemAction)
		action = split(dm.ReadIni("action", itemAction(1), ".\QMScript\tthbn.ini"),",")	
		arr = itemCode(itemAction(0))
		For i = 0 To UBound(arr)
			simpleCall hwnd, tthbn + 57568, array(action(1), action(0), arr(i))
		Next
	End Function
	Function itemCode(name)
		Dim itemStart,itemAddr,arr
		itemStart = dm.FindData(hwnd, "00000000-FFFFFFFF", "20 4E 00 00 1E 00 00 00 00 00 00 00 BB C8 A8 E2")
		itemAddr = dm.FindData(hwnd, itemStart + "-" + hex(clng("&H" + itemStart) + 4000000), "00 00" + big5(name) + "00 00 00 00")
		arr = split(itemAddr, "|")
		For i = 0 To UBound(arr)
			arr(i) = dm.readint(hwnd, hex(clng("&H" + arr(i)) - 10), 0)
		Next
		itemCode = arr
	End Function
	Function big5(word)
		dim i
		For i = 0 To len(word) - 1
			big5 = big5 & HEX(asc(right(word,len(word)-i)))
		Next
	End Function
	
	Public function reset()
		simpleCall hwnd, addr("reset",-70), array()
	End Function
	
	Public function apply(name)
		dim item
		set item = getBag(array(name))(0)
		simpleCall hwnd, addr("apply",0), array(-01,item.item("sn"),item.item("id"))
		delay 500
	End Function
	
	Public function map(mapID)
		dim src,name,big5
		src = "../JD_FCS21.60/config/"
		name = t.id & "_飛雁山莊(花)"
		big5 = replace(dm.StringToData(name, 0), " ", "")
		if dm.IsFileExist(src & big5 & ".dbt") then
			src = src & big5 & ".dbt"
		else
			src = src & name & ".dbt"
		end if
		dm.DeleteFile src
		dm.WriteFile src, "1,1," & mapID
		Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+10C", 0, 0)
		Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+11C", 0, mapID)
		Call dm.WriteInt(hwnd, "[[[<ttha.bin>+4EF82C]+98]+F4]", 0, mapID)	
	End Function
	
	Public function frame(arr)
		dim x,y,w,h
		x = arr(0)
		y = arr(1)
		w = arr(2)
		h = arr(3)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+78", 0, x)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+7C", 0, y)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+80", 0, 0+x+w)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+84", 0, 0+y+h)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+1E8", 0, x)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+1EC", 0, y)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+1F0", 0, 0+x+w)
		Call dm.WriteInt(hwnd, "[[[[<ttha.bin>+D80DC]+414]+108]+694]+1F4", 0, 0+y+h)
	End Function
	
	'hwnd
	Public Function getAllHwnds()
		getAllHwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1+4), ",")
	End Function
	
	Public Function getPartHwnds()
		Dim w, h, col, row, x, y, nx, ny, hwnd,result
		w = dm.getscreenwidth()
		h = dm.getscreenheight() - 40
		col = 5 : row = 2
		result = ""
		Call GetCursorPos(x, y)
		For i = 0 To col-1
			For j = 0 To row - 1
				nx = (x + w * i / col) mod w
				ny = (y + h * j / row) mod h
				hwnd = dm.GetPointWindow(nx, ny)
				If instr(dm.GetWindowProcessPath(hwnd), "ttha") <> 0 Then 
					result = result & " " & hwnd
				End If
			Next
		Next
		getPartHwnds = split(trim(result)," ")
	End Function
	
	Public Function getArrHwnds(names)
		Dim hwndstr,hwnds
		hwndstr = ""
		hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1+4), ",")
		For Each hwnd In hwnds
			id = dm.ReadString(hwnd, "<tthbn.bin>+109720", 0, 16)
			If Instr(names, id) > 0 Then 
				hwndstr = hwndstr & hwnd & " "
			End If 
		Next
		getArrHwnds = split(trim(hwndstr), " ")
	End Function

	Public Function getHwnd()
		getHwnd = dm.FindWindow("", "絕代方程式")
	End Function
	
	'破解
	Public Function crackAll(atk,move,range)
		crack()
		moveSpeed atk,move
		atkRange(range)
		fixBank()
		if range=7 then
			fullSupport()
		end if
	End Function
	
	Public Function crack()
		Call dm.WriteData(hwnd, "<tthbn.bin>+ACD08", "E1E6F4B784ACEDB5E0E2A09CA9FBECE4B68BB10")
		Call dm.WriteInt(hwnd, "<tthbn.bin>+ACD20", 0, timegettime)
		Call dm.WriteData(hwnd,addr("crack",0),"EB")
	End Function

	'280
	Public Function moveSpeed(atk,move)
		Call dm.WriteInt(hwnd,addr("moveSpeedAtk",0),0,atk)
		Call dm.WriteInt(hwnd,addr("moveSpeed",0),0,move)
	End Function

	Public Function atkRange(range)
		reDim codes(5)
		codes(0) = "mov eax,0" & HEX(range ^ 2)
		codes(1) = "cmp edx,eax"
		codes(2) = "jg 0" + addr("atkRange",30)
		codes(3) = "cmp ecx,eax"
		codes(4) = "jg 0" + addr("atkRange",30)
		codes(5) = "jmp 0" + addr("atkRange",35)
		Call asm(addr("atkRange",0) , codes)
	End Function
	
	Public Function reLoginMin(min)
		Randomize
		min = Int(6 * Rnd() + 2)
		call dm.WriteInt(hwnd, "<ttha.bin>+3C334", 0, min * 60000)
	End Function

	Public function fixBank()
		Call dm.WriteData(hwnd, addr("distance",0), "10")	'距離
		
		reDim codes(8)
		codes(0) = "mov ecx,[esp+20]"
		codes(1) = "push 0"
		codes(2) = "push ecx"
		codes(3) = "lea ecx,[esp+20]"
		codes(4) = "call 0" + addr("fixbank",12 + dm.ReadInt(hwnd, addr("fixbank",8), 0))
		codes(5) = "call 0" + addr("fixbank",17 + dm.ReadInt(hwnd, addr("fixbank",13), 0))
		codes(6) = "test eax,eax"
		codes(7) = "je 0" + addr("fixbank",30)
		codes(8) = "jmp 0" + addr("fixbank",21)
		Call inAsm(addr("fixbank",12), codes)
		
		result = findAddr("81 FA E8 18")
		reDim codes(10)
		codes(0) = "cmp edx,000018E8"
		codes(1) = "je 0" + HEX(result+8)
		codes(2) = "cmp edx,00001851"
		codes(3) = "jne 0" + HEX(result+50)
		codes(4) = "mov edx,[ebp+0C]"
		codes(5) = "push edx"
		codes(6) = "mov ax,[ebp+08]"
		codes(7) = "push eax"
		codes(8) = "push 05"
		codes(9) = "call 0" + HEX(dm.ReadInt(hwnd, HEX(result + 34), 0) + result + 34 + 4)
		codes(10) = "jmp 0" + HEX(result+50)
		Call inAsm(HEX(result), codes)
		
		reDim codes(12)
		codes(0) = "push ecx"
		codes(1) = "call 0" + addr("fixbank2",dm.ReadInt(hwnd, addr("fixbank2",-4), 0))
		codes(2) = "mov ecx,[ebp+08]"
		codes(3) = "and ecx,0000FFFF"
		codes(4) = "cmp ecx,00001851"
		codes(5) = "jne 0" + addr("fixbank2",0)
		codes(6) = "mov eax,[ebp+0C]"
		codes(7) = "push eax"
		codes(8) = "mov cx,[ebp+08]"
		codes(9) = "push ecx"
		codes(10) = "push 05"
		codes(11) = "call 0" + addr("fixbank2",dm.ReadInt(hwnd, addr("fixbank2",-37), 0)-33)
		codes(12) = "jmp 0" + addr("fixbank2",0)
		Call inAsm(addr("fixbank2",-6), codes)
	End Function
	
	' 解開打怪範圍最小限制
	Public Function freeFrame(min)
		Call dm.WriteData(hwnd, "<ttha.bin>+29C0A",HEX(256-min))
		Call dm.WriteData(hwnd, "<ttha.bin>+29C1F",HEX(256-min))		
	End Function
	
	' 完全補給
	Public Function fullSupport()
		Call dm.WriteData(hwnd, addr("fullSupport",0),"EB")
	End Function

	Public Function atkSpeed(speed)
		Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+12100", 0,speed)	
	End Function

	Public Function addHP()
		Call dm.WriteInt(hwnd, "<ttha.bin>+22F5F", 0, &H20A)
		Call dm.WriteData(hwnd, "<ttha.bin>+22E5F", "90 90 90 90 90 90")
	End Function

	Public Function addMP()
		' 安裝變身
		Call dm.WriteData(hwnd, "<tthbn.bin>+A1580", "0C 02 00 00 00 00 00 00 A8 AD 00 00 00 00 00 00 00 00 00 00 00 00 00 00")
		Call dm.WriteData(hwnd, "<ttha.bin>+2293C", "D6 15 42 00")
		Call dm.WriteData(hwnd, "<ttha.bin>+1EF20", "90 90")
		Call dm.WriteData(hwnd, "<ttha.bin>+1EF3D", "90 90 90 90")
		Call dm.WriteData(hwnd, "<ttha.bin>+21614", "90 90 90")
		
		' 導入原本流程
		reDim codes(2)
		codes(0) = "and eax, 0F"
		codes(1) = "jmp 0" + HEX(ttha + &H213BF)
		codes(2) = "nop"
		Call asm("ttha.bin+213B6", codes)
		
		' 50%變身
		reDim codes(0)
		codes(0) = "jmp 0" + HEX(ttha + &H215ED)
		Call asm("ttha.bin+215DC", codes)
		
		reDim codes(4)
		codes(0) = "lea ecx,[eax+000001AC]"
		codes(1) = "mov eax,[ecx]"
		codes(2) = "imul eax,eax,02"
		codes(3) = "cmp eax,[ecx+04]"
		codes(4) = "ja 0" + HEX(ttha + &H21671)
		Call asm("ttha.bin+21608", codes)
	End Function

	Public Function defBuff(have8)
		' 安裝替身
		Call dm.WriteData(hwnd, "<tthbn.bin>+A1568", "0B 02 00 00 00 00 00 00 B4 C0 A8 AD 00 00 00 00 00 00 00 00 00 00 00 00")
		Call dm.WriteData(hwnd, "<ttha.bin>+1E9B1", "90 90")
		Call dm.WriteData(hwnd, "<ttha.bin>+1E9CE", "90 90 90 90")
		
		' 導入原本流程
		reDim codes(2)
		codes(0) = "and eax, 0F"
		codes(1) = "jmp 0" + HEX(ttha + &H213BF)
		codes(2) = "nop"
		Call asm("ttha.bin+213B6", codes)
		
		If have8 Then 
			' 有八卦才放替身
			reDim codes(5)
			codes(0) = "mov eax,[tthbn.bin+A171C]" ' 有八卦1
			codes(1) = "add eax,eax"
			codes(2) = "add eax,[ebx+04]"		' 沒替身0
			codes(3) = "cmp eax,02" 			' 1+1+0
			codes(4) = "jne 0" + HEX(ttha + &H21461)
			codes(5) = "jmp 0" + HEX(ttha + &H213DD)
			Call inAsm("ttha.bin+213C6", codes)
			
			' 替身消失時 清空八卦
			reDim codes(4)
			codes(0) = "cmp ecx,00000858"' 是清空替身流程
			codes(1) = "jne newmem+12"
			codes(2) = "mov dword ptr [tthbn.bin+A171C],0"' 清空八卦
			codes(3) = "mov dword ptr [ecx+tthbn.bin+A0D14],0"
			codes(4) = "jmp 0" + HEX(tthbn + &H421B6)
			Call inAsm("tthbn.bin+421AC", codes)		
		End If
	End Function

	Public Function asm(addr, codes)
		Dim code
		Call dm.AsmClear
		For Each code In codes	
			dm.AsmAdd toNum(code)
		Next
		code = dm.AsmCode(clng("&H" & toNum(addr)))
		Call dm.WriteData(hwnd, addBracket(addr), code)
	End Function

	Public Function toNum(str)
		Dim addr,num,match
		re.Pattern = "\w+\.\w+\+\w+"
		set match = re.Execute(str)
		If (match.count > 0) Then 
			addr = split(match(0), "+")
			num = dm.GetModuleBaseAddr(hwnd, addr(0)) + clng("&H" & addr(1))
		end if 
		toNum = re.Replace(str, HEX(num))
	End Function

	Public Function addBracket(str)
		re.Pattern = "(\w+\.\w+)"
		addBracket = re.Replace(str, "<$1>")
	End Function

	Public Function inAsm(addr, codes)
		Dim pid,hProcess,newmem,inCodes(1),i
		pid = dm.GetWindowProcessId(hwnd)
		hProcess = dw.OpenProcess(2035711, False, pid)
		newmem = HEX(dw.VirtualAllocEx(hProcess, 0, 50, &H3000, &H40))
			
		inCodes(0) = "jmp 0" & newmem
		inCodes(1) = "nop"
		Call asm(addr, inCodes)
		
		For i = 0 To UBound(codes)
			codes(i) = inToNum(codes(i),newmem)
		Next
		Call asm(newmem, codes)
	End Function

	Public Function inToNum(str,newmem)
		Dim addr,num,match
		re.Pattern = "newmem\+\w+"
		set match = re.Execute(str)
		If (match.count > 0) Then 
			addr = split(match(0), "+")
			num = clng("&H" & newmem) + clng("&H" & addr(1))
		end if 
		inToNum = re.Replace(str, "0"& HEX(num))
	End Function
	
	'util
	private Function simpleCall(hwndd,addr,parameters)
		dm.AsmClear 
		For Each p In parameters
			dm.AsmAdd "push 0" & toHEX(p)
		Next
		dm.AsmAdd "call 0" + toHEX(addr)
		dm.AsmCall hwndd, 1
		Delay 100
	End Function
	
	Private Function toHEX(data)
		If TypeName(data) <> "String" Then 
			data = HEX(data)
		End If
		toHEX = data
	End Function
	
	private Function getObjs(addr, cntAddr, length, conds, keys)
		Dim i,cnt,obj,key,objs,j,target
		addr = clng("&H"+addr)
		j = 0 : objs = array()
		if cntAddr = null Then
			cnt = 199
		else
			cnt = dm.ReadInt(hwnd, cntAddr, 0) - 1
		end if
		For i = 0 to cnt
			' 製作物件
			Set obj = CreateObject("Scripting.Dictionary")
			obj.add "no", i
			For Each key In keys
				If key(0) = "name" Then 
					obj.add key(0), dm.ReadString(hwnd, HEX(addr + length * i + key(1)), 0, 16)
				Else 
					obj.add key(0), dm.ReadInt(hwnd, HEX(addr + length * i + key(1)), 0)
				End If
			Next
			' id為零時結束
			if obj.item("id") = 0 Then
				exit For
			end if
			' 決定要不要放入
			target = obj(keys(0)(0))
			for each cond in conds
				re.Pattern = cond
				If typename(cond) = "Integer" and (target = cond) or typename(cond) = "String" and re.Test(target) Then 
					Redim Preserve objs(j) : Set objs(j) = obj : j = j + 1
					exit for
				End If				
			Next
		Next
		getObjs = objs
	End Function
	
	private function findAddr(code)
		if not addrs.Exists(code) then
			addrs.Add code, dm.FindData(hwnd,"00000000-FFFFFFFF",code)
		end if
		findAddr = CLNG("&H" & addrs.Item(code))
	end function
	
	Public Function updateAddr()
		dim code
		
		Set addrs = CreateObject("Scripting.Dictionary")
		addrs.Add "monster", "tthbn,00 8B 4D FC 6B C9 1C C7 81"
		addrs.Add "getIdByName", "tthbn,E0 69 C9 68 02 00 00 C7 81"
		addrs.Add "account", "tthbn,C4 6A 32 6A 00 68"
		addrs.Add "name", "tthbn,68 88 0C 00 00 6A 00 68"
		addrs.Add "isDisConnect", "tthbn,75 4D 83 3D"
		addrs.Add "bagCnt", "tthbn,F3 AB C7 05"
		addrs.Add "bag", "tthbn,E0 69 C0 18 01 00 00 8B 88"
		addrs.Add "bankCnt", "tthbn,54 FA FF FF A3"
		addrs.Add "bank", "tthbn,FA FF FF 69 C9 18 01 00 00 8B 91"
		addrs.Add "npcCnt", "tthbn,74 FF FF FF 3B 15"
		addrs.Add "npc", "tthbn,E1 05 8B 91"
		addrs.Add "getMsg", "tthbn,6A 78 68"
		addrs.Add "cash", "tthbn,89 45 98 8B 15"
		addrs.Add "isOffLine", "tthbn,6B D2 4A 81 C2"
		addrs.Add "level", "tthbn,25 33 C0 A0"
		For Each key In addrs.Keys
			code = split(addrs.item(key),",")
			result = dm.FindData(hwnd, "00000000-F0000000", code(1))
			result = split(result,"|")(0)
			l = Len(Replace(code(1), " ", "", 1, - 1 )) / 2
			result = dm.ReadInt(hwnd, HEX(CLNG("&H" & result) + l), 0) - eval(code(0))
			dm.WriteIni "addr", key, code(0) & "," & result, ".\QMScript\tthbn.ini"
		Next
		
		Set addrs = CreateObject("Scripting.Dictionary")
		addrs.Add "login", "ttha,53 56 57 8B F1 89 44 24 10"
		addrs.Add "crack", "ttha,74 7C 83 E8 02 74 1F 48"
		addrs.Add "moveSpeedAtk", "ttha,18 01 00 00 6A 02 50"
		addrs.Add "moveSpeed", "ttha,18 01 00 00 6A 02 51"
		addrs.Add "atkRange", "ttha,03 D1 89 54 24 04 DB 44 24 04 D9 FA E8"
		addrs.Add "fixbank", "ttha,6A 10 51 8D 4C 24 20"
		addrs.Add "fixbank2", "tthbn,66 C7 45 F0 1B"
		addrs.Add "distance", "ttha,02 81 E2 FF FF"
		addrs.Add "fullSupport", "ttha,75 29 3B C7"
		addrs.Add "invite", "tthbn,55 8B EC 83 EC 68 53 56 57 8D 7D 98 B9 1A 00 00 00 B8 CC CC CC CC F3 AB 66 C7 45 F4 0F 00 C6 45 E4 0D C6 45 E5 00 C6 45 E6 FF C6 45 E7 04 C6 45 E8 2C"
		addrs.Add "accept", "tthbn,55 8B EC 83 EC 68 53 56 57 8D 7D 98 B9 1A 00 00 00 B8 CC CC CC CC F3 AB 66 C7 45 F4 0F 00 C6 45 E4 0D C6 45 E5 00 C6 45 E6 FF C6 45 E7 04 C6 45 E8 2D"
		addrs.Add "give", "tthbn,55 8B EC 83 EC 6C 53 56 57 8D 7D 94 B9 1B 00 00 00 B8 CC CC CC CC F3 AB 66 C7 45 F4 13 00"
		addrs.Add "comfirm1", "tthbn,55 8B EC 83 EC 58 53 56 57 8D 7D A8 B9 16 00 00 00 B8 CC CC CC CC F3 AB 66 C7 45 FC 05 00 C6 45 F4 03 C6 45 F5 00 C6 45 F6 FF C6 45 F7 04 C6 45 F8 30"
		addrs.Add "comfirm2", "tthbn,55 8B EC 83 EC 58 53 56 57 8D 7D A8 B9 16 00 00 00 B8 CC CC CC CC F3 AB 66 C7 45 FC 05 00 C6 45 F4 03 C6 45 F5 00 C6 45 F6 FF C6 45 F7 04 C6 45 F8 2B"
		addrs.Add "apply", "tthbn,55 8B EC 83 EC 74 53 56 57 "
		addrs.Add "pop", "tthbn,C6 45 E8 34"
		addrs.Add "openBank", "tthbn,55 8B EC 83 EC 78 53 56 57 8D 7D 88 B9 1E 00 00 00 B8 CC CC CC CC F3 AB C7 45 EC"
		addrs.Add "buy", "tthbn,55 8B EC 83 EC 74 53 56 57 8D 7D 8C B9 1D 00 00 00 B8 CC CC CC CC F3 AB 8B 45"
		addrs.Add "sell", "tthbn,55 8B EC 83 EC 78 53 56 57 8D 7D 88 B9 1E 00 00 00 B8 CC CC CC CC F3 AB C7 45 F4"
		addrs.Add "reset", "tthbn,00 00 00 00 8B F4 FF 15"
		For Each key In addrs.Keys
			code = split(addrs.item(key),",")
			result = dm.FindData(hwnd, "00000000-F0000000", code(1))
			result = split(result,"|")(0)
			dm.WriteIni "addr", key, code(0) & "," & CLNG("&H" & result) - eval(code(0)), ".\QMScript\tthbn.ini"
		Next
	End Function
	
	Public Function IDs(key)
		IDs = split(dm.ReadIni("IDs",key,".\QMScript\tthbn.ini"),",")
	End Function
	Public Function teamST(key)
		teamST = split(dm.ReadIni("teamST",key,".\QMScript\tthbn.ini"),",")
	End Function
	Public Function memberST(key)
		memberST = split(dm.ReadIni("memberST",key,".\QMScript\tthbn.ini"),",")
	End Function
	Public function addr(key,offset)
		dim data : data = split(dm.ReadIni("addr", key, ".\QMScript\tthbn.ini"),",")
		addr = HEX(eval(data(0)) + Clng(data(1)) + offset)
	End function
	Public function addrStr(key,offset)
		dim data : data = split(dm.ReadIni("addr", key, ".\QMScript\tthbn.ini"),",")
		addrStr = "<" + data(0) + ".bin>+" + HEX(Clng(data(1)) + offset)
	End function
	Public function teamIDs()
		teamIDs = split(dm.ReadIni("setting","teamIDs",".\QMScript\local.ini"),",")
	End function
	Public function test()
		test = eval("tthbn")
	End function
End Class
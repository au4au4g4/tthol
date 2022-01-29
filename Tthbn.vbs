Class Tthbn

	Private dm,dw,re,hwnd,ttha,tthbn
	
	Public Sub Class_Initialize()
		Set dm = createobject("dm.dmsoft")
		Set dw = CreateObject("DynamicWrapper")
		Set re = New RegExp
		dw.Register "kernel32.dll", "OpenProcess", "i=uuu", "r=h" 
		dw.Register "kernel32.dll", "VirtualAllocEx", "i=lllll", "r=l"
    	End Sub
	
   	Public default function Init(p_hwnd)
		hwnd = p_hwnd
		ttha = dm.GetModuleBaseAddr(hwnd, "ttha.bin")
		tthbn = dm.GetModuleBaseAddr(hwnd, "tthbn.bin")
   	End function

	'flag
	Public Function isOffLine()
		isOffLine = (dm.ReadInt(hwnd, "<tthbn.bin>+111964", 0) = 0)
	End Function
	
	'read
	Public Function id()
		id = dm.ReadString(hwnd, "<tthbn.bin>+1099D0", 0, 16)
	End Function

	Public Function level()
		level = dm.ReadInt(hwnd, "<tthbn.bin>+10A6C0", 1)
	End Function
	Public Function money()
		money = dm.ReadInt(hwnd, "<tthbn.bin>+ACD24", 0) / 10 ^ 6
	End Function	
	Public Function expp()
		expp = dm.ReadInt(hwnd, "<tthbn.bin>+ACD28", 0) / 10 ^ 6
	End Function
	Public Function place()
		place = dm.ReadString(hwnd, "<tthbn.bin>+1099E4", 0, 16)
	End Function
	Public Function start()
		start = dm.ReadInt(hwnd, "<tthbn.bin>+ACD20", 0)
	End Function
	Public Function monster()
		sum = 0 : monId = 1 : cnt = 1 : i = 0
		While monId > 0
			base = tthbn + &HAEBE0 + 28 * i
			monId = dm.ReadInt(hwnd, HEX(base - 4), 0) / 10 ^ 4
			cnt = dm.ReadInt(hwnd, HEX(base), 0) / 10 ^ 4
			sum = sum + cnt
			i=i+1
		Wend
		monster = sum
	End Function
	Public Function cash ()
		cash = dm.ReadInt(hwnd, "<tthbn.bin>+109A5C", 0) / 10 ^ 6
	End Function
	Public Function deposit()
		deposit = dm.ReadInt(hwnd, "<tthbn.bin>+AFCB8", 0)
	End Function
	Public Function location()
		location = dm.ReadInt(hwnd, "[[<ttha.bin>+004EF82C]+98]+164", 0)
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
	
	Public Function getMsg()
		getMsg = dm.ReadString(hwnd, "<tthbn.bin>+10A758", 0, 115)
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
		getRecords = getObjs(tthbn + offset, null, 28, array(""), keys)
	End Function
	
	Public Function findNPC(name)
		Dim keys
		keys = array(array("name", 12),array("id", 4), array("sn", 0))
		set findNPC = getObjs(tthbn + &H105820, tthbn + &H1118B0, 32, array(name), keys)(0)
	End Function
	
	Public Function getBag(names)
		getBag = getItems(tthbn + &H102148, tthbn + &H1118B4, names)
	End Function
	
	Public Function getBank(names)
		getBank = getItems(tthbn + &HAFCC0, tthbn + &HAFCBC, names)
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
	
	Public Function stops()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,[" + HEX(ttha + &H4EF82C) + "]"
		dm.AsmAdd "mov ecx,[ecx+98]"
		dm.AsmAdd "call "+ HEX(ttha+&H54E80)
		dm.AsmCall hwnd,1
	End Function
	
	Public Function learn(code)
		code = split(code, chr(9))
		for i = 1 to code(1)
			simpleCall hwnd, tthbn + &H2BA10, array(i,code(0))
		Next
	End Function
	
	Public Function trades(buyer, keywords, ByVal cnt)
		dim bag : bag = getBag(keywords)
		For i = 0 To UBound(bag)
			Set item = bag(i)
			itemCnt = item.item("cnt") 
			if cnt = -1 Then
				trade buyer, item, itemCnt
			elseif itemCnt - cnt >= 0 then
				trade buyer, item, cnt
				exit for
			else
				trade buyer, item, itemCnt
				cnt = cnt - itemCnt
			end if
		Next
	End Function	
	
	Public Function trade(buyer, item, cnt)	
		if hwnd - buyer = 0 then
			exit Function
		end if

		bTthbn = dm.GetModuleBaseAddr(buyer, "tthbn.bin")
		sAddr = clng("&H" & dm.FindString(buyer, HEX(bTthbn + &HE9788) & "-" & HEX(bTthbn + &HE9788 + &H7850), id, 0)) - 16
		sID = dm.ReadInt(buyer, HEX(sAddr), 0)

		bName = dm.ReadString(buyer, "<tthbn.bin>+1099D0", 0, 16)
		bAddr = clng("&H" & dm.FindString(hwnd, HEX(tthbn + &HE9788) & "-" & HEX(tthbn + &HE9788 + &H7850), bName, 0)) - 16
		bID = dm.ReadInt(hwnd, HEX(bAddr), 0)
		
		sn = item.item("sn")
		iID = item.item("id")				
		simpleCall hwnd, tthbn + &H27030, array(bID, &HEA64)	' 邀請	
		simpleCall buyer, bTthbn + &H27200, array(sID, &HEA64)	' 接受	
		simpleCall hwnd, tthbn + &H273D0, array(cnt, sn, iID)	' 交付	
		simpleCall hwnd, tthbn + &H27BC0, array()	' 確定1
		simpleCall buyer, bTthbn + &H27BC0, array()
		simpleCall hwnd, tthbn + &H27CF0, array()	' 確定2
		simpleCall buyer, bTthbn + &H27CF0, array()
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
			simpleCall hwnd, tthbn + &H259B0, array(0,item.item("cnt"),item.item("sn"),item.item("id"))
		Next
	End Function
	
	Public Function push(names)
		dim items, item
		openBank()
		items = getBag(names)
		for each item in items
			simpleCall hwnd, tthbn + &H259B0, array(1,item.item("cnt"),item.item("sn"),item.item("id"))
		Next
	End Function
	
	Public Function openBank()
		Set npc = findNPC("伙計")
		simpleCall hwnd, tthbn + &H24470, array(npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function buy(npc,item)
		simpleCall hwnd, tthbn + &H24E90, array(item.item("cnt"),item.item("id"),npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function sell(npc,item)
		simpleCall hwnd, tthbn + &H251A0, array(item.item("cnt"),item.item("sn"),item.item("id"),npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function kill(x,y,sn,id,skill)
		simpleCall hwnd, tthbn + &H23900, array(x,y,sn,id,skill)
	End Function

	Public Function login()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,0" + HEX(ttha + &H4EF82C)
		dm.AsmAdd "mov ecx,[ecx]"
		dm.AsmAdd "push 0"
		dm.AsmAdd "push 03EC"
		dm.AsmAdd "call 0" + HEX(ttha + &H8B010)
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
	
	Public Function reLoginMin(min)
		call dm.WriteInt(hwnd, "<ttha.bin>+3C334", 0, min * 60000)
	End Function

	Public function fixBank()
		Call dm.WriteData(hwnd, "<ttha.bin>+69CDA", "10")
		reDim codes(8)
		codes(0) = "mov ecx,[esp+20]"
		codes(1) = "push 0"
		codes(2) = "push ecx"
		codes(3) = "lea ecx,[esp+20]"
		codes(4) = "call 0" + HEX(ttha + &H8871A)
		codes(5) = "call 0" + HEX(ttha + &H8DBD1)
		codes(6) = "test eax,eax"
		codes(7) = "je ttha.bin+69C8E"
		codes(8) = "jmp ttha.bin+69C85"
		Call inAsm("ttha.bin+69C7C ", codes)
	End Function
	
	Public function removePlayer()
		Call dm.WriteData(hwnd, "<ttha.bin>+66EAD", "909090909090")
		reDim codes(4)
		codes(0) = "push edx"
		codes(1) = "push 0000013C"
		codes(2) = "call 0" + HEX(ttha + &H7E764)
		codes(3) = "call 0" + HEX(tthbn + &H26180)
		codes(4) = "jmp ttha.bin+67396"
		Call inAsm("ttha.bin+6738B", codes)
	End Function
	
	Public function updateItem(name, action)
		dim id
		id = dm.ReadIni("id", name, ".\item.ini")
		action = split(dm.ReadIni("action", action, ".\item.ini"),",")
		simpleCall hwnd, tthbn + 57568, array(action(1), action(0), id)
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
	Public Function crackAll(speed,range,def)
		crack()
		moveSpeed(speed)
		atkRange(range)
		atkSpeed(100)
		addHP()
		addMP()
		defBuff(def)
		reLoginMin(2)
	End Function
	
	Public Function crack()
		Call dm.WriteData(hwnd, "<tthbn.bin>+ACD08", "E1 E7 F6 B7 84 AC ED B4 E1 E0 A0 9C A9 FB ED E6 B6 8B B1")
		reDim codes(0) : codes(0) = "jmp ttha.bin+3B1E0"
		Call asm("ttha.bin+3B162", codes)	
	End Function

	'280
	Public Function moveSpeed(speed)
		Call dm.WriteInt(hwnd,"<ttha.bin>+3FD3D",0,speed)
		Call dm.WriteInt(hwnd,"<ttha.bin>+3FD53",0,speed)
	End Function

	Public Function atkRange(range)
		reDim codes(5)
		codes(0) = "mov eax,0" & HEX(range ^ 2)
		codes(1) = "cmp edx,eax"
		codes(2) = "jg ttha.bin+535AA"
		codes(3) = "cmp ecx,eax"
		codes(4) = "jg ttha.bin+535AA"
		codes(5) = "jmp ttha.bin+535AF"
		Call asm("ttha.bin+5358C", codes)
	End Function
	
	Public Function setRange(range)
		call dm.WriteInt(hwnd, "<ttha.bin>+5358E", 0, range ^ 2)
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
		codes(1) = "jmp ttha.bin+213BF"
		codes(2) = "nop"
		Call asm("ttha.bin+213B6", codes)
		
		' 50%變身
		reDim codes(0)
		codes(0) = "jmp ttha.bin+215ED"
		Call asm("ttha.bin+215DC", codes)
		
		reDim codes(4)
		codes(0) = "lea ecx,[eax+000001AC]"
		codes(1) = "mov eax,[ecx]"
		codes(2) = "imul eax,eax,02"
		codes(3) = "cmp eax,[ecx+04]"
		codes(4) = "ja ttha.bin+21671"
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
		codes(1) = "jmp ttha.bin+213BF"
		codes(2) = "nop"
		Call asm("ttha.bin+213B6", codes)
		
		If have8 Then 
			' 有八卦才放替身
			reDim codes(5)
			codes(0) = "mov eax,[tthbn.bin+A171C]" ' 有八卦1
			codes(1) = "add eax,eax"
			codes(2) = "add eax,[ebx+04]"		' 沒替身0
			codes(3) = "cmp eax,02" 			' 1+1+0
			codes(4) = "jne ttha.bin+21461"
			codes(5) = "jmp ttha.bin+213DD"
			Call inAsm("ttha.bin+213C6", codes)
			
			' 替身消失時 清空八卦
			reDim codes(4)
			codes(0) = "cmp ecx,00000858"' 是清空替身流程
			codes(1) = "jne newmem+12"
			codes(2) = "mov dword ptr [tthbn.bin+A171C],0"' 清空八卦
			codes(3) = "mov dword ptr [ecx+tthbn.bin+A0D14],0"
			codes(4) = "jmp tthbn.bin+421B6"
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
	private Function simpleCall(hwnd,addr,parameters)
		dm.AsmClear 
		For Each p In parameters
			dm.AsmAdd "push 0" & HEX(p)
		Next
		dm.AsmAdd "call " + HEX(addr)
		dm.AsmCall hwnd, 1
		Delay 100
	End Function
	
	private Function getObjs(addr, cntAddr, length, conds, keys)
		Dim i,cnt,obj,key,objs,j,target
		j = 0 : objs = array()
		if cntAddr = null Then
			cnt = 199
		else
			cnt = dm.ReadInt(hwnd, HEX(cntAddr), 0) - 1
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
				If typename(cond) = "Integer" and (target = cond) or typename(cond) = "String" and instr(target,cond)>0 Then 
					Redim Preserve objs(j) : Set objs(j) = obj : j = j + 1
					exit for
				End If				
			Next
		Next
		getObjs = objs
	End Function
End Class
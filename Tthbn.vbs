Class Tthbn

	Private dm,hwnd,ttha,tthbn
	
	Public Sub Class_Initialize()
		Set dm = createobject("dm.dmsoft")
    End Sub
	
    Public default function Init(p_hwnd)
		hwnd = p_hwnd
		ttha = dm.GetModuleBaseAddr(hwnd, "ttha.bin")
		tthbn = dm.GetModuleBaseAddr(hwnd, "tthbn.bin")
    End function

	'flag
	Public Function isOffLine()
		isOffLine = (dm.ReadInt(hwnd, "<tthbn.bin>+1116B6", 0) = 0)
	End Function
	
	'read
	Public Function id()
		id = dm.ReadString(hwnd, "<tthbn.bin>+109720", 0, 16)
	End Function
	Public Function level()
		level = dm.ReadInt(hwnd, "<tthbn.bin>+10A410", 1)
	End Function
	Public Function money()
		money = dm.ReadInt(hwnd, "<tthbn.bin>+ACD24", 0) / 10 ^ 6
	End Function	
	Public Function expp()
		expp = dm.ReadInt(hwnd, "<tthbn.bin>+ACD28", 0) / 10 ^ 6
	End Function
	Public Function place()
		place = dm.ReadString(hwnd, "<tthbn.bin>+109734", 0, 16)
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
		cash = dm.ReadInt(hwnd, "<tthbn.bin>+1097AC", 0)
	End Function
	Public Function deposit ()
		deposit = dm.ReadInt(hwnd, "<tthbn.bin>+AFCB8", 0)
	End Function
	
	Public Function getMsg()
		getMsg = dm.ReadString(hwnd, "<tthbn.bin>+10A4A8", 0, 115)
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
		getRecords = getObjs(tthbn + offset, 28, array(""), keys)
	End Function
	
	Public Function findNPC(names)
		Dim keys
		keys = array(array("name", 12),array("id", 4), array("sn", 0))
		set findNPC = getObjs(tthbn + &H105570, 32, names, keys)(0)
	End Function
	
	Public Function getBag(names)
		getBag = getItems(&H101E98, names)
	End Function
	
	Public Function getBank(names)
		getBank = getItems(&HAFCC0, names)
	End Function
	
	private Function getItems(offset, names)
		Dim keys
		keys = array(array("name", 32),array("sn", 0), array("id", 4), array("cnt", 8))
		getItems = getObjs(tthbn + offset, 280, names, keys)
	End Function
	
	'function
	Public Function talkNPC()
		call dm.WriteInt(hwnd,"[[<ttha.bin>+004EF82C]+98]+150",0,&HE7) 'resetTalk
		simpleCall hwnd, ttha + &H44D90, array()
	End Function
	 
	'1開始 2對話 45678選項
	Function talkOption(npc, arr)
		For Each a In arr
			If a = 0 Then 
				Delay 500
			Else 
				simpleCall hwnd, tthbn + &H24720, array(1, npc, a)
				Delay 50
			End If
		Next
	End Function
	
	Function stops()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,[" + HEX(ttha + &H4EF82C) + "]"
		dm.AsmAdd "mov ecx,[ecx+98]"
		dm.AsmAdd "call "+ HEX(ttha+&H54E80)
		dm.AsmCall hwnd,1
	End Function
	
	Public Function learn(code)
		code = split(code, chr(9))
		simpleCall hwnd, tthbn + &H2BA10, array(code(1),code(0))
	End Function
	
	Public Function trades(buyer, keywords)
		dim bag : bag = t.getBag(keywords)
		For i = 0 To UBound(bag)
			Set item = bag(i)
			t.trade buyer, item
		Next
	End Function	
	
	Private Function trade(buyer, item)	
		if hwnd - buyer = 0 then
			exit Function
		end if

		bTthbn = dm.GetModuleBaseAddr(buyer, "tthbn.bin")
		sAddr = clng("&H" & dm.FindString(buyer, HEX(bTthbn + &HE94D8) & "-" & HEX(bTthbn + &HF0D28), id, 0)) - 16
		sID = dm.ReadInt(buyer, HEX(sAddr), 0)

		bName = dm.ReadString(buyer, "<tthbn.bin>+109720", 0, 16)
		bAddr = clng("&H" & dm.FindString(hwnd, HEX(tthbn + &HE94D8) & "-" & HEX(tthbn + &HF0D28), bName, 0)) - 16
		bID = dm.ReadInt(hwnd, HEX(bAddr), 0)
		
		sn = item.item("sn")
		iID = item.item("id")
		cnt = item.item("cnt")
		
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
		Set npc = t.findNPC("伙計")
		simpleCall hwnd, tthbn + &H24470, array(npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function buy(npc,item)
		simpleCall hwnd, tthbn + &H24E90, array(item.item("cnt"),item.item("id"),npc.item("sn"),npc.item("id"))
	End Function
	
	Public Function sell(npc,item)
		simpleCall hwnd, tthbn + &H251A0, array(item.item("cnt"),item.item("sn"),item.item("id"),npc.item("sn"),npc.item("id"))
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
	
	private Function getObjs(addr, length, conds, keys)
		Dim i,flag,obj,key,objs,j,target
		i = 0 : j = 0 : objs = array()
		flag = dm.ReadInt(hwnd, HEX(addr + length * i), 0)
		Do While flag <> 0
			Set obj = CreateObject("Scripting.Dictionary")
			For Each key In keys
				If key(0) = "name" Then 
					obj.add key(0), dm.ReadString(hwnd, HEX(addr + length * i + key(1)), 0, 16)
				Else 
					obj.add key(0), dm.ReadInt(hwnd, HEX(addr + length * i + key(1)), 0)
				End If
			Next
			target = obj(keys(0)(0))
			for each cond in conds
				If typename(cond) = "Integer" and (target = cond) or typename(cond) = "String" and instr(target,cond)>0 Then 
					Redim Preserve objs(j) : Set objs(j) = obj : j = j + 1
					exit for
				End If				
			next
			i = i + 1
			flag = dm.ReadInt(hwnd, HEX(addr + length * i), 0)
		Loop
		getObjs = objs
	End Function
	
End Class
Class Tthol

	Private dm,hwnd
	
    Public default function Init(p_hwnd)
		hwnd = p_hwnd
		Set dm = createobject("dm.dmsoft")
		Call dm.WriteInt(hwnd, "<tthola.dat>+36B3FC", 0, 150)
    End function

	' ------------------------------ram flag------------------------------

	Public Function getUsingBuff()
		Dim headAddr,lenAddr
		lenAddr = "[[<tthola.dat>+3EABBC]+10]+20FC"
		headAddr = read("[<tthola.dat>+3EABBC]+10") + &H2104
		getUsingBuff = getAddrs(headAddr,lenAddr)
	End Function

	Public Function getXY()
		dim x,y,xy
		x = read("[<tthola.dat>+21A030]+B2C")/40
		y = read("[<tthola.dat>+21A030]+B30")/40
		getXY = array(x,y)
	End Function

	Public Function getSkill(code)
		getSkill = read(array("[<tthola.dat>+3EABBC]+10", &H3050 + (code - 1) * 8))
	End Function
	
	Public Function getTarget()
		getTarget = read("[[<tthola.dat>+3EABBC]+10]+20E8")
	End Function

	Public Function talkWnd()
		talkWnd = (read("[[[[<tthola.dat>+3EABBC]+10]+8]+104]+20") = 3)
	End Function
	
	public function selectWnd()
		selectWnd = (read("<tthola.dat>+1FAE84") = 0)
	end function 

	Public Function location()
		location = readString("[[[<tthola.dat>+3EABBC]+10]+AE4]+2B04",16)
	End Function

	Public Function name()
		name = readString("[[<tthola.dat>+3EED84]+12C]+8",16)
	End Function

	Public Function task()
		task = (read("[[<tthola.dat>+3EABBC]+10]+1768") <> 0)
	End Function
	
	Public Function walking()
		walking = (read(getMainAddr()+&HA0) = 1)
	End Function
	
	Public Function getMainAddr()
		Dim edx, eax
		edx = read("[[[<tthola.dat>+3ED97C]+10]+C]+8")
		eax = read("[[<tthola.dat>+3ED97C]+10]+AFC") and "&HFFFF"
		getMainAddr = read(edx + eax * 4)
	End Function
	
	' 無攻擊0 普功1 技能3
	Public Function attackType()
		attackType = read(getMainAddr() + &H10C8)
	End Function
	
	Public Function getBag(keywords)
		dim item
		set bag = getItems("bag")
		For i = bag.Count - 1 To 0 step -1
			set item = bag.GetByIndex(i)
			for each keyword in keywords
				if instr(item.item("name"),keyword) < 1 then
					bag.RemoveAt(i)
				end if
			next
		Next
		set getBag = bag
	End Function
	
	Function getItems(place)
		Dim bagAddr, lenAddr, headAddr, itemAddr, item, id,bag
		Set bag = CreateObject("System.Collections.SortedList")
		bagAddr = getBagInfo(place)
		
		lenAddr = bagAddr - 4
		headAddr = read(bagAddr)
		itemAddrs = getAddrs(headAddr, lenAddr)
		
		For Each itemAddr In itemAddrs
			Set item = CreateObject("Scripting.Dictionary")
			id = read(itemAddr + 4)
			item.Add "addr", itemAddr
			item.Add "no", dm.ReadInt(hwnd,HEX(itemAddr + 3), 1)
			item.Add "id", read(itemAddr + 5)
			item.Add "sn", read(itemAddr + 9)
			item.Add "amount", read(itemAddr + 16)
			item.Add "name", readString(itemAddr + 20, 12)
			While bag.ContainsKey(id)
				id = id + 1
			Wend
			bag.Add id, item
		Next
		set getItems = bag
	End Function

	Function getBagInfo(place)
		If place = "bank" Then 
			getBagInfo = read("[[<tthola.dat>+003ED978]+10]+2134") + &H1A8
		Else 
			Dim eax,esi
			If place = "bag" Then 
				getBagInfo = read("[[[<tthola.dat>+003ED978]+10]+87C]") + &H360
			Else 
				getBagInfo = read("[[[<tthola.dat>+003ED978]+10]+87C]") + &H368	'pet
			End If
		End If
	End Function
	
	Public Function getMonster()
		Set getMonster = getNpc(array(array(20, &H20000001),array(&H10C1,0)))
	End Function
	
	Public Function getNpc(conds)
		Set npc = CreateObject("Scripting.Dictionary")
		dim pass, addr : addr = getMainAddr()
		do While addr <> 0
			pass = true
			for each cond in conds
				pass = pass and (read(addr + cond(0)) = cond(1))
			next
			if pass Then
				npc.Add "addr", HEX(addr)
				npc.Add "no", dm.ReadInt(hwnd, HEX(addr + 300), 1)
				npc.Add "id", read(addr + 302)
				npc.Add "sn", read(addr + 306)
				npc.Add "key", read(addr + 28)
				npc.Add "x", read(addr + 38)
				npc.Add "y", read(addr + 42)
				npc.Add "tp", read(addr + 220)
				npc.Add "dead", dm.ReadInt(hwnd, HEX(addr + &H10C3), 2)
				npc.Add "name", readString(addr + 484,10)
				exit do
			End If
			addr =  read(addr + &H124)
		Loop
		set getNpc = npc
	End Function
	
	Public Function printNpc(conds)
		dim pass, addr : addr = getMainAddr()
		do While addr <> 0
			pass = true
			for each cond in conds
				pass = pass and (read(addr + cond(0)) = cond(1))
			next
			if pass Then
				TracePrint join(array(HEX(addr),HEX(dm.ReadInt(hwnd, HEX(addr + 300), 1)),HEX(read(addr + 302)),HEX(read(addr + 306)),HEX(read(addr + 28)),readString(addr + 484,10),read(addr + 38)/40,read(addr + 42)/40,read(addr + 220)))
			End If
			addr =  read(addr + &H124)
		Loop
	End Function

	' ------------------------------ram 動作------------------------------

	' 移動
	Public Function go(x,y)
		dm.AsmClear 
		dm.AsmAdd "mov eax,"+HEX(getMainAddr())
		dm.AsmAdd "mov ebx,0"+Hex(x * 40)
		dm.AsmAdd "mov edi,0"+Hex(y * 40)
		dm.AsmAdd "push 00000001"
		dm.AsmAdd "call 00407480"
		dm.AsmCall hwnd,1
	End Function
		
	' 移動
	Public Function go1(xy)
		send "1E","8",array(typ,3,id,0,sn,0),array(1,2,2,2,2,2)
	End Function

	' 去找NPC
	Public Function goNpc(id)
		dm.AsmClear 
		dm.AsmAdd "mov eax,[7EABBC]"
		dm.AsmAdd "mov eax,[eax+10]"
		dm.AsmAdd "mov eax,[eax+37D8]"
		dm.AsmAdd "mov ebx,[626D14]"
		dm.AsmAdd "mov ecx,0000028F"
		dm.AsmAdd "mov [ebx+000001F0],ecx"
		dm.AsmAdd "mov ecx,000001EF"
		dm.AsmAdd "mov [ebx+000001F4],ecx"
		dm.AsmAdd "mov ecx,01"
		dm.AsmAdd "mov [eax+40],ecx"
		dm.AsmAdd "mov ebx,"+HEX(id)
		dm.AsmAdd "mov [eax+64],ebx"
		dm.AsmAdd "call 0042F550"
		dm.AsmCall hwnd, 1
	End Function
	
	Public Function talkNpc1()
		dm.AsmClear 
		dm.AsmAdd "mov eax,[07E97D0]"
		dm.AsmAdd "mov eax,[eax+0000012C]"
		dm.AsmAdd "call 0043F8C0"
		dm.AsmCall hwnd, 1
	End Function
	
	Public Function talkNpc(id,arr)
		dim i,t,typ, npc
		set npc = getNpc(array(array(302,id)))
		if npc.count > 0 Then
			for each a in arr
				t = 0
				if IsArray(a) Then
					t = a(1)-1
					typ = a(0)
				Else
					typ = a
				end if
				for i = 0 to t
					send "28","B",array(typ,3,id,0,npc.item("sn"),0),array(1,2,2,2,2,2)
				next
			next
		end if
	End Function
	
	'傳送
	Public Function trans()
		dim portal, sn
		set portal = getNpc(array(array(302, &H65)))
		no = portal.item("no")
		sn = portal.item("sn")
		if portal.count > 0 Then
			send "28","B",array(3,no,&H65,sn),array(1,2,4,4)
		end if
	End Function
	
	'穿裝
	Public Function wear(eqpt)
		dim no,id,sn
		no = eqpt.item("no")
		id = eqpt.item("id")
		sn = eqpt.item("sn")
		send "2A","B",array(no,id,sn,3),array(2,4,4,1)
	End Function
	
	Public Function selNpc(no)
		dm.AsmClear 
		dm.AsmAdd "mov ecx, [7EABBC]"
		dm.AsmAdd "mov ecx ,[ecx+10]"
		dm.AsmAdd "mov ecx ,[ecx+3C]"		
		dm.AsmAdd "add ecx ,2A0"
		dm.AsmAdd "push " & (no-1)
		dm.AsmAdd "call 00435180"
		dm.AsmCall hwnd, 1
	End Function

	' 鎖定怪物
	Public Function aimMonster(yes)
		If yes Then 
			Call dm.WriteData(hwnd, "<tthola.dat>+4312E", "8B4424048B4014C20C00")
		Else 
			Call dm.WriteData(hwnd, "<tthola.dat>+4312E", "C20C0090909090909090")
		End If
	End Function
	
	'學技能
	Public function learnSkills(codes)
		For Each code In codes
			learnSkill (HEX(code(0)*100 + code(1)))
			Delay 200
		Next
	end Function
	
	'學技能
	Public function learnSkill(code)
		dm.AsmClear 
		dm.AsmAdd "mov eax,0" & code
		dm.AsmAdd "call 00440780"
		dm.AsmCall hwnd, 1
	end Function
	
	'移除技能
	Public function reomveSkill(code)
		dm.AsmClear 
		dm.AsmAdd "mov eax,0" & code
		dm.AsmAdd "mov ecx,07081"
		dm.AsmAdd "call 004407A0"
		dm.AsmCall hwnd, 1
	end Function
	
	' 逛攤販
	Public function shopping(addr)
		dm.AsmClear
		dm.AsmAdd "mov eax," & addr
		dm.AsmAdd "mov eax,[eax+000000D4]"
		dm.AsmAdd "mov edi,"& HEX(dm.ReadInt(hwnd,"[<tthola.dat>+003ED978]+10", 0))
		dm.AsmAdd "call 004855D0"
		dm.AsmCall hwnd, 1
	end Function
	
	' 煉化
	Public Function compound(id,addr)
		Call dm.WriteInt(hwnd, "[[<tthola.dat>+003ED978]+10]+18B4", 0, id)
		Call dm.WriteInt(hwnd, "[[<tthola.dat>+003ED978]+10]+18CC", 0, addr + 3)
		dm.AsmClear 
		dm.AsmAdd "mov eax,01"
		dm.AsmAdd "mov esi,0"& HEX(dm.ReadInt(hwnd, "[<tthola.dat>+003ED978]+10", 0))
		dm.AsmAdd "call 004E7D80"
		dm.AsmCall hwnd,1
	End Function
	
	' 煉化
	Public Function compound0(cid,item)
		dim no,id,sn
		no = item.item("no")
		id = item.item("id")
		sn = item.item("sn")
		send "44","F",array(cid,0,0,0,1),array(4,4,4,2,1)
		send "45","32",array(no,id,0,sn,0,0,0,0,0,0,0,0,0,0),array(2,2,2,2,4,4,4,4,4,4,4,4,4,4)
		send "AD","1E",array(0,0,0,0,0,0,0,0),array(4,4,4,4,4,4,4,4)
		send "49","0",array(),array()
	End Function
	
	' 普攻
	Public Function atk(monster)
		dm.AsmClear 
		dm.AsmAdd "push " & HEX(monster.item("key"))
		dm.AsmAdd "push " & HEX(dm.ReadInt(hwnd, "[<tthola.dat>+3ED97C]+10", 0))
		dm.AsmAdd "call 00489C60"
		dm.AsmCall hwnd, 1	
	End Function
	
	'領取
	Public Function pop(item)
		Dim id, no, sn, amount
		id = item.item("id")
		no = item.item("no")
		sn = item.item("sn")
		amount = item.item("amount")
		send "34","B",array(no,id,0,sn,0,amount),array(2,2,2,2,8,2)
	End Function

	'存入
	Public Function push(item)
		Dim id, no, sn, amount
		id = item.item("id")
		no = item.item("no")
		sn = item.item("sn")
		amount = item.item("amount")
		send "32","B",array(no,id,0,sn,0,amount),array(2,2,2,2,8,2)
	End Function
		
	' 登入
	Public Function login1(id,pwd)
		dm.AsmClear 
		dm.AsmAdd "sub esp,30"
		dm.AsmAdd "mov ecx, 65677774"
		dm.AsmAdd "mov [esp], ecx"
		dm.AsmAdd "mov ecx, 7361706E"
		dm.AsmAdd "mov [esp+4], ecx"
		dm.AsmAdd "mov ecx, 00190073"
		dm.AsmAdd "mov [esp+8], ecx"
		dm.AsmAdd "mov ecx, 386A6780"
		dm.AsmAdd "mov [esp+14], ecx"	
		dm.AsmAdd "mov ecx, 346A6433"
		dm.AsmAdd "mov [esp+18], ecx"
		dm.AsmAdd "mov ecx, 05570034"
		dm.AsmAdd "mov [esp+1C], ecx"
		
		dm.AsmAdd "push 0"
		dm.AsmAdd "push 2A"
		dm.AsmAdd "push 0"

		dm.AsmAdd "lea ebx,[esp+0C]"
		dm.AsmAdd "mov ecx,00000003"		
		dm.AsmAdd "call 0043EBD0"
		dm.AsmAdd "add esp,30"
		dm.AsmAdd "ret"
		dm.AsmCall hwnd,1
	End Function
	
	' 登入
	Public Function login(id,pwd)
		dm.AsmClear
		dm_ret = dm.WriteString(hwnd,"0076C2C0",0,"gj83dj44")
		dm.AsmAdd "mov ecx, 1704FBA0"
		dm.AsmAdd "call 0439C40"
		dm.AsmCall hwnd,1
	End Function

	Public Function socket()
		dm.AsmClear 
		dm.AsmAdd "mov ecx,[07EE054]"
		dm.AsmAdd "mov edx,[07E97D0]"
		dm.AsmAdd "mov eax,01"
		dm.AsmAdd "imul eax,eax,3C"
		dm.AsmAdd "add eax,ecx"
		dm.AsmAdd "mov ecx,[edx+04]"
		dm.AsmAdd "mov edx,[eax+30]"
		dm.AsmAdd "push edx"
		dm.AsmAdd "add eax,20"
		dm.AsmAdd "push eax"
		dm.AsmAdd "push ecx"
		dm.AsmAdd "call 0596880"
		dm.AsmCall hwnd,1
	End Function
	
	' ------------------------------記憶體------------------------------

	Private Function read(addrs)
		read = dm.ReadInt(hwnd, getScript(addrs), 0)
	End Function

	Private Function readString(addrs,len)
		readString = dm.ReadString(hwnd,getScript(addrs),0,len)
	End Function

	Private Function getScript(addrs)
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

	Private Function toHEX(data)
		If TypeName(data) <> "String" Then 
			data = HEX(data)
		End If
		toHEX = data
	End Function

	Private Function getAddrs(headAddr,lenAddr)
		Dim i,len
		Redim result(0)
		len = read(lenAddr)
		For i = 0 To len - 1
			Redim Preserve result(i)
			 result(i) = read(headAddr+4*i)
		Next
		getAddrs = result
	End Function
	
	' ------------------------------備用------------------------------
	
	Public Function talkNpc3(a,b,c,d)
		dm.AsmClear 
		dm.AsmAdd "sub esp,0C"
		dm.AsmAdd "mov al,0" & a			'開始01,交談02,傳點03,選項04...
		dm.AsmAdd "mov [esp+04],al"
		dm.AsmAdd "mov eax," & b	'1
		dm.AsmAdd "mov [esp+05],eax"
		dm.AsmAdd "mov eax," & c	'1
		dm.AsmAdd "mov [esp+09],eax"
		dm.AsmAdd "push 0B"
		dm.AsmAdd "lea ebx,[esp+08]"
		dm.AsmAdd "mov ecx,00000028"
		dm.AsmAdd "mov ax," & d			'1
		dm.AsmAdd "mov [esp+11],ax"
		dm.AsmAdd "call 004309D0"
		dm.AsmAdd "add esp,0C"
		dm.AsmCall hwnd, 1
	End Function
	
	Public Function talking()
		talking = read("[[<tthola.dat>+003E77D4]+10]+1534")<>0
	End Function
	
	' 開啟 左鍵無法點擊NPC
	Public Function lockLClickNPC(yes)
		If yes Then 
			Call dm.WriteData(hwnd, "<tthola.dat>+5FC8A", "90E9")
		Else 
			Call dm.WriteData(hwnd, "<tthola.dat>+5FC8A", "0F84")
		End If
	End Function
		
	' TOFIX
	Public Function skilling()
		edx = read("[[[<tthola.dat>+3EABBC]+10]+C]+8")
		eax = read("[[<tthola.dat>+3EABBC]+10]+214C") and "&HFFFF"
		skilling = (read(array(edx + eax * 4, "10")) = 2)
	End Function
	
	Private Function send(cid,pid,datas,lens)
		dim i, j : j = 0
		dm.AsmClear 
		dm.AsmAdd "sub esp,0"& HEX(UBound(lens)*4+4)
		dm.AsmAdd "lea ebx,[esp]"
		for i = 0 to UBound(datas)
			dm.AsmAdd "mov ecx,0"& hex(datas(i))
			dm.AsmAdd "mov [esp+0"& HEX(j)&"],ecx"
			j =  j+ lens(i)
		next		
		dm.AsmAdd "push 0"& pid
		dm.AsmAdd "push 0"
		dm.AsmAdd "mov ecx,0"& cid
		dm.AsmAdd "call 0043EAB0"
		dm.AsmAdd "add esp,0"& HEX(UBound(lens)*4+4)
		dm.AsmCall hwnd, 1
		Delay 200
	End Function
	
End Class
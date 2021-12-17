Class Tthol

	Private dm,hwnd
	
    Public default function Init(p_hwnd)
		hwnd = p_hwnd
		Set dm = createobject("dm.dmsoft")
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
	
	' 取得玄天NPC ID
	Public Function getNpc()
		getNpc = read("[[[[[<tthola.dat>+3EABBC]+10]+37D8]+60]+2244]+8")
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
	
	Function getItems(place)
		Dim bagAddr, lenAddr, headAddr, itemAddr, item, id,bag
		Set bag = CreateObject("System.Collections.SortedList")
		bagAddr = getBagInfo(place)
		
		lenAddr = bagAddr - 4
		headAddr = read(bagAddr)
		itemAddrs = getAddrs(headAddr, lenAddr)
		
		For Each itemAddr In itemAddrs
			Set item = CreateObject("Scripting.Dictionary")
			id = read(itemAddr + 4 * 1)
			item.Add "id", HEX(id / &H100)
			item.Add "addr", itemAddr
			item.Add "amount", read(itemAddr + 4 * 4)
			item.Add "name", readString(itemAddr + 4 * 5, 12)
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

	' ------------------------------ram 動作------------------------------

	' 移動
	Public Function go(xy)
		dm.AsmClear 
		dm.AsmAdd "mov eax,"+HEX(getMainAddr())
		dm.AsmAdd "mov ebx,0"+Hex(xy(0) * 40)
		dm.AsmAdd "mov edi,0"+Hex(xy(1) * 40)
		dm.AsmAdd "push 00000001"
		dm.AsmAdd "call 00407480"
		dm.AsmCall hwnd,1
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
	
	Public Function talkNpc()
		dm.AsmClear 
		dm.AsmAdd "mov eax,[007EED84]"
		dm.AsmAdd "mov eax,[eax+0000012C]"
		dm.AsmAdd "call 004318F0"
		dm.AsmCall hwnd, 1
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
	
	Public function learnSkill(code)
		dm.AsmClear 
		dm.AsmAdd "mov eax,0" & code
		dm.AsmAdd "call 00440780"
		dm.AsmCall hwnd, 1
	end Function
	
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
	
	' 登入
	Public Function login(id,pwd)
		dm_ret = dm.WriteString(hwnd,"<tthola.dat+3E97E4>",0,id)
		dm_ret = dm.WriteString(hwnd,"<tthola.dat+36C2C0>",0,pwd)
		dm.AsmClear 
		dm.AsmAdd "push 007E97E4"
		dm.AsmAdd "mov esi,0076C2C0"
		dm.AsmAdd "call 00441390"
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

	' ??
	Public Function hasMonster()
		hasMonster = read("<tthola.dat>+3EB930")<>0
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
	
End Class
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
MacroID=ea432741-492b-429e-b5a6-192926d42b51
Description=TEST
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=1
QUI=Form1
[Relative]
SetupOCXFile=
[Comment]

[UIPackage]
UEsDBBQAAgAIAKweg1P/Sm75uwMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAfClKlhwpSpYcKUqWHtWUtPU0EUPrfl0ZZaWoTyFMtLkZXujYklEjGIhlajRk0KViH2gQUN7vwZblm7coGJP8C17tziyrhzxU7qN/fOtVMEvGfmQm3CRw5T2s79Zs6Z8xqC5ODL58T25vuBb7QHVyhIu9UwtSnvWVJsxIkC8u/darXqvl09QVPhF6RF2lDYuhXSDglDQpAI5BSkAxKFxBzTUyckIeftnqixabFAZfysU4quUQljhV4TB0mcGPdZ1j+++/PFw7kP379aQbx+E3beu0OzdJH0ESLLcvkDB3xn+4bD6/Krn01TluYMVhCR/EnpO4fxu6P6WRb6XqU8zVOOihi56KKAzZ+QsdfrvBY5uv77v8kJjgcz5UrRwP1IzL9kMN8iM5jON8XW5qcfJvtnuOyRzNeNO7X4F6ir+bzazI2VIuvkaJEKmmuIIf51yDrFK39AiX93aYXWILoriGvsP6jwL4A1R0/oFvRQYOZeN/5H5TO98rco/FnwbyDvG9jf4u6/VbH/NKqPAqSiuZJuDf42Zf81/jTssETP+fx2vdzJ0H/7HvsvGepf1OchBn+InBpexSM4xQ41H9obzL+zv87ttHQzk8rkSmupTL6y8vSI+CfePTaa/wCRp4L4k8dvHYj42yN9yuv5Cyvnfw6+x/e6+vMfl8/zyh/Z1/8zWMcy9MDLhn3w/9Mk4rB3/g6F/yqq/1U7B1QgJayBl4WS2H+YnL7cK39U4Z/B3ktm8Yed/0S8uvxn/wVk32dYg9B7SSf+s/ljdfbPQf/rWIPQQ9q2g5sJamfjYExq5J9OhT9tP9/xPy98+5w/q5tq925e5sQVftP4YVp/mmK6XFwsp8sbpvN1a/hm71+yOPnrdu2rc/rs+MP2v4Ry/m6DfwPsBVThOohp8HfV+b+4/9Hv4CKwoMvdf8B3Drt/8uP+5xzT/0W+CsnXjej/RLyS13++9H9nMQ6T9/zXQ772f1aKqf8k+dv/DWDsZey/l/zv/yYY/Or/i/zq/6YY/H3kb/83iHGMwd9P/vZ/IxiHGPwD9Hf/pwvT/H8Pch/yEfJWY36j82ewwfym/acP/R87/g+Sv/3fuPRpr/xD5G//N4nxPIP/DPnb/4nnjTL4h6mx/Z/I127/NwtuofM1Tf6oBn+KnDtggYxdezoxOA8/4KJLg39E0f+8ffaKsH1Bu//j8o8q/KaoGiYA0/5zamvZaP7gZpYauf5G4zrO2yvt2wcn/lwgXv07tif+LNFL+P/x9Z/jCv8s2EUE1GPXu3+a8NH/iIjNr+I3UEsBAhcLFAACAAgArB6DU/9Kbvm7AwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHwpSpYVBLBQYAAAAAAQABAEAAAADzAwAAAAA=


[Script]
Set dm = createobject("dm.dmsoft")
//Import "Tthbn.vbs" : Set t = New Tthbn
//hwnd = dm.FindWindow("", "絕代方程式")
Import "Tthol.vbs" : Set t = New Tthol
hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
t.init (hwnd)

VBSBegin


keyword = "級"
Set bag = t.getItems("bag")
For i = 0 To bag.Count - 1
	Set item = bag.GetByIndex(i)	
	If instr(item.item("name"), keyword) > 0 Then 
		TracePrint item.item("name")
		push item
	End If
Next

Function pop(item)
	id = item.item("id")
	sn = item.item("sn")
	amount = item.item("amount")
	
	dm.AsmClear 
	dm.AsmAdd "sub esp,16"
	dm.AsmAdd "mov ecx,0" & hex(id) & "0007"
	dm.AsmAdd "mov [esp],ecx"
	dm.AsmAdd "mov ecx,0" & hex(sn) & "0000"
	dm.AsmAdd "mov [esp+4],ecx"
	dm.AsmAdd "mov cx,0" & hex(amount)
	dm.AsmAdd "mov [esp+10],cx"
	dm.AsmAdd "lea ebx,[esp]"
	dm.AsmAdd "mov ecx,034"
	dm.AsmAdd "push 0"
	dm.AsmAdd "push 0B"
	dm.AsmAdd "push 0"
	dm.AsmAdd "call 0043EAB0"
	dm.AsmAdd "add esp,16"
	dm.AsmAdd "ret"
	dm.AsmCall hwnd, 1
End Function

Function push(item)
	id = item.item("id")
	sn = item.item("sn")
	amount = item.item("amount")
	
	dm.AsmClear 
	dm.AsmAdd "sub esp,16"
	dm.AsmAdd "mov ecx,0" & hex(id * &H10000 + 7)
	dm.AsmAdd "mov [esp],ecx"
	dm.AsmAdd "mov ecx,0" & hex(sn * &H10000)
	dm.AsmAdd "mov [esp+4],ecx"
	dm.AsmAdd "mov cx,0" & hex(amount)
	dm.AsmAdd "mov [esp+10],cx"
	dm.AsmAdd "lea ebx,[esp]"
	dm.AsmAdd "mov ecx,032"
	dm.AsmAdd "push 01"
	dm.AsmAdd "push 0B"
	dm.AsmAdd "push 0"
	dm.AsmAdd "call 0043EAB0"
	dm.AsmAdd "add esp,16"
	dm.AsmAdd "ret"
	dm.AsmCall hwnd, 1
End Function

Function npc()
	dm.AsmClear 
	dm.AsmAdd "sub esp,16"
	dm.AsmAdd "mov byte ptr [esp],01"
	dm.AsmAdd "mov ecx,19300007"
	dm.AsmAdd "mov [esp+1],ecx"
	dm.AsmAdd "mov ecx,00040000"
	dm.AsmAdd "mov [esp+5],ecx"
	dm.AsmAdd "mov cx,0000"
	dm.AsmAdd "mov [esp+9],cx"
	dm.AsmAdd "lea ebx,[esp]"
	dm.AsmAdd "mov ecx,028"
	dm.AsmAdd "push 0"
	dm.AsmAdd "push 0B"
	dm.AsmAdd "push 0"
	dm.AsmAdd "call 0043EAB0"
	dm.AsmAdd "add esp,16"
	dm.AsmAdd "ret"
	dm.AsmCall hwnd, 1
End Function

VBSEnd
' ------------------------------跑馬燈------------------------------
//str = "高嘉瑜對林男提告傷害、妨害秘密、妨害電腦使用等罪，檢察官訊問後，認定林男另涉嫌私行拘禁、及強制罪，檢方認定犯罪嫌疑重大。"
//scroller str,10

Function scroller(str, length)
	slen = len(str)
	str = str & mid(str,1,length)
	For i = 0 To slen
		ad (mid(str, i + 1, length))
		i = i mod slen
		Delay 300
	Next
End Function


Function ad(str)
	dm_ret = dm.WriteString(hwnd, "[[<tthola.dat>+003ED978]+10]+35A8", 0, str)
	dm.AsmClear 
	dm.AsmAdd "push 0"
	dm.AsmAdd "push 00000029"
	dm.AsmAdd "push 0"
	dm.AsmAdd "lea ebx,[esp+0C]"
	dm.AsmAdd "mov ecx,07F"
	dm.AsmAdd "call 0043EAB0"
	dm.AsmAdd "ret"
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

//hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4 + 8 + 16), ",")
//For Each hwnd In hwnds
//	Call dm.WriteData(hwnd, "<ttha.bin>+215D7", "8097AE02")
//Next
// ------------------------------TEST------------------------------
// 取得小方家族名單
//hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4 + 8 + 16), ",")
//For Each hwnd In hwnds
//	hwnd = 1901288
//	base = dm.GetModuleBaseAddr(hwnd, "tthbn.bin") + &HA9E80
//	family = readString(base, 16)
//	count = dm.ReadInt(hwnd, HEX(base + 40), 1)
//	For i = 0 To count-1
//		name = readString(base + 60 + i * 44, 16)
//		level = dm.ReadInt(hwnd, HEX(base + 76 + i * 44), 1)
//		job = dm.ReadInt(hwnd, HEX(base + 78 + i * 44), 1)
//		loginTime = dm.ReadInt(hwnd, HEX(base + 84 + i * 44), 0)
//		off =  DateDiff("d", secToDate( loginTime + 8 * 60 * 60), now)
//		str = name & "|" & level & "|" & job & "|" & off & "|" & family
//		dm.WriteFile "C:\Users\god\Desktop\123.txt", str + chr(10)
//	Next
//Next

// send msg
//hwnd = dm.FindWindow("", "Tthol")
//
//dm.AsmClear
//dm.AsmAdd "mov eax,00000005"
//dm.AsmAdd "mov ebx,0019F784"
//dm.AsmAdd "mov ecx,0000001C"
//dm.AsmAdd "mov edx,042D69A8"
//dm.AsmAdd "mov esi,17024A50"
//dm.AsmAdd "mov edi,1DB28568"
//dm.AsmAdd "mov ebp,17024780"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 00000000"
//dm.AsmAdd "push 77A677A6"
//dm.AsmAdd "push 042D69A8"
//dm.AsmAdd "push 00453253"
//dm.AsmAdd "push -1"
//dm.AsmAdd "push 00000005"
//dm.AsmAdd "call 004309D0"
//dm.AsmCall hwnd,1

Function go(xy)
	ProcessId = 11016
 	lpBaseAddress = Lib.AsmCode.申請指定進程空間(ProcessId, 120)
	call Lib.AsmCode.寫入四字節內存整數(ProcessId, lpBaseAddress + 100, 0)
	call Lib.AsmCode.AsmClear()
	call Lib.AsmCode.Mov_EAX_Value(t.getMainAddr())
	call Lib.AsmCode.Mov_EBX_Value(xy(0) * 40)
	call Lib.AsmCode.Mov_EDI_Value(xy(1) * 40)
	call Lib.AsmCode.Push(1)
	call Lib.AsmCode.Mov_ECX_Value(&H00492F40)
	call Lib.AsmCode.Call_ECX()
	call Lib.AsmCode.Mov_EAX_Value(1)
	call Lib.AsmCode.Mov_DWORD_Ptr_Addr_EAX(lpBaseAddress+100)
	call Lib.AsmCode.Mov_ESP_EBP()
	call Lib.AsmCode.pop_ebp()
	call Lib.AsmCode.Ret()
	call Lib.AsmCode.RunAsmCodetoMainThread(ProcessId,lpBaseAddress)
End Function

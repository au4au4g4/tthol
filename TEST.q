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
MacroID=59f289d0-0ce9-47a4-9914-0d1525d37be7
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
UEsDBBQAAgAIAG4UglN1XzTCvgMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAfrMKhh6zCoYeswqGHtWUtPU0EUPrfl0ZZaWoTyFMtLkZXGrTGxRCIJoqHVqInGglWIfZCCBndu/A9uWevGBQt/gGvZuVUXJv4BdlK/uXeunSLgPTMXahM+cpjSdu43c86c1xAkB9ufE183Pwx8oz24SkHarYapTXnPkmIjThSQf+9Wq1X37eoJmgq/IC3ShsLWrZB2SBgSgkQgpyAdkCgk5pieOiEJOW/3RI1NiwUq42edUnSdShgr9Io4SOLEuM+y/vHd1ffbl7//+GIF8fp12HnvDs3SRdJHiCzL5Q8c8J3Hbxxel1/9bJqyNGewgojkT0rfOYzfHdXPstD3KuVpnnJUxMhFFwVs/oSMvV7ntcjR9d//TU5wPJgpV4oG7kdi/iWD+RaZwXS+KbY2P/002T/DZY9kvm7cqcW/QF3N59VmbqwUWSdHi1TQXEMM8a9D1ile+QNK/LtLK7QG0V1BXGP/QYV/Aaw5ekK3oIcCM/e68T8qn+mVv0Xhz4J/A3nfwP4Wd/+tiv2nUX0UIBXNlXRr8Lcp+6/xp2GHJXrO57fr5U6G/tv32H/JUP+iPg8x+EPk1PAqHsIpdqj50N5g/p39dW6npZuZVCZXWktl8pWVp0fEP/HukdH8B4g8FcSfPH7rQMTfHulTXs9fWDn/c/A9vtfVn/+4fJ5X/si+/p/BOpahB1427IP/nyYRh73zdyj811D9r9o5oAIpYQ28LJTE/sPk9OVe+aMK/wz2XjKLP+z8J+LVlT/7LyD7PsMahN5LOvGfzR+rs38O+l/HGoQe0rYd3ExQOxsHY1Ij/3Qq/Gn7+Y7/eeHb5/xZ3VS7d/MyJ67wm8YP0/rTFNPl4mI5Xd4wna9bwzd7/5LFyV+3a1+d02fHH7b/JZTzdxv8G2AvoArXQUyDv6vO/8X9j34HF4EFXe7+A75z2P2TH/c/55j+L/JVSL5uRP8n4pW8/vOl/zuLcZi8578e8rX/s1JM/SfJ3/5vAGMvY/+95H//N8HgV/9f5Ff/N8Xg7yN/+79BjGMM/n7yt/8bwTjE4B+gv/s/XZjm/3uQ+5CPkLca8xudP4MN5jftP33o/9jxf5D87f/GpU975R8if/u/SYznGfxnyN/+TzxvlME/TI3t/0S+dvu/WXALna9p8kc1+FPk3AELZOza04nBefgBF10a/COK/ufts1eE7Qva/R+Xf1ThN0XVMAGY9p9TW8tG8wc3s9TI9TcaN3DeXmrfPjjx5wLx6t+xPfFniV7A/4+v/xxX+GfBLiKgHrve/dOEj/5HRGx+Fb8BUEsBAhcLFAACAAgAbhSCU3VfNMK+AwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAH6zCoYVBLBQYAAAAAAQABAEAAAAD2AwAAAAA=


[Script]
Set dm = createobject("dm.dmsoft")

//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//Import "Tthol.vbs"
//Set t = New Tthol
//t.init (hwnd)
'783890 1111532 652856 914934 783898 915088 589498 652808 327518 718388 524134 524142 6160244 261862 456386 652882

Call dm.setwindowstate(652882,3)
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

//While true
//	go (array(50, 37))
//	Delay 2000
//	go (array(42, 40))
//	Delay 2000
//Wend


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


// ------------------------------其他------------------------------
Function secToDate(numSeconds)
  Dim dEpoch
  dEpoch = DateSerial(1970,1,1)  
  secToDate = DateAdd("s",numSeconds,dEpoch) 
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

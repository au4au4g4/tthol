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
MacroID=3781a671-1216-431f-be63-fdde4c1be28d
Description=開小方
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
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Import "Tthbn.vbs" : Set t = New Tthbn
Set re = New RegExp
Set dm = createobject("dm.dmsoft")

w = dm.getscreenwidth()
h = dm.getscreenheight() - 40
c = 5
r = 2

three = array(110, 3, 1, array("SnowFlakes", "FrozenBubbles", "SnowRollers","safedepositbox"))
sky = array(110, 3, 1, array("tongxia", "xingzhi", "cangcui", "fenxue", "qiuyun", "yaolan", "yuanxiang", "yunmiao", "guiyan", "daohua"))
fox  = array(110, 3, 1, array("ef7a82","eaff56","ca6924","a4e2c6","b0a4e3","e4c6d0","e3f9fd","afdd22","a2cc8c","f1c143"))
kylin  = array(110, 4, 1, array("tzuyeh","rihdan","posiao","jihhsi","riiyuu","jhongrih","yangrih","chenrih","puurih","renjing"))
bad =  array(110, 2, 1, array("Hydrogenn", "Nitrogenn", "Oxygenn", "Fluorine", "Helium", "Neonnn", "Argonn", "Krypton", "Xenonn", "Radonn"))
flower =  array(250, 10, 1, array("hwasheng","jhihma","lianru","heitang","fongni","hungtou","mochaa","cisihh","haitai","chaiyu"))
teams = array(fox)
'teams = array(kylin, flower, three)

hwnd = 0
windowCnt = UBound(t.getAllHwnds())
For j = 0 To UBound(teams)
	team = teams(j)
	For i = 0 To UBound(team(3))
		// 開程式
		zmRunApp("C:\Users\god\Desktop\JD_FCS2110\tthfcs.exe")
		While windowCnt >= UBound(t.getAllHwnds())	
			Delay 500
		Wend
		windowCnt = windowCnt + 1

		// 調整位置
		hwnd = dm.FindWindow("", "絕代方程式")
		dm_ret = dm.SetClientSize(hwnd, w / c - 2, h / r - 32)
		x = (i mod 5) * w / c - 7
		y = (i \ 5) * h / r
		dm.MoveWindow hwnd, x, y
		
		// 輸入帳密
		login = dm.FindWindow("","登錄")
		edits = split(dm.EnumWindow(login, "", "Edit", 2 + 4), ",")
		dm_ret = dm.BindWindow(edits(0), "normal", "windows", "windows", 0)
		call del(15)
		dm.SendString edits(0), team(3)(i)
		dm.SendString edits(1), "gj83dj4"
		comboBoxs = split(dm.EnumWindow(login, "", "ComboBox", 2 + 4), ",")
		dm_ret = dm.BindWindow(login, "normal", "windows", "windows", 0)
		dm.SendString comboBoxs(0), "飛雁山莊(花)"
		dm.SendString comboBoxs(1), "1"
		
		// 破解
		t.init hwnd
		Call crack()
		Call moveSpeed(team(0))
		Call atkRange(team(1))
		Call atkSpeed(400)
		Call addHP()
		Call addMP()
		Call defBuff(team(2))
		't.login
	Next
	
	// 換頁
	KeyDown "Ctrl", 1
	KeyDown "Win", 1
	KeyPress "Right", 1
	KeyUp "Win", 1
	KeyUp "Ctrl", 1
Next

Function del(cnt)
	Dim i
	For i = 0 To cnt
		dm.KeyPress 8
		dm.KeyPress 46
	Next
End Function

Function crack()
	Call dm.WriteData(hwnd, "<tthbn.bin>+ACD08", "E1 E7 F6 B7 84 AC ED B4 E1 E0 A0 9C A9 FB ED E6 B6 8B B1")
	reDim codes(0) : codes(0) = "jmp ttha.bin+3B1E0"
	Call asm("ttha.bin+3B162", codes)	
End Function

// 280
Function moveSpeed(speed)
	Call dm.WriteInt(hwnd,"<ttha.bin>+3FD53",0,speed)
End Function

Function atkRange(range)
	reDim codes(5)
	codes(0) = "mov eax,0" & HEX(range ^ 2)
	codes(1) = "cmp edx,eax"
	codes(2) = "jg ttha.bin+535AA"
	codes(3) = "cmp ecx,eax"
	codes(4) = "jg ttha.bin+535AA"
	codes(5) = "jmp ttha.bin+535AF"
	Call asm("ttha.bin+5358C", codes)
End Function

Function atkSpeed(speed)
	Call dm.WriteInt(hwnd, "[[<ttha.bin>+4EF82C]+98]+12100", 0,speed)	
End Function

Function addHP()
	Call dm.WriteInt(hwnd, "<ttha.bin>+22F5F", 0, &H20A)
	Call dm.WriteData(hwnd, "<ttha.bin>+22E5F", "90 90 90 90 90 90")
End Function

Function addMP()
	// 安裝變身
	Call dm.WriteData(hwnd, "<tthbn.bin>+A1580", "0C 02 00 00 00 00 00 00 A8 AD 00 00 00 00 00 00 00 00 00 00 00 00 00 00")
	Call dm.WriteData(hwnd, "<ttha.bin>+2293C", "D6 15 42 00")
	Call dm.WriteData(hwnd, "<ttha.bin>+1EF20", "90 90")
	Call dm.WriteData(hwnd, "<ttha.bin>+1EF3D", "90 90 90 90")
	Call dm.WriteData(hwnd, "<ttha.bin>+21614", "90 90 90")
	
	// 導入原本流程
	reDim codes(2)
	codes(0) = "and eax, 0F"
	codes(1) = "jmp ttha.bin+213BF"
	codes(2) = "nop"
	Call asm("ttha.bin+213B6", codes)
	
	// 50%變身
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

Function defBuff(have8)
	// 安裝替身
	Call dm.WriteData(hwnd, "<tthbn.bin>+A1568", "0B 02 00 00 00 00 00 00 B4 C0 A8 AD 00 00 00 00 00 00 00 00 00 00 00 00")
	Call dm.WriteData(hwnd, "<ttha.bin>+1E9B1", "90 90")
	Call dm.WriteData(hwnd, "<ttha.bin>+1E9CE", "90 90 90 90")
	
	// 導入原本流程
	reDim codes(2)
	codes(0) = "and eax, 0F"
	codes(1) = "jmp ttha.bin+213BF"
	codes(2) = "nop"
	Call asm("ttha.bin+213B6", codes)
	
	If have8 Then 
		// 有八卦才放替身
		reDim codes(5)
		codes(0) = "mov eax,[tthbn.bin+A171C]" // 有八卦1
		codes(1) = "add eax,eax"
		codes(2) = "add eax,[ebx+04]"		// 沒替身0
		codes(3) = "cmp eax,02" 			// 1+1+0
		codes(4) = "jne ttha.bin+21461"
		codes(5) = "jmp ttha.bin+213DD"
		Call inAsm("ttha.bin+213C6", codes)
		
		// 替身消失時 清空八卦
		reDim codes(4)
		codes(0) = "cmp ecx,00000858"// 是清空替身流程
		codes(1) = "jne newmem+12"
		codes(2) = "mov dword ptr [tthbn.bin+A171C],0"// 清空八卦
		codes(3) = "mov dword ptr [ecx+tthbn.bin+A0D14],0"
		codes(4) = "jmp tthbn.bin+421B6"
		Call inAsm("tthbn.bin+421AC", codes)		
	End If
End Function

Function asm(addr, codes)
	Dim code
	Call dm.AsmClear
	For Each code In codes	
		dm.AsmAdd toNum(code)
	Next
	code = dm.AsmCode(clng("&H" & toNum(addr)))
	Call dm.WriteData(hwnd, addBracket(addr), code)
End Function

Function toNum(str)
	Dim addr,num,match
	re.Pattern = "\w+\.\w+\+\w+"
	set match = re.Execute(str)
	If (match.count > 0) Then 
		addr = split(match(0), "+")
		num = dm.GetModuleBaseAddr(hwnd, addr(0)) + clng("&H" & addr(1))
	end if 
	toNum = re.Replace(str, HEX(num))
End Function

Function addBracket(str)
	re.Pattern = "(\w+\.\w+)"
	addBracket = re.Replace(str, "<$1>")
End Function

Function inAsm(addr, codes)
	Dim pid,hProcess,newmem,inCodes(1),i
	pid = dm.GetWindowProcessId(hwnd)
	hProcess = OpenProcess(2035711, False, pid)
	newmem = HEX(VirtualAllocEx(hProcess, 0, 100, &H3000, &H40))
		
	inCodes(0) = "jmp 0" & newmem
	inCodes(1) = "nop"
	Call asm(addr, inCodes)
	
	For i = 0 To UBound(codes)
		codes(i) = inToNum(codes(i),newmem)
	Next
	Call asm(newmem, codes)
End Function

Function inToNum(str,newmem)
	Dim addr,num,match
	re.Pattern = "newmem\+\w+"
	set match = re.Execute(str)
	If (match.count > 0) Then 
		addr = split(match(0), "+")
		num = clng("&H" & newmem) + clng("&H" & addr(1))
	end if 
	inToNum = re.Replace(str, "0"& HEX(num))
End Function

Sub zmRunApp(path)
    Dim p, DirPath, FileName
    p = InStrRev(path, "\")
    DirPath = Left(path, p)
    FileName = Right(path, Len(path) - p)
    ShellExecute GetDesktopWindow, "open", FileName, vbNullString, DirPath, 5
End Sub

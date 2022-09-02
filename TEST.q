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
MacroID=c559cb34-2c95-4b80-8095-4fa3cd169748
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
UEsDBBQAAgAIAO6TGVW0ZYi5sgMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAcbwAdjG8AHYxvAB2PtWUtPU0EUPrcFaUstLUJ5iuWlyEr3xsQSiSSIhlajJpoUrELsgxQ0uPNnuGXjxpULEt25ca0bE5e6cOEfYCf1m3vn2ilCuGfmwrUJHzlMaTv3mzlnzmsIk4Mvn1Pft94N/qA9uEph2q1H6ZTyniXFRpIoJP/erdfr7tv1E7QUfkPapA2FrdshHZAoJAKJQU5DOiFxSMIxPXVBUnLe7okaWxaLVMXPBmXoOlUw1uglcZDGiXGfZR3y3Y9vvv4sv/9mhfH6VdR57w7N0SXSR4Qsy+UPHcLrjupnM5SneYMVxCR/WvoOlz8Pfa9RkRaoQGWMXHRTyOZPydjrdV6bHF3//d/kBMeD2WqtbOB+JOZfNphvkRlM55tie+vTL5P9M1z2SObrxp1G/As11XxebebGSpF1CrREJc01JBD/OmWd4pU/pMS/u7RK6xDdFSQ19h9W+BfBWqDHdAt6KDFzrxv/4/KZXvnbFP48+DeR9w3sb3H3367YfwbVRwlS01xJjwb/KWX/Df4s7LBMz/j8dr3cxdB/xx77LxvqX9TnEQZ/hJwaXsVDOMUOtR46Aubf2V/ndlq6mcvkCpX1TK5YW31yRPyTbx8ZzX+AyFND/Cnitw5E/O2VPuX1/EWV8z8P3+N7XfP5T8rneeWP7ev/OaxjBXrgZcN++P8ZEnHYO3+nwn8N1f+anQNqkArWwMtCaew/Sk5f7pU/rvDPYu8Vs/jDzn8iXl35u/8Ssu9TrEHovaIT/9n8iSb7F6D/DaxB6CFr28HNBI2zcTCmNPJPl8KftZ/v+J8Xvn3On9VDjXs3L3OSCr9p/DCtP00xUy0vVbPVTdP5ujV8q/cveZz8Dbv21Tl9dvxh+19KOX+3wb8J9hKqcB0kNPi7m/xf3P/od3AxWNDlHjjgO0d9/3Oe6f8iX0Xk6yD6PxGv5PWfL/3fOYwj5D3/9ZKv/Z+VYeo/Tf72f4MY+xj77yP/+79JBr/6/yK/+r9pBn8/+dv/DWEcZ/APkL/93yjGYQb/IP3b/+nCNP/fg9yHfIC81pgfdP4MB8xv2n/60P+x4/8Q+dv/TUif9so/TP72f1MYLzD4z5K//Z943hiDf4SC7f9Evnb7vzlwC52va/LHNfgz5NwBC+Ts2tOJwUX4ARfdGvyjiv4X7LNXhu1L2v0fl39M4TdF3TABmPaf09srRvOHtvIU5PqDxg2ctxfatw9O/LlIvPp3fE/8Wabn8P/j6z8nFP45sIsIqMeud/806aP/ERGbX8UfUEsBAhcLFAACAAgA7pMZVbRliLmyAwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHG8AHY1BLBQYAAAAAAQABAEAAAADqAwAAAAA=


[Script]
//Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
Set dm = createobject("dm.dmsoft")
hwnd = dm.FindWindow("", "絕代方程式")
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
//t.init hwnd

For 10000
	Beep 
	Delay 1000
Next

//result = dm.FindData(hwnd, "00000000-FFFFFFFF", "51 E8 C2 BD")
//TracePrint result


//Set addrs = CreateObject("Scripting.Dictionary")
//addrs.Add "login", "8B C3 6B C0 3C"
//addrs.Add "send", "80 3D 34 C8 7E 00 01 75 6A 8B 44 24 08 56 50 51 E8 3B"
//addrs.Add "send1", "80 3D 34 C8 7E 00 01 75 6A 8B 44 24 08 56 50 51 E8 4B"
//addrs.Add "go", "56 8B F0 8B 86 04 03 00  00 85 C0 74"
//addrs.Add "atk", "83 EC 08 53 8B 5C 24 10 8B"
//addrs.Add "learnSkill", "51 53 6A 04 51 8D 5C 24  0C B9 3C"
//addrs.Add "speed", "7D 05 B9 64"
//For Each key In addrs.Keys
//	result = dm.FindData(hwnd, "00400000-00500000", addrs.item(key))
//	dm.WriteIni "addr", key, CLNG("&H" & result), ".\QMScript\tthbn.ini"
//Next

//arr = split(dm.EnumWindow(0, "挂機路線", "Static", 1 + 2), ",")
//For Each a In arr
//	a = clng(dm.GetWindow(clng(a),0))
//	hwnd = clng(dm.GetWindow(a, 7))
//	TracePrint HEX(dm.ReadInt(hwnd, "[[[<ttha.bin>+006FDE48]+C]+8]+14", 0)+4)
//	dm.AsmClear 
//	dm.AsmAdd "push 0" & HEX(a)
//	dm.AsmAdd "mov ecx,0" & HEX(dm.ReadInt(hwnd, "[[[<ttha.bin>+006FDE48]+C]+8]+14", 0)+4)
//	dm.AsmAdd "call 0" & HEX(dm.GetModuleBaseAddr(hwnd, "ttha.bin") + &H888C5)
//	TracePrint HEX(dm.AsmCall(hwnd, 1))
//Next

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

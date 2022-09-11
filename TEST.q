[General]
SyntaxVersion=2
BeginHotkey=49
BeginHotkeyMod=2
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
UEsDBBQAAgAIAKANKFVSAtCpugMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAc8SRljPEkZYzxJGWPtWUtPU0EUPrfl0ZelRShPtbwUWenemFgikQTR0GrURJOCVYh9kIIGd/4El25Zu3LBwp0xcS3RjUtJ3PgH2En95t65dooQ75m5UJvwkcOUtnO/mXPmvIYgOdj+nPy++W5wh/bhKgVprxamDuU9S4qNBFFA/r1Xq9Xct2snaCn8grRJGwpbt0M6IWFICBKBnIJEITFI3DE9dUGSct7eiRpbFgtUwc86pek6lTFW6SVxkMKJcZ9l/eO729Gvr3c+frOCeP0q7Lx3h2bpEukjRJbl8gcO+c6HH19sXpdf/WyacjRnsIKI5E9J3+Hy56DvVSrQPOWphJGLbgrY/EkZe73Oa5Oj67//m5zgeDBTqZYM3I/E/MsG8y0yg+l8U2xtfvppsn+Gyx7JfN24U49/gYaaz6vN3Fgpsk6eFqmouYY44l9U1ile+QNK/LtLK7QG0V1BQmP/QYV/Aax5eky3oIciM/e68T8mn+mVv03hz4F/A3nfwP4Wd//tiv2nUX0UIVXNlfRo8Hco+6/zZ2CHJXrG57fr5S6G/jv32X/JUP+iPg8x+EPk1PAqHsIpdqn10Nlk/t2DdW6npZvZdDZfXktnC9WVJ0fEP/H2kdH8B4g8VcSfAn7rQMTfXulTXs9fWDn/c/A9vtc1nv+EfJ5X/siB/p/FOpahB1427If/nyYRh73zRxX+a6j+V+0cUIWUsQZeFkph/2Fy+nKv/DGFfwZ7L5vFH3b+E/Hqyp/9F5F9n2INQu9lnfjP5o832D8P/a9jDUIPGdsObiaon43DMamRf7oU/oz9fMf/vPAdcP6sHqrfu3mZk1D4TeOHaf1piulKabGSqWyYztet4Vu9f8nh5K/bta/O6bPjD9v/ksr5uw3+DbAXUYXrIK7B393g/+L+R7+Di8CCLvfAId9x732O6v7nPNP/Rb4KydfN6P9EvJLXf770f+cwniXv+a+XfO3/rDRT/ynyt/8bxNjH2H8f+d//TTD41f8X+dX/TTH4+8nf/m8I4xiDf4D87f9GMA4z+Afp7/5PF6b5/x7kPuQ95I3G/Gbnz2CT+U37Tx/6P3b8HyJ/+79x6dNe+YfJ3/5vEuMFBv8Z8rf/E88bZfCfpeb2fyJfu/3fLLiFztc0+WMa/Gly7oAFsnbt6cTgAvyAi24N/hFF//P22SvB9kXt/o/LP6rwm6JmmABM+8+prWWj+UObOWrm+puNGzhvL7RvH5z4c5F49e/YvvizRM/h/8fXf44r/LNgFxFQj13v/mnCR/8jIja/it9QSwECFwsUAAIACACgDShVUgLQqboDAAAIJgAACQAJAAAAAAAAAAAAAIAAAAAAVUlQYWNrYWdlVVQFAAc8SRljUEsFBgAAAAABAAEAQAAAAPIDAAAAAA==


[Script]
Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
//Set dm = createobject("dm.dmsoft")
//hwnd = dm.FindWindow("", "絕代方程式")
//Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
//t.init hwnd

MsgBox 1


//TracePrint dm.FindData(hwnd, "00400000-00600000", "74 13 8B 06 8B CE FF 50 70")
//Set addrs = CreateObject("Scripting.Dictionary")
//addrs.Add "login", "8B C3 6B C0 3C"
//addrs.Add "send", "80 3D 34 C8 7E 00 01 75 6A 8B 44 24 08 56 50 51 E8 3B"
//addrs.Add "send1", "80 3D 34 C8 7E 00 01 75 6A 8B 44 24 08 56 50 51 E8 4B"
//addrs.Add "go", "56 8B F0 8B 86 04 03 00  00 85 C0 74"
//addrs.Add "atk", "83 EC 08 53 8B 5C 24 10 8B"
//addrs.Add "learnSkill", "51 53 6A 04 51 8D 5C 24  0C B9 3C"
//addrs.Add "speed", "7D 05 B9 64"
//addrs.Add "logout", "6A FF 68 7F 2A"
//addrs.Add "wear2", "53 55 56 8B D8 57 8D"
//addrs.Add "shopping", "83 EC 0C 53 8D 88"
//addrs.Add "freeWindow", "74 13 8B 06 8B CE FF 50 70"
//addrs.Add "freeWindowLimit", "0F 84 81 04 00 00 E8"
//For Each key In addrs.Keys
//	result = dm.FindData(hwnd, "00400000-00600000", addrs.item(key))
//	dm.WriteIni "addr", key, CLNG("&H" & result), ".\QMScript\tthol.ini"
//Next
//
//Set addrs = CreateObject("Scripting.Dictionary")
//addrs.Add "location", "4B 20 8B 0D"
//addrs.Add "shop", "53 75 15 A1"
//For Each key In addrs.Keys
//	result = dm.FindData(hwnd, "00400000-00500000", addrs.item(key))
//	l = Len(Replace(addrs.item(key), " ", "", 1, - 1 )) / 2
//	result = dm.ReadInt(hwnd, HEX(CLNG("&H" & result) + l), 0)
//	dm.WriteIni "addr", key, result, ".\QMScript\tthol.ini"
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

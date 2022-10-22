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
UEsDBBQAAgAIAKoSV1XoO0hPrwMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAdBpVRjQaVUY0GlVGPtWUtPU0EUPrcFaUstLUJ5iuWlyEr3xsQ2EkkQDa1GTTQpWIXYBylocKU/wy1rVy5Y+ANc69Ilbkz8A+ykfnPvXDtFCPfMXLg24SOHKW3nfjPnzHkNYXLw7Wtqd/vT8A86gJsUpv1GlM4p71lSbCSJQvLv/Uaj4b7dOENb4TekQ9pQ2LoT0gWJQiKQGOQ8pBsShyQc01MPJCXn7Z+psW2xRDX8bFKGblMVY53eEgdpnBj3WdYx391915jJ/fxuhfH6fdR57wHN0zXSR4Qsy+UPHcPrjupnOSrQgsEKYpI/LX2Hy1+AvtepRItUpApGLnopZPOnZOz1Oq9Djq7//m9yhtPBXK1eMXA/EvOvG8y3yAym802xs/3ll8n+GS57IvN1404z/oVaaj6vNnNjpcg6RVqmsuYaEoh/3bJO8cofUuLfQ1qjDYjuCpIa+w8r/EtgLdJzugc9lJm5143/cflMr/wdCn8B/FvI+wb2t7j771Tsn0P1UYbUNVfSp8F/Ttl/kz8LO6zQKz6/XS/3MPTfdcD+K4b6F/V5hMEfIaeGV/EUTrFH7YeugPn3Dte5nZbu5jP5YnUjky/V116cEP/0x2dG858g8tQRf0r4rQMRf/ulT3k9f1Hl/C/A9/he13r+k/J5Xvljh/p/HutYhR542XAQ/n+BRBz2zt+t8N9C9b9u54A6pIo18LJQGvuPktOXe+WPK/xz2HvVLP6w85+IVzf+7r+M7PsSaxB6r+rEfzZ/osX+Reh/E2sQesjadnAzQfNsHI0ZjfzTo/Bn7ec7/ueF75DzZ/VR897Ny5ykwm8aP0zrT1PkapXlWra2ZTpft4Zv9/6lgJO/ade+OqfPjj9s/0sp5+8++LfAXkYVroOEBn9vi/+L+x/9Di4GC7rcQ0d856Tvfy4z/V/kq4h8HUT/J+KVvP7zpf+7hHGMvOe/fvK1/7MyTP2nyd/+bxjjAGP/A+R//zfN4Ff/X+RX/zfL4B8kf/u/EYyTDP4h8rf/G8c4yuAfpn/7P12Y5v9HkMeQz5APGvODzp/hgPlN+08f+j92/B8hf/u/KenTXvlHyd/+bwbjFQb/RfK3/xPPm2Dwj1Gw/Z/I127/Nw9uofMNTf64Bn+GnDtggbxdezoxuAQ/4KJXg39c0f+iffYqsH1Zu//j8k8o/KZoGCYA0/5zdmfVaP7IdoGCXH/QuIPz9kb79sGJP1eJV/9OHog/K/Qa/n96/eeUwj8PdhEB9dj17p+mffQ/ImLzq/gDUEsBAhcLFAACAAgAqhJXVeg7SE+vAwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHQaVUY1BLBQYAAAAAAQABAEAAAADnAwAAAAA=


[Script]
//Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
Set dm = createobject("dm.dmsoft")
hwnd = dm.FindWindow("", "絕代方程式")
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
t.init hwnd

TracePrint t.account

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

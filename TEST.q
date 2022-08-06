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
UEsDBBQAAgAIAOBgBlWQM+ljrwMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAeEWe5ihFnuYoRZ7mLtWUtPU0EUPrfl0ZZSWoTyFMtLgZXujYklEjGIhlajJpoUrELsgxQ0uPNnuGVpXLlg4Q9wLUuXuvQPsJP6zb1z7RQh3jNz4dqEjxymtJ37zZwz5zWEycH+19T33U/DP+gIblCYDutR6lDes6TYSBKF5N+H9Xrdfbt+jpbCL0ibtKGwdTukExKFRCAxSDekCxKHJBzTUw8kJecdnquxZbFCVfxsU4ZuUQVjjd4SB2mcGPdZ1j++uz+73H3nwzcrjNfvos57D2iRrpI+ImRZLn/oH7zuqH42T3laMlhBTPKnpe9w+fPQ9yYVaZkKVMbIRS+FbP6UjL1e57XJ0fXf/03OcTZYqNbKBu5HYv41g/kWmcF0vin2dr/8NNk/w2VPZb5u3GnEv1BTzefVZm6sFFmnQKtU0lxDAvGvS9YpXvlDSvx7SBu0BdFdQVJj/2GFfwWsBXpO96CHEjP3uvE/Lp/plb9N4c+Dfwd538D+Fnf/7Yr951F9lCA1zZX0afB3KPtv8GdhhzV6xee36+Uehv47j9h/zVD/oj6PMPgj5NTwKp7CKQ6o9dAZMP/B8Tq309LdXCZXqGxlcsXaxotT4p/++Mxo/hNEnhriTxG/dSDib7/0Ka/nL6qc/yX4Ht/rms9/Uj7PK3/sWP/PYR3r0AMvGw7C/y+QiMPe+bsU/puo/jftHFCDVLAGXhZKY/9Rcvpyr/xxhX8Be6+YxR92/hPx6vqf/ZeQfV9iDULvFZ34z+ZPNNm/AP1vYw1CD1nbDm4maJyNkzGjkX96FP6s/XzH/7zwHXP+rD5q3Lt5mZNU+E3jh2n9aYr5anm1mq3umM7XreFbvX/J4+Rv27Wvzumz4w/b/1LK+bsP/h2wl1CF6yChwd/b5P/i/ke/g4vBgi730AnfOe37n8tM/xf5KiJfB9H/iXglr/986f8uYRwj7/mvn3zt/6wMU/9p8rf/G8Y4wNj/APnf/00z+NX/F/nV/80x+AfJ3/5vBOMkg3+I/O3/xjGOMviH6e/+Txem+f8R5DHkM+S9xvyg82c4YH7T/tOH/o8d/0fI3/5vSvq0V/5R8rf/m8F4hcF/kfzt/8TzJhj8YxRs/yfytdv/LYJb6HxLkz+uwZ8h5w5YIGfXnk4MLsIPuOjV4B9X9L9sn70ybF/S7v+4/BMKvynqhgnAtP+c21s3mj+ym6cg1x80buO8vdG+fXDizyzx6t/JI/FnjV7D/8+u/5xS+BfBLiKgHrve/dO0j/5HRGx+Fb8BUEsBAhcLFAACAAgA4GAGVZAz6WOvAwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHhFnuYlBLBQYAAAAAAQABAEAAAADnAwAAAAA=


[Script]
//Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
//Set dm = createobject("dm.dmsoft")
//hwnd = dm.FindWindow("", "絕代方程式")
//Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
//t.init hwnd

hwnds = t.getAllHwnds()
For each hwnd in hwnds
	t.init hwnd
	t.reLoginMin(4)
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

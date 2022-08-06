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
UEsDBBQAAgAIACZhBlVnt9bIrgMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAfvWe5i71nuYu9Z7mLtWUtPU0EUPrfl0ZZaWoTyFMtLkZXujYklEjGIhlajJpoUrELsgxQ0uPMvuHPL0rhywcIf4FqXLnXp3rCT+s29c+kUId4zc+HahI8cptx07jdzzpzXECYHX7+kvm9/HPpBB3CNwrRXj1KH8sySYiNJFJJ/79Xrdfdx/RQthd+QNmlDYet2SCckColAYpAzkC5IHJJwTE/dkJSct3eqxpbFElXxs0kZukEVjDV6TRykcWLcd1n/+vLbXwu33n+zwvj4Juo8ukfzdJn0ESHLcvlD/+DdHxXMUp4WDFYQk/xp6Ttc/jz0vU5FWqQClTFy0UMhmz8lY6/XeW1ydP33f5NTnAzmqrWygfuRmH/FYL5FZjCdb4qd7c8/TfbPcNljma8bdxrxL9RU83m1mRsrRdYp0DKVNNeQQPzrknWKV/6QEv/u0xptQHRXkNTYf1jhXwJrgZ7SHeihxMy9bvyPy3d65W9T+PPg30LeN7C/xd1/u2L/WVQfJUhNcyW9Gvwdyv4b/FnYYYVe8Pntermbof/OA/ZfMdS/qM8jDP4IOTW8isdwil1qPXQGzL97uM7ttHQ7l8kVKhuZXLG29uyY+Kc+PDGa/wiRp4b4U8RvHYj42yd9yuv5iyrnfwG+x/e65vOflO/zyh871P9zWMcq9MDLhgPw/7Mk4rB3/i6F/zqq/3U7B9QgFayBl4XS2H+UnL7cK39c4Z/D3itm8Yed/0S8urq//xKy73OsQei9ohP/2fyJJvsXoP9NrEHoIWvbwc0EjbNxNKY18k+3wp+13+/4nxe+Q86f1UuNezcvc5IKv2n8MK0/TTFbLS9Xs9Ut0/m6NXyr9y95nPxNu/bVOX12/GH7X0o5f3fBvwX2EqpwHSQ0+Hua/F/c/+h3cDFY0OUePOpLx3z/c4Hp/yJfReTnIPo/Ea/k9Z8v/d95jKPkPf/1ka/9n5Vh6j9N/vZ/Qxj7GfvvJ//7vykGv/r/Ir/6vxkG/wD52/8NY5xg8A+Sv/3fGMYRBv8Q/d3/6cI0/z+APIR8grzTmB90/gwHzG/af/rQ/7Hj/zD52/9NSp/2yj9C/vZ/0xgvMvjPkb/9n3jfOIN/lILt/0S+dvu/eXALnW9o8sc1+DPk3AEL5Oza04nBRfgBFz0a/GOK/hfts1eG7Uva/R+Xf1zhN0XdMAGY9p8zO6tG84e38xTk+oPGTZy3V9q3D078uUS8+nfiQPxZoZfw/5PrPycV/nmwiwiox653/zTlo/8REZtfxR9QSwECFwsUAAIACAAmYQZVZ7fWyK4DAAAIJgAACQAJAAAAAAAAAAAAAIAAAAAAVUlQYWNrYWdlVVQFAAfvWe5iUEsFBgAAAAABAAEAQAAAAOYDAAAAAA==


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

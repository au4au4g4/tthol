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
MacroID=11f10087-a4ce-43dd-a8ec-6151cbd3a0d7
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
UEsDBBQAAgAIANgDnlPb5pEtrgMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAe0/cxhtP3MYbT9zGHtWUtPU0EUPrct0pZaWoTyFMtLkZXujYklEkkQCa1GTTQpWIXYBylocOfPcMvalQsW/gDX6sqtLv0DrJT6zb1z7RQh3DNz4dqEjxymtJ37zZwz5zWEycGXz+nvOx+GftAB3KQw7TdidE55z5JiI0UUkn/vNxoN9+3GGdoKvyERaUNh6w5IJyQGiULikPOQLkgCknRMT92QtJy3f6bGtsUy1fCzRVm6TVWMdXpDHGRwYtxnWcd8dylSrXz99c0K4/XbmPPefZqna6SPKFmWyx86htcd1c9mqUALBiuIS/6M9B0ufwH63qASLVKRKhi56KGQzZ+WsdfrvIgcXf/93+QMp4O5Wr1i4H4k5l83mG+RGUznm2J359NPk/0zXPZE5uvGnWb8C7XUfF5t5sZKkXWKtEJlzTUkEf+6ZJ3ilT+kxL8HtE6bEN0VpDT2H1b4l8FapGd0D3ooM3OvG/8T8ple+SMKfwH828j7Bva3uPvvUOw/i+qjDKlrrqRXg/+csv8mfw52WKWXfH67Xu5m6L/zgP1XDfUv6vMogz9KTg2v4gmcYo/aD50B8+8drnM7Ld3NZ/PF6mY2X6qvPz8h/qn3T43mP0bkqSP+lPBbByL+9kmf8nr+Ysr5X4Dv8b2u9fyn5PO88scP9f881rEGPfCy4QD8/wKJOOydv0vhv4Xqf8POAXVIFWvgZaEM9h8jpy/3yp9Q+Oew96pZ/GHnPxGvbvzdfxnZ9wXWIPRe1Yn/bP5ki/2L0P8W1iD0kLPt4GaC5tk4GtMa+adb4c/Zz3f8zwvfIefP6qXmvZuXOSmF3zR+mNafppitVVZqudq26XzdGr7d+5cCTv6WXfvqnD47/rD9L62cvyXwb4O9jCpcB0kN/p4W/xf3P/odXBwWdLkHj/jOSd//XGb6v8hXUfk6iP5PxCt5/edL/3cJ4yh5z3995Gv/Z2WZ+s+Qv/3fEMZ+xv77yf/+b4rBr/6/yK/+b4bBP0D+9n/DGCcY/IPkb/83hnGEwT9E//Z/ujDN/w8hjyAfIe805gedP8MB85v2nz70f+z4P0z+9n+T0qe98o+Qv/3fNMYrDP6L5G//J543zuAfpWD7P5Gv3f5vHtxC55ua/AkN/iw5d8ACebv2dGJwCX7ARY8G/5ii/0X77FVg+7J2/8flH1f4TdEwTACm/efM7prR/OGdAgW5/qBxB+fttfbtgxN/rhKv/p04EH9W6RX8//T6z0mFfx7sIgLqsevdP0356H9ExOZX8QdQSwECFwsUAAIACADYA55T2+aRLa4DAAAIJgAACQAJAAAAAAAAAAAAAIAAAAAAVUlQYWNrYWdlVVQFAAe0/cxhUEsFBgAAAAABAAEAQAAAAOYDAAAAAA==


[Script]
Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
hwnd = dm.FindWindow("", "絕代方程式")
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
t.init hwnd

hwnds = t.getPartHwnds()
For Each hwnd In hwnds
	t.init hwnd
	t.skill 550
Next

// 入門派
//While true
//	key = WaitKey()
//	Select case key
//	Case 97
//		t.learnSkills(array(array(1,1),array(4,1),array(4,2),array(4,3),array(5,2),array(5,3),array(9,1),array(7,1)))
//		t.talkNPC &H17B9, &H1, array(1, 2, 2, 2, 4, array(2, 10))
//	Case 98
//		t.talkNPC &H17B9, &H1, array(1, 2, 2, 2)
//		t.talkNPC &H197A, &H1, array(1, 4, 6, 4)
//	Case 99
//		t.talkNPC &H19DF, &H1, array(1, 2, 5, 4, 4, 2, 2)
//	Case 100
//		t.talkNPC &H1789, &H1, array(1, 2, 8, 2, 4, 4, 2, 2)
//	Case 101
//		t.talkNPC &H1DD4, &H1, array(1, array(2, 12))
//		Delay 500
//		t.talkNPC &H1DD5, &H27, array(1, array(2, 6))
//	Case 102
//		t.talkNPC &H1DD4, &H1, array(1, array(2, 5))
//	End Select
//Wend

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

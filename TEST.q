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
UEsDBBQAAgAIAGhNa1iZyCSKrgMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAco0u5lKNLuZSjS7mXtWUtPU0EUPrfl0ZelRShPsbwUWenSxJhYIpEE0dBi1ESTglWIfZAWDO78GS5l7coFC3+Aa1270qV/gJ3Ub+6da6cI4Z6ZC7UJHzlMaTv3mzlnzmsIkoNvX5M/dj8N/6RDuE1BOqiHqUt5z5JiI0EUkH8f1Ot19+36OdoKvyEd0obC1p2QbkgYEoJEIBcgUUgMEndMTz2QpJx3cK7GtsUyVfCzRWm6S2WMVXpLHKRwYtxnWSd8N33zw/ZK7bsVxOt3Yee9FVqg66SPEFmWyx84gdcd1c/mKEeLBiuISP6U9B0ufw763qQCLVGeShi56KWAzZ+UsdfrvA45uv77v8k5zgbzlWrJwP1IzL9hMN8iM5jON8Xe7pdfJvtnuOypzNeNO434F2iq+bzazI2VIuvkaZWKmmuII/5FZZ3ilT+gxL9HtEE1iO4KEhr7Dyr8y2DN0wt6AD0UmbnXjf8x+Uyv/B0Kfw78O8j7Bva3uPvvVOw/h+qjCKlqrqRPg79L2X+DPwM7rNFrPr9dL/cw9N99yP5rhvoX9XmIwR8ip4ZX8QxOsU/th+4W8+8frXM7Ld3PprP5ci2dLVQ3Xp4S//TH50bznyLyVBF/CvitAxF/+6VPeT1/YeX8L8L3+F7XfP4T8nle+SNH+n8W61iHHnjZcBD+f5FEHPbOH1X476D637RzQBVSxhp4WSiF/YfJ6cu98scU/nnsvWwWf9j5T8SrW3/3X0T2fYU1CL2XdeI/mz/eZP889L+FNQg9ZGw7uJmgcTaOx4xG/ulR+DP28x3/88J3xPmz+qhx7+ZlTkLhN40fpvWnKeYqpdVKprJjOl+3hm/3/iWHk79l1746p8+OP2z/Syrn7yH4d8BeRBWug7gGf2+T/4v7H/0OLgILutxDx3zntO9/rjD9X+SrkHzdiv5PxCt5/edL/3cZ4xh5z3/95Gv/Z6WZ+k+Rv/3fMMYBxv4HyP/+b5rBr/6/yK/+b5bBP0j+9n8jGCcZ/EPkb/83jnGUwT9M//Z/ujDN/48hTyCfIe815rc6fwZbzG/af/rQ/7Hj/wj52/9NSZ/2yj9K/vZ/MxivMvgvkb/9n3jeBIN/jFrb/4l87fZ/C+AWOq9p8sc0+NPk3AELZO3a04nBBfgBF70a/OOK/pfss1eC7Yva/R+Xf0LhN0XdMAGY9p+ze+tG80d2c9TK9bca93De3mjfPjjx5xrx6t/JQ/Fnjbbh/2fXf04p/AtgFxFQj13v/mnaR/8jIja/ij9QSwECFwsUAAIACABoTWtYmcgkiq4DAAAIJgAACQAJAAAAAAAAAAAAAIAAAAAAVUlQYWNrYWdlVVQFAAco0u5lUEsFBgAAAAABAAEAQAAAAOYDAAAAAA==


[Script]
//Set dm = createobject("dm.dmsoft")
//Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
//hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
//dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
Set dm = createobject("dm.dmsoft")
hwnd = dm.FindWindow("", "絕代方程式")
Import "QMScript/Util.vbs" : Set u = New Util
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
t.init hwnd

hwnds = t.getAllHwnds()
datas = array()
For Each hwnd In hwnds
	t.init(hwnd)
	Redim Preserve datas(ubound(datas) + 1)
	datas(ubound(datas)) = array(t.id, t.place, t.deposit + t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6)
//		datas(ubound(datas)) = array(t.id, t.level, t.place, t.period, t.monster, t.expp, t.money, t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6, t.getItemCnt("特貢令"))	
Next
u.post "掛機", datas
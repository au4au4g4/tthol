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
UEsDBBQAAgAIAMoZUFiz/E13vAMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAcd085lHdPOZR3TzmXtWUtPU0EUPrfl0ZZaWoTyFMtLkZWu3BgT20gkQTS0GjXRULAKsQ9S0ODOn+DSLWtXLFj4A1zr2pUu/QPspH5z71w7RcB7Zi7UJnzkMKXt3G/mnDmvIUgOvn5JfN/eGfpBB3CTgrRfC1OH8p4lxUacKCD/3q/Vau7btTO0FH5B2qQNha3bIZ2QMCQEiUDOQbogUUjMMT11QxJy3v6ZGlsWi1TBzyal6DaVMVbpLXGQxIlxn2X947vr73euZ5a+WUG8fhd23ntAc3SV9BEiy3L5A0d8Z2nN4XX51c8ylKN5gxVEJH9S+s5x/O6ofpaDvtepQAuUpxJGLnooYPMnZOz1Oq9Njq7//m9yhtPBbKVaMnA/EvOvGcy3yAym802xu/35p8n+GS57IvN14049/gUaaj6vNnNjpcg6eVqmouYaYoh/XbJO8cofUOLfQ1qjDYjuCuIa+w8q/ItgzdNzugc9FJm5143/UflMr/xtCn8O/FvI+wb2t7j7b1fsn0H1UYRUNVfSq8Hfoey/zp+GHVboFZ/frpe7GfrvPGD/FUP9i/o8xOAPkVPDq3gKp9ij1kNnk/n3Dte5nZbuZlPZfHkjlS1U116cEP/Ux2dG858g8lQRfwr4rQMRf/ukT3k9f2Hl/M/D9/he13j+4/J5Xvkjh/p/FutYhR542XAA/n+eRBz2zt+l8N9C9b9u54AqpIw18LJQEvsPk9OXe+WPKvyz2HvZLP6w85+IVzf+7L+I7PsSaxB6L+vEfzZ/rMH+eeh/E2sQekjbdnAzQf1sHI1pjfzTrfCn7ec7/ueF75DzZ/VS/d7Ny5y4wm8aP0zrT1NkKqXlSrqyZTpft4Zv9f4lh5O/ade+OqfPjj9s/0so5+8++LfAXkQVroOYBn9Pg/+L+x/9Di4CC7rcg0d857j7Jz/ufy4x/V/kq5B83Yz+T8Qref3nS/93EeMoec9/feRr/2elmPpPkr/93xDGfsb++8n//m+Kwa/+v8iv/m+GwT9A/vZ/wxgnGPyD5G//N4ZxhME/RH/3f7owzf+PII8hnyAfNOY3O38Gm8xv2n/60P+x4/8w+dv/TUqf9so/Qv72f9MYLzP4L5C//Z943jiDf5Sa2/+JfO32f3PgFjrf0OSPavCnyLkDFsjatacTgwvwAy56NPjHFP0v2GevBNsXtfs/Lv+4wm+KmmECMO0/Z3ZXjeYPb+eometvNu7gvL3Rvn1w4s8V4tW/Ewfizwq9hv+fXv85qfDPgV1EQD12vfunKR/9j4jY/Cp+A1BLAQIXCxQAAgAIAMoZUFiz/E13vAMAAAgmAAAJAAkAAAAAAAAAAAAAgAAAAABVSVBhY2thZ2VVVAUABx3TzmVQSwUGAAAAAAEAAQBAAAAA9AMAAAAA


[Script]
Set dm = createobject("dm.dmsoft")
Import "QMScript/Tthol.vbs" : Set t = New Tthol : t.init
Import "QMScript/Util.vbs" : Set u = New Util
hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
hwnds = t.getAllHwnds()
dm_ret = dm.BindWindow(hwnd, "normal", "windows3", "windows", 0)
//
//Set dm = createobject("dm.dmsoft")
//hwnd = dm.FindWindow("", "絕代方程式")
//Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
//t.init hwnd

datas = array()
For Each hwnd In hwnds
	t.init(hwnd)
	Redim Preserve datas(ubound(datas) + 1)
	datas(ubound(datas)) = array(t.id, t.place, t.deposit + t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6)
//		datas(ubound(datas)) = array(t.id, t.level, t.place, t.period, t.monster, t.expp, t.money, t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6, t.getItemCnt("特貢令"))	Next
u.post "掛機", datas
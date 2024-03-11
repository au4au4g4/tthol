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
UEsDBBQAAgAIABRMa1j1DNlDtQMAAAgmAAAJABEAVUlQYWNrYWdlVVQNAAfKz+5lys/uZcrP7mXtWUtPU0EUPrcFaUstLUJ5iuWlyEr3xsQSiSSIhlajJpoUrEDsg7RocOfPcMvalQsW/gDXujNxpUv/ADtbv7l3rp0ixHtmLlyb8JHDlLZzv5lz5ryGMDn48jn1fe/D6A86hJsUpkYzSueU9ywpNpJEIfl3o9lsum83z9BR+AXpkjYUtu6G9ECikAgkBjkP6YXEIQnH9NQHScl5jTM1dixWqYqfHcrQbapgrNEb4iCNE+M+y/rHd982NuL5+jcrLF5Hnfce0BJdI31EyLJc/tAx39n+6vC6/OpnC5SnZYMVxCR/WvoOlz8PfW9TkVaoQGWMXPRTyOZPydjrdV6XHF3//d/kDKeDxWqtbOB+JOZfN5hvkRlM55tif+/TT5P9M1z2RObrxp1W/Au11XxebebGSpF1CrRGJc01JBD/emWd4pU/pMS/h7RFdYjuCpIa+w8r/KtgLdBzugc9lJi5143/cflMr/xdCn8e/LvI+wb2t7j771bsv4DqowSpaa5kQIP/nLL/Fn8Wdlinl3x+u17uY+i/55D91w31L+rzCIM/Qk4Nr+IpnOKAOg89AfMfHK1zOy3dzWVyhUo9kyvWtl6cEP/s+2dG858g8tQQf4r4rQMRfwelT3k9f1Hl/C/D9/he137+k/J5XvljR/p/DuvYhB542XAY/n+BRBz2zt+r8N9C9b9t54AapII18LJQGvuPktOXe+WPK/yL2HvFLP6w85+IVzf+7L+E7LuBNQi9V3TiP5s/0Wb/AvS/gzUIPWRtO7iZoHU2jsecRv7pU/iz9vMd//PCd8T5swaode/mZU5S4TeNH6b1pykWquW1ara6azpft4bv9P4lj5O/Y9e+OqfPjj9s/0sp5+8++HfBXkIVroOEBn9/m/+L+x/9Di4GC7rcI8d8x733Oan7n8tM/xf5KiJfB9H/iXglr/986f8uYZwg7/lvkHzt/6wMU/9p8rf/G8U4xNj/EPnf/80y+NX/F/nV/80z+IfJ3/5vDOM0g3+E/O3/JjGOM/hH6e/+Txem+f8R5DHkI+Sdxvyg82c4YH7T/tOH/o8d/8fI3/5vRvq0V/5x8rf/m8N4hcF/kfzt/8Tzphj8ExRs/yfytdv/LYFb6LyuyR/X4M+QcwcskLNrTycGF+EHXPRr8E8q+l+xz14Zti9p939c/imF3xRNwwRg2n/O728azR/by1OQ6w8ad3DeXmvfPjjx5yrx6t/pQ/FnnV7B/0+v/5xR+JfALiKgHrve/dOsj/5HRGx+Fb8BUEsBAhcLFAACAAgAFExrWPUM2UO1AwAACCYAAAkACQAAAAAAAAAAAACAAAAAAFVJUGFja2FnZVVUBQAHys/uZVBLBQYAAAAAAQABAEAAAADtAwAAAAA=


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
//		datas(ubound(datas)) = array(t.id, t.level, t.place, t.period, t.monster, t.expp, t.money, t.cash + t.getItemCnt("百萬官幣") * 10 ^ 6, t.getItemCnt("特貢令"))	
Next
u.post "掛機", datas
[General]
SyntaxVersion=2
BeginHotkey=121
BeginHotkeyMod=0
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=105
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=39c5f415-fecf-4813-b6c3-5b92fb1e7c66
Description=小方排序
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
Import "Util.vbs" : Set u = New Util
Import "Tthbn.vbs" : Set t = New Tthbn

VBSBegin

Set dm = createobject("dm.dmsoft")
hwnds = t.getAllHwnds()
For Each hwnd In hwnds
	tthbns = dm.GetModuleBaseAddr(hwnd, "tthbn.bin")
	' 排序錢莊
	Call order(&HAFCC0, &HAFCBC,GetRef("compare"))
	' 排序背包
	Call order(&H101E98, &H111604,GetRef("compare"))
Next

Function order(addr1, addr2,compare)
	count = dm.ReadInt(hwnd, HEX(tthbns + addr2), 1)
	redim bag(count)
	For i = 0 To count - 1
		bag(i) = dm.ReadData(hwnd, HEX(tthbns + addr1 + i * 280), 280)
	Next
	u.sort bag, GetRef("compare")
	For i = 0 To count - 1
		Call dm.WriteData(hwnd, HEX(tthbns + addr1 + i * 280), bag(i))
	Next
End Function

Function compare(a, b)
	compare = mid(a, 16, 2) & mid(a, 13, 2) > mid(b, 16, 2) & mid(b, 13, 2)
End Function

VBSEnd
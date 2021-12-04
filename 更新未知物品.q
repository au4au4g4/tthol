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
MacroID=3726b690-a17e-4af8-a6f3-94b2c1787fb8
Description=更新未知物品
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
id = array("BD81","BE81","BF81","C081")
name = array("奶香冰棒","清涼冰塊","雪大福","薄荷草")

Set dm = createobject("dm.dmsoft")
hwnds = split(dm.EnumWindow(0, "絕代方程式", "", 1 + 4 + 8 + 16), ",")
For Each hwnd In hwnds
	base = dm.GetModuleBaseAddr(hwnd, "tthsj.bin")
	lower = HEX(base + &H1617000)
	upper = HEX(base + &H18F0000)
	For i = 0 To UBound(id)
		result = dm.FindData(hwnd, lower & "-" & upper, id(i) & "00001D")
		If (len(result) = 0) + (InStr(result, "|") <> 0) Then 
			TracePrint "錯誤|" & name(i) & "|" & result & "|" & hwnd
		Else 
			dm_ret = dm.WriteString(hwnd, result & "+C", 0, name(i))
		End If
	Next
Next

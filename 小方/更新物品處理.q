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
MacroID=05978246-8fe2-4176-8df6-3f04489367d6
Description=��s���~�B�z
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
Set dm = createobject("dm.dmsoft")

itemActions = split(dm.GetClipboard(), chr(10))
hwnds = split(dm.EnumWindow(0, "���N��{��", "", 1 + 4 + 8 + 16), ",")
For Each hwnd In hwnds
	base = dm.GetModuleBaseAddr(hwnd, "tthbn.bin")
	range = HEX(base + &HC1010) & "-" & HEX(base + &HE4120)
	For Each itemAction In itemActions
		If itemAction <> "" Then 
    		itemAction = split(itemAction, "=")
    		id = "&H" & itemAction(0)
    		result = dm.FindInt(hwnd, range, id, id, 0)
    		If (len(result) = 0) + (InStr(result, "|") <> 0) Then 
    			TracePrint len(result)
				TracePrint "���~|" & itemAction(0)
			Else 
				dm_ret = dm.WriteInt(hwnd, HEX(clng("&H"&result) + 4), 0, itemAction(1))
				dm_ret = dm.WriteInt(hwnd, HEX(clng("&H"&result) + 8), 0, itemAction(2))
			End If	
		End If
	Next
Next
Beep

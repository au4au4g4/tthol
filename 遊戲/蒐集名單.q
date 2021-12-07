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
MacroID=87a78e58-6ece-482b-944b-fa44acc8d2b6
Description=»`¶°¦W³æ
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
Import "Tthol.vbs"
Set t = New Tthol

hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
t.init hwnd
While true
	addr = t.getMainAddr()
	While addr <> 0
		level = dm.ReadInt(hwnd, "[" & HEX(addr + 924) & "]+EA", 0)
		If (level > 140) * (level < 201) Then 
			name = dm.ReadString(hwnd, HEX(addr + 484), 0, 16)
			If dm.ReadIni("list", name, ".\list.ini") = "" Then 
				family = dm.ReadString(hwnd, HEX(addr + 545), 0, 16)
				dm.WriteIni "list", name, family, ".\list.ini"
			End If
		End If
		addr = dm.ReadInt(hwnd, HEX(addr + 292), 0)
		Delay 200
	Wend
Wend

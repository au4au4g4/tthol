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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set dm = createobject("dm.dmsoft")

itemActions = split(dm.GetClipboard(), chr(13) & chr(10))
hwnds = split(dm.EnumWindow(0, "���N��{��", "", 1 + 4 + 8 + 16), ",")
For Each hwnd In hwnds
	t.init hwnd
	For Each itemAction In itemActions
		If itemAction <> "" Then 		
    		itemAction = split(itemAction, chr(9))
    		t.updateItem itemAction(0), itemAction(1)
		End If
	Next
Next

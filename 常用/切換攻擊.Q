[General]
SyntaxVersion=2
BeginHotkey=98
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=98
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=369742b4-a928-4ace-bf82-b2b70f42bdf6
Description=¤Á´«§ðÀ»
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
Import "Tthbn.vbs" : Set t = New Tthbn

hwnds = t.getPartHwnds()
While true
	flag = (Hour(now) < 18) * - 1 
	For Each hwnd In hwnds
		t.groupAtk flag
		t.init hwnd
		range = iif(flag,t.gRange,t.nRange)
		t.setRange(range)
	Next
	Delay 60 * 60 * 1000
Wend

Function iif(flag, tResult, fResult)
	If flag Then 
		iif = tResult
	Else 
		iif = fResult
	End If
End Function
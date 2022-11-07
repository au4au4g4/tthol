[General]
SyntaxVersion=2
BeginHotkey=56
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=21bf90ac-ccf2-4fb6-90c9-23cb0b5d4566
Description=¯}¸Ñ
Enable=1
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

t.init t.getHwnd()
memberST = t.memberST(t.account)
If UBound(memberST) > - 1  Then 
	teamID = t.memberST(t.account)(4)
	teamSt = t.teamSt(teamID)
	t.crackAll teamSt(0), teamSt(1), teamSt(2)
Else 
	t.crack
End If


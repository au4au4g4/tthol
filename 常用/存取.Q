[General]
SyntaxVersion=2
BeginHotkey=104
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=dc7ce714-1fb1-43b7-9efc-a6dd3a7b1ac7
Description=存取
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
Set dm = createobject("dm.dmsoft")
Import "Tthbn.vbs" : Set t = New Tthbn

UserVar isPop=DropList{"取":1|"存":0}=0 "動作"
UserVar keywords "關鍵字" : keywords = split(keywords," ")
hwnd = t.getHwnd()
t.init hwnd
If isPop Then 
    t.pop keywords
Else
    t.push keywords
End If


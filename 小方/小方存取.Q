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
MacroID=dc7ce714-1fb1-43b7-9efc-a6dd3a7b1ac7
Description=小方存取
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
hwnd = dm.GetForegroundWindow()
t.init hwnd
UserVar isPop=DropList{"取":1|"存":0}=0 "動作"
UserVar keywords "關鍵字"

t.openbank
keywords = split(keywords," ")
For Each keyword In keywords
    If isPop Then 
        t.popItem keyword
    Else
        t.pushItem keyword
    End If
    Delay 50
Next


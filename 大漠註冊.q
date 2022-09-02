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
MacroID=aca3796a-df62-482c-8c49-1f527f7089fb
Description=¤jºzµù¥U
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
set ws=createobject("Wscript.Shell")
ws.run "regsvr32 /s .\QMScript\dm.dll "
ws.run "regsvr32 /s .\QMScript\dynwrap.dll "
set ws=nothing

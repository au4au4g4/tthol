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
MacroID=4833c75b-8416-457a-a496-00cbe0a18a95
Description=複製設定
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
src = "C:\Users\god\Desktop\tool\JD_FCS21.60\config\"
source = "雷山小過"
accounts = split(dm.ReadIni("addr", "m", ".\QMScript\tthbn.ini"),",")
For Each account In accounts
	id = split(dm.ReadIni("addr", account, ".\QMScript\tthbn.ini"), ",")(4)
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).dat", src & id & "_飛雁山莊(花).dat"	
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).dbt", src & id & "_飛雁山莊(花).dbt"
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).ddt", src & id & "_飛雁山莊(花).ddt"
Next

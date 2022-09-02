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
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set re = New RegExp
Set dm = createobject("dm.dmsoft")

hwnds = array(t.getHwnd())
//hwnds = t.getAllHwnds()
For each hwnd in hwnds
	t.init hwnd
//	t.crack
//	t.moveSpeed 100,230
//	t.atkrange(2)
//	t.fixbank 
//	t.fullsupport
	t.crackall 100, 230, 2
Next

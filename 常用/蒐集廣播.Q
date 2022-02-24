[General]
SyntaxVersion=2
BeginHotkey=103
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=103
StopHotkeyMod=4
RunOnce=1
EnableWindow=
MacroID=a8017fbe-f143-41fe-a597-74fd52b52602
Description=蒐集廣播
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
hwnd = dm.FindWindow("", "絕代方程式")
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
t.init hwnd

Dim fileStr, count, str, typ, last, temp
While true
	fileStr = dm.ReadFile("C:\Users\god\Desktop\系統對話.txt")
	last = mid(fileStr, InstrRev(fileStr, " ["))
	count = dm.ReadInt(hwnd, "[[<ttha.bin>+D80DC]+580]+6C", 0)
	TracePrint count
	For i = 0 To count - 1
		str = dm.ReadString(hwnd, "[[[[<ttha.bin>+D80DC]+580]+68]+" & HEX((count - 1 - i) * 4) & "]", 0, 100) & chr(10)
		typ = dm.ReadInt(hwnd, "[[[<ttha.bin>+D80DC]+580]+90]+" & HEX((count - 1 - i) * 4), 0)
		If str = last Then 
			Exit For
		ElseIf typ = 200 Then
			temp = str  & temp
		End If
	Next
	dm.WriteFile "C:\Users\god\Desktop\系統對話.txt", temp
	TracePrint 999
	Delay 12 * 60 * 60 * 1000
Wend


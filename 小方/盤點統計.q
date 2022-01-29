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
MacroID=bc4a6b0b-1aa2-4b01-9a01-fd8fe69e615a
Description=盤點統計
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
Import "Util.vbs" : Set u = New Util
Import "Tthbn.vbs" : Set t = New Tthbn

VBSBegin
itemTable = array()
statisticsTable = array()
hwnds = t.getAllHwnds()
For Each hwnd In hwnds
	t.init (hwnd)
	bag = t.getBag(array(""))
	bank = t.getBank(array(""))
	monster = t.getMonster()
	reward = t.getReward()
	arrayToTable bag, itemTable, array("id", "name", "cnt"), array(t.id, "背包")
	arrayToTable bank, itemTable, array("id", "name", "cnt"), array(t.id, "錢莊")
	arrayToTable monster, statisticsTable, array("id", "name", "cnt"), array(t.id, "怪物")
	arrayToTable reward, statisticsTable, array("id","name", "cnt"), array(t.id, "物品")
Next

u.post "盤點", itemTable
'u.post "統計", statisticsTable

Function arrayToTable(arr, table, keys, tails)
	Dim row
	For i = 0 To UBound(arr)
		row = array()
		Set item = arr(i)
		For Each key In keys
			Redim Preserve row(UBound(row) + 1)
			row(UBound(row)) = item.item(key)
		Next
		For Each tail In tails
			Redim Preserve row(UBound(row) + 1) : row(UBound(row)) = tail
		Next
		Redim Preserve table(UBound(table) + 1) : table(UBound(table)) = row
	Next
End Function
VBSEnd

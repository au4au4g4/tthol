[General]
SyntaxVersion=2
BeginHotkey=101
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=bc4a6b0b-1aa2-4b01-9a01-fd8fe69e615a
Description=盤點統計
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
Import "Util.vbs" : Set u = New Util
Import "Tthbn.vbs" : Set t = New Tthbn

k = 0 : l = 0
hwnds = t.getAllHwnds()
For Each hwnd In hwnds
	t.init (hwnd)

	bag = t.getBag("")
	bank = t.getBank("")
	monster = t.getMonster()
	reward = t.getReward()
	
	For i = 0 To UBound(bag)
		Set item = bag(i)
		Redim Preserve arr1(k) : arr1(k) = array(item.item("id"),item.item("name"), item.item("cnt"), t.id(), "背包")
		k = k + 1
	Next
	For i = 0 To UBound(bank)
		set item = bank(i)
		Redim Preserve arr1(k) : arr1(k) = array(item.item("id"),item.item("name"), item.item("cnt"), t.id(), "錢莊")
		k = k + 1
	Next
	For i = 0 To UBound(monster)
		set item = monster(i)
		Redim Preserve arr2(l) : arr2(l) = array(item.item("name"), item.item("cnt"), t.id(), "怪物")
		l = l + 1
	Next
	For i = 0 To UBound(reward)
		set item = reward(i)
		Redim Preserve arr2(l) : arr2(l) = array(item.item("name"), item.item("cnt"), t.id(), "物品")
		l = l + 1
	Next
Next

task1 = array("盤點", 1, arr1)
task2 = array("統計", 1, arr2)
u.postTasks(array(task1,task2))

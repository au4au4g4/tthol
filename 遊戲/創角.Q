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
MacroID=2205ec83-8fb2-4ea5-8854-6811db7d0838
Description=創角
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
Import "Tthol.vbs" : Set t = New Tthol
hwnd = dm.FindWindow("_UJONLINE_", "Tthol")
t.init hwnd

t.trans &H147'進入傳點
t.talkNPC &HEA64, t.mySN, array(2, 2, 2, 4, 2)'進入對話
t.go (array(26, 9))
Delay 1000
t.talkNPC &H1901, &H1, array(1, array(2,5))'明若蘭1
t.go (array(44, 7))
Delay 3000
t.trans &H156'*入弓手傳點

t.go (array(6, 40))
Delay 3000
t.trans &H15A	'*出弓手傳點
t.talkNPC &HEA64, t.mySN, array(array(2,20))'出弓手對話
t.talkNPC &H1901, &H3, array(1,array(2,40))'明若蘭3	
sn = t.getBag(array("陸天灝衣")).GetByIndex(0).item("sn")
t.wear &H5E9A,sn
t.talkNPC &H1901, &H3, array(1,array(2,30))'明若蘭3
t.learnSkills(array(array(8,1),array(4,1),array(5,1),array(2,1),array(3,1),array(6,1),array(522,1),array(523,1),array(524,1),array(525,1),array(550,1),array(586,1),array(588,1),array(587,1)))
t.trans &H151'*惡乞丐傳點

t.go (array(20, 69))
Delay 2000
t.talkNPC &H1902, &H1, array(1,array(2,30))'惡乞丐	 
t.trans &H158	'*入山洞傳點
Delay 1000
t.go (array(17, 12))
Delay 3000

t.trans &H15D	'*泉水
t.talkNPC &HEA64, t.mySN, array(array(2,10))
sn = t.getBag(array("泉水")).GetByIndex(0).item("sn")
t.compound0 &H4E5F, &H5E9B, sn
t.go (array(8, 10))
Delay 2000

t.trans &H166'*出山洞傳點
t.talkNPC &HEA64, t.mySN, array(array(2,10))	'出山洞對話
t.go (array(61, 68))
Delay 2000

t.talkNPC &H1901, &H2, array(1,array(2,90))	'明若蘭2
t.go (array(109, 64))
Delay 1000
t.trans &H14B	'*結束傳點
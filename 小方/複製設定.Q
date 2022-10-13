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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn
Set dm = createobject("dm.dmsoft")
root = "C:\Users\god\Desktop\tool\JD_FCS21.60\config\"

teamIDs = t.teamIDs
For each teamID in teamIDs
	IDs = t.IDs(teamID)
	teamST = t.teamST(teamID)
	For Each ID In IDs
		memberST = t.memberST(ID)
		
		// dat
		dat = root & fileName(memberST(1)) & ".dat"
		Plugin.File.CopyFile root & "template.dat", dat
		dm.WriteIni "BASE", "gocity", teamST(3), dat
		dm.WriteIni "SKILL","skillgslq",teamST(4),dat
		dm.WriteIni "SKILL", "skillswbgswz", teamST(5), dat
		dm.WriteIni "HIT", "loopskill1", teamST(6), dat
		dm.WriteIni "MAP","gjmapname",memberST(2),dat
		dm.WriteIni "MAP", "arrrect", replace(memberST(3), "|", ","), dat
		
		// ddt
		ddt = root & fileName(memberST(1)) & ".ddt"
		dm.DeleteFile ddt
		bagMeds = split(teamST(7),"|")
		For Each bagMed In bagMeds
			dm.WriteFile ddt, Replace(bagMed,"x",",") & ",從倉庫取,賣掉,1" + chr(10)
		Next
		petMeds = split(teamST(8),"|")
		For Each petMed In petMeds
			dm.WriteFile ddt, Replace(petMed,"x",",") & ",購買,不處理,2" + chr(10)
		Next
	Next
Next

Function fileName(name)
	fileName = name & "_飛雁山莊(花)"
	If instr("四象,六十四卦,枯里珍,夏枯草", name) Then 
		fileName = big5(fileName)
	End If
End Function

Function big5(str)
	big5 = ""
	For i = 1 To Len(str)
    	big5 = big5 + lcase(hex(asc(Mid(str, i, 1))))
	Next
End Function
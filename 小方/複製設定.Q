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
src = "C:\Users\GOD\Desktop\JD_FCS21.60\config\"
source = ""
ids = array("乾為天","天地否","天澤履","天山遯","天火同人","天水訟","天雷無妄","天風姤","地天泰","坤為地","地澤臨","地山謙","地火明夷","地水師","地雷復","澤天夬","澤地萃","兌為澤","澤山咸","澤火革","澤水困","澤雷隨","澤風大過","山天大畜","山地剝","山澤損","艮為山","山火賁","山水蒙","山雷頤","山風蠱","火天大有","火地晉","火澤睽","火山旅","離為火","火水未濟","火雷噬嗑","火風鼎","水天需","水地比","水澤節","水山蹇","水火既濟","坎為水","水雷屯","水風井","雷天大壯","雷地豫","雷澤歸妹","雷山小過","雷火豐","雷水解","震為雷","雷風恒","風天小畜","風地觀","風澤中孚","風山漸","風火家人","風水渙")
For Each id In ids
//	Call Plugin.File.DeleteFile(src & id & "_飛雁山莊(花).dat")
//	Call Plugin.File.DeleteFile(src & id & "_飛雁山莊(花).dbt")
//	Call Plugin.File.DeleteFile(src & id & "_飛雁山莊(花).ddt")
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).dat", src & id & "_飛雁山莊(花).dat"	
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).dbt", src & id & "_飛雁山莊(花).dbt"
	Plugin.File.CopyFile src & source & "_飛雁山莊(花).ddt", src & id & "_飛雁山莊(花).ddt"
Next

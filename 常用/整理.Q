[General]
SyntaxVersion=2
BeginHotkey=54
BeginHotkeyMod=2
PauseHotkey=0
PauseHotkeyMod=0
StopHotkey=123
StopHotkeyMod=0
RunOnce=1
EnableWindow=
MacroID=0e8977bb-a552-405d-8d8d-608e4cf1529b
Description=整理
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
Import "QMScript/Tthbn.vbs" : Set t = New Tthbn

hwnd = t.getHwnd()
tasks = split(dm.getclipboard(), chr(13) & chr(10))
For Each task In tasks
	order = split(task, "-")
	params = split(order(1), " ")
	
	If instr("",order(0)) > 0 Then 
		sellers = array(t.getHwnd)
	Elseif instr("itgou",order(0)) > 0 then
		sellers = t.getAllHwnds()
	Else 
		sellers = t.getPartHwnds()
	End If

	For Each seller In sellers
		t.init seller
	 	Select Case order(0)
    		Case "o"
        		t.pop params
    		Case "u"
        		t.push params
    		Case "t"
				t.trade hwnd, params, 100000
			Case "g"
				t.init hwnd
				t.trade seller, array(params(0)), params(1)
			Case "a"
        		t.apply params(0)
			Case "l"
        		t.learn params
			Case "i"
    			t.updateItem params
    		Case "m"
    			t.map params(0)
    		Case "f"
    			t.frame params
    		Case "b"
    			If params(1) = "百萬官幣" Then 
    				params(2) = int(t.cash / 10 ^ 6)
    			End If
				t.buy params(0), params(1), params(2)
				t.reset
    		Case "s"
				t.sell params(0), params(1)
		End Select
	Next  
Next

For 10
	
Next
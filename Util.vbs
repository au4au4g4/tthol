Class Util
	
	Private dm,hwnd,http,json
	
	Public Sub Class_Initialize
		Set dm = createobject("dm.dmsoft")
		Set http = createObject("winhttp.winhttprequest.5.1")
    End Sub
	
    Public default function Init()
		
    End function

	Public Function post(sheet,clear,data)
		postTasks array(array(sheet, clear,data))
	End Function
	
	Public Function postTasks(data)
		include "VbsJson.vbs"
		Set json = New VbsJson
		http.open "POST", "https://script.google.com/macros/s/AKfycbwXYvPlGW7wt5GIfkD-a7xPWDjpzdOhNHNEDFqTznO4aQpr4oRUz_bmaBKxe3D0VuOcPQ/exec", False
		http.setrequestheader "Content-Type","application/json"
		http.send json.encode(data)
	End Function
	
	' line notify
	Function pushMsg(str)
		http.open "POST", "https://notify-api.line.me/api/notify", False
		http.setrequestheader "Content-Type", "application/x-www-form-urlencoded"
		http.setrequestheader "Authorization","Bearer MSakOPLvpgbeYtuiU8K2cRUjgiHruOQiQ37Jd7CNoKz"
		http.send "message="&str
	End Function
	
	Private function include(className)
		Dim classFile : classFile = className
		Dim fsObj : Set fsObj = CreateObject("Scripting.FileSystemObject")
		Dim vbsFile : Set vbsFile = fsObj.OpenTextFile(classFile, 1, False)
		Dim myFunctionsStr : myFunctionsStr = vbsFile.ReadAll
		vbsFile.Close
		Set vbsFile = Nothing
		Set fsObj = Nothing
		ExecuteGlobal myFunctionsStr	
	end function
	
	Sub sort(ByRef arr, ByRef compare)
		QuickSortRecursive arr, LBound(arr), UBound(arr), GetRef("compare")
	End Sub
	Sub QuickSortRecursive(ByRef arr, ByVal l, ByVal r, ByRef compare)
		Dim p, ls, rs, temp
		
		'Zero or 1 item to sort
		If r - l < 1 Then Exit Sub
		
		'Only 2 items to sort
		If r - l = 1 Then
			If compare(arr(l), arr(r)) Then
				temp = arr(l)
				arr(l) = arr(r)
				arr(r) = temp
			End If
			Exit Sub
		End If
		
		'3 or more items to sort
		p = arr(Int((l + r) / 2))
		arr(Int((l + r) / 2)) = arr(l)
		
		ls = l + 1
		rs = r
		
		Do
			'Find the right ls
			While ls < rs And not compare(arr(ls), p)
				ls = ls + 1
			Wend
			'Find the right rs
			While l < rs And compare(arr(rs), p)
				rs = rs - 1
			Wend
			'Swap values if ls is less than rs
			If ls < rs then
				temp = arr(ls)
				arr(ls) = arr(rs)
				arr(rs) = temp
			End If
		Loop While ls < rs
		
		arr(l) = arr(rs)
		arr(rs) = p
		
		'Recursively call function
		'2 or more items in first section
		If l < (rs - 1) Then QuickSortRecursive arr, l, rs - 1, GetRef("compare")
		'2 or more items in second section
		If rs + 1 < r Then QuickSortRecursive arr, rs + 1, r, GetRef("compare")	
	End Sub
	
End Class
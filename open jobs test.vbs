If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
'19032-256-01-RBLD
'Material Description
'Earliest Requirements Date
'Number of Latest Receipt Element
'Latest Receipt Date

Dim objShell, ex, wb, ws
Dim x, i, n, row
Dim tierProds, t
Dim failCase, initWidth
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105

Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory
failCase = "Unmodified"


Call Main

Sub Main
	Set ex = CreateObject("Excel.Application")
	ex.Visible= False
	Set wb = ex.Workbooks.Add
	wb.SaveAs(strpath & "\Adam's Book.xlsx")
	
	Set ws = wb.Sheets(1)
	
	ws.Cells.Select
	ex.Selection.NumberFormat = "@"
	ws.Cells(1,1).select
	
	ex.ScreenUpdating = False
	ex.Calculation = xlCalculationManual
	ex.EnableEvents = False
	ex.DisplayAlerts = False
	
	Call MakeDictionary
	Call GetHierarchy
	
	ex.Visible= True
	ex.ScreenUpdating = True
	ex.Calculation = xlCalculationAutomatic
	ex.EnableEvents = True
	ex.DisplayAlerts = True
	
	ex.Columns.AutoFit
	wb.Save
	
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End Sub



Sub GetHierarchy
	t = 1
	session.findById("wnd[0]/usr").VerticalScrollBar.Position = 0
	initWidth = session.findById("wnd[0]/usr/lbl[8,3]").CharWidth
	ws.Cells(2,1).Value = session.findById("wnd[0]/usr/lbl[8,3]").Text
	ws.Cells(3,2).Value = session.findById("wnd[0]/usr/lbl[12,5]").Text
	ws.Cells(3,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & ",5]").Text
	ws.Cells(3,5).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 50 & ",5]").Text
	ws.Cells(3,6).Value = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + 68 & ",5]").Text)
	tierProds.Item(CStr(t)) = ws.Cells(3,6).Value
	On Error Resume Next
	ws.Cells(3,7).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 88 & ",5]").Text
	If Err.Number <> 0 Then
		Err.Clear
	End If
	On Error Goto 0
	x = 16
	n = 0
	i = 7
	row = 4
	
	Do While failCase <> "Done"
	
		On Error Resume Next
		ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
		
		
		If Err.Number <> 0 Then
			Err.Clear
			Call FailCheck
			Select Case failCase
				Case "Shift Left"
					i = i + 1
					n = n - 4
					row = row + 1
					t = CInt(t) - 1
					'MsgBox("Shifting Left.")
					ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
					ws.Cells(row,2).Value = tierProds.Item(CStr(t))
				Case "Shift Right"
					i = i + 1
					n = n + 4
					row = row + 1
					t = CInt(t) + 1
					tierProds.Item(CStr(t)) = ws.Cells(row - 2,6).Value
					'MsgBox("Shifting Right.")
					ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
					ws.Cells(row,2).Value = tierProds.Item(CStr(t))
				Case "Done"
					MsgBox("Exiting via end of hierarchy")
					Exit Do
			End Select			
		Else
			ws.Cells(row,2).Value = tierProds.Item(CStr(t))
		End If
		On Error Goto 0
		
		
		ws.Cells(row,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & "," & i & "]").Text
		
		ws.Cells(row,5).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 50 & "," & i & "]").Text
		
		ws.Cells(row,6).Value = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + 68 & "," & i & "]").Text)
		
		On Error Resume Next
		ws.Cells(row,7).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 88 & "," & i & "]").Text
		
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		
		If i = 21 Then
			session.findById("wnd[0]/usr").VerticalScrollBar.Position = session.findById("wnd[0]/usr").VerticalScrollBar.Position + 15
			i = 7
		Else
			i = i + 1
		End If
		row = row + 1
	Loop
	
	On Error Goto 0
End Sub



Sub MakeDictionary
	Set tierProds = CreateObject("Scripting.Dictionary")
	tierProds.RemoveAll
	tierProds.Add "1",""
	tierProds.Add "2",""
	tierProds.Add "3",""
	tierProds.Add "4",""
	tierProds.Add "5",""
	tierProds.Add "6",""
	tierProds.Add "7",""
	tierProds.Add "8",""
	tierProds.Add "9",""
	tierProds.Add "10",""
	tierProds.Add "11",""
	tierProds.Add "12",""
	tierProds.Add "13",""
	tierProds.Add "14",""
	tierProds.Add "15",""
	tierProds.Add "16",""
	tierProds.Add "17",""
	tierProds.Add "18",""
	tierProds.Add "19",""
	tierProds.Add "20",""
	tierProds.Add "21",""
	tierProds.Add "22",""
	tierProds.Add "23",""
	tierProds.Add "24",""
	tierProds.Add "25",""
	tierProds.Add "26",""
	tierProds.Add "27",""
	tierProds.Add "28",""
	tierProds.Add "29",""
	tierProds.Add "30",""
End Sub



Sub FailCheck
	On Error Resume Next
	session.findById("wnd[0]/usr/lbl[" & (x + n) - 4 & "," & i + 1 & "]").SetFocus
	
	If Err.Number <> 0 Then
		Err.Clear
		session.findById("wnd[0]/usr/lbl[" & (x + n) + 4 & "," & i + 1 & "]").SetFocus
		If Err.Number <> 0 Then
			Err.Clear
			failCase = "Done"
		Else
			failCase = "Shift Right"
		End If	
	Else
		failCase = "Shift Left"
	End If
	On Error Goto 0
End Sub


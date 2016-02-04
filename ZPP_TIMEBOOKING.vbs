'******************************************
'******************************************
'** Author: Adam Damke					 **
'** 									 **
'** Description: This Script runs the 	 **
'** ZPP_TIMEBOOKING transaction for		 **
'** each plant 500C through 500I and	 **
'** exports the results to an excel 	 **
'** file.  Then it formats the file		 **
'** that was exported as per the job aid **
'** 									 **
'** File: ZPP_TIMEBOOKING.vbs			 **
'** Date Last Updated: June 26th, 2013	 **
'******************************************
'******************************************

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

Dim plant, success, status
Dim savePath, saveFile, saveAsFile, filePath, newFilePath, fso
Dim fso, ex, wb, ws
Dim row, maxRow, entries, delStr

Call Main


Sub Main

	MsgBox("Please wait for the completion message." & chr(13) & "This will take several minutes." & chr(13) & chr(13) & "Click OK to begin.")
	
	plant = "500C"
	Call ExportOpenJobsReport
	plant = "500D"
	Call ExportOpenJobsReport
	plant = "500E"
	Call ExportOpenJobsReport
	plant = "500F"
	Call ExportOpenJobsReport
	plant = "500G"
	Call ExportOpenJobsReport
	plant = "500H"
	Call ExportOpenJobsReport
	plant = "500I"
	Call ExportOpenJobsReport
	
	MsgBox("The requested reports have been created." & chr(13) & chr(13) & "Thank you.")					
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit	

End Sub


Sub ExportOpenJobsReport
	success = False
	status = ""
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_TIMEBOOKING"
	session.findById("wnd[0]").sendVKey 0

	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_ENTRY:1200/chkP_KZ_E1").selected = true
	session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:ZPPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "PP_500X_000"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = plant
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_ENTRY:1200/ctxtP_SYST1").text = "TECO"

	session.findById("wnd[0]").sendVKey 8
	While success = False
	WScript.Sleep(10000)
	If session.findById("wnd[0]/titl").text = "Time Booking Report - Individual Object List" Then
		success = True
	ElseIf session.findById("wnd[0]/titl").text = "SAP" Then
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[0]/tbar[0]/btn[15]").press
		Exit Sub
	End If
	Wend
	
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
	WScript.Sleep(100)
	
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(100)
	
	savePath = "O:\CustSvc\Parts\SRVCTRS\Scripting\ZPP_TIMEBOOKING\Reports\" & plant
	saveFile = plant & " Open Jobs " & Month(Date) & "_" & Day(Date) & "_" & Year(Date) & ".xls"
	saveAsFile = plant & " Open Jobs " & Month(Date) & "_" & Day(Date) & "_" & Year(Date) & ".xlsx"
	filePath = savePath & "\" & saveFile
	newFilePath = savePath & "\" & saveAsFile
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = savePath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = saveFile
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").Text
	If status = "File O:\CustSvc\Parts\SRVCTRS\Scripting\ZPP_TIMEBOOKING already exists" Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press

	'Call ExcelCleanUp

End Sub


Sub ExcelCleanUp

	'******************
	'Steps 1-3*********
	'******************
	Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
	Set ex = WScript.CreateObject("Excel.application")
	ex.Visible = True
	Set wb = ex.Workbooks.Open(filePath)
	Set ws = wb.Sheets(wb.ActiveSheet.Name)
	
	
	'******************
	'Step 4************
	'******************
	ws.Columns("A:A").Select
	ex.Selection.Delete -4131
	ws.Rows("7:7").Select
    ex.Selection.Delete -4162
    ws.Rows("1:5").Select
    ex.Selection.Delete -4162
	
	
	'******************
	'Step 5-6**********
	'******************
	Call updateEntries
	delStr = ""
	For i = 1 To entries
		If ws.Cells(row,5).Value = "" Then
			delStr = delStr & row & ":" & row & ","
		End If
		row = row + 1
	Next
	
	delStr = Left(delStr,Len(delStr)-1)
	ws.Range(delStr).Select
	ex.Selection.Delete -4162
	
	
	'******************
	'Step 7-9**********
	'******************
	Call updateEntries
	delStr = ""
	For i = 1 To entries
		If InStr(CStr(ws.Cells(row,19).Value),"CNF") <> 0 And InStr(CStr(ws.Cells(row,19).Value),"PCNF") = 0 Then
			delStr = delStr & row & ":" & row & ","
		End If
		row = row + 1
	Next
	
	delStr = Left(delStr,Len(delStr)-1)
	ws.Range(delStr).Select
	ex.Selection.Delete -4162
	
	Call updateEntries
	WScript.Echo entries
	
	
	'WScript.Sleep(50000)
	'wb.SaveAs(newFilePath)
	'WScript.Sleep(100)
	'fso.DeleteFile(filePath)
	
	
    

	
	
	
	
	ex.Save
	ex.Quit 
	
	Set fso = Nothing
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing

End Sub	

Sub updateEntries
	row = 2
	While ws.Cells(row,1).Value <> ""
		row = row + 1
	Wend
		maxRow=row
		row = 2
		entries = maxRow - row

End Sub
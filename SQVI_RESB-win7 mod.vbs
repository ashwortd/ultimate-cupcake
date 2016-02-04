' File:			ZMB52.vbs
' Author:		Jesse Colucci
' Edit Date:	08/14/2013
Option Explicit

' File Definitions
Const fileName = "SQVI_RESB_EXPORT"
Const fileDirectory = "\\winvault01\briodata\BrioReps\ETL\PMx_Export\"
Const tempDirectory = "C:\Temp\"
Const tempName = "Data_Export_Temp"
Const userName = "jtanguay"
Const password = "alstomspring2015"
Const showWindow = true		' Show excel window?



Const FOR_READING = 1		' File IO function inputs
Const FOR_WRITING = 2
Const xlAddIn = 18			' Excel save as file format input



' File locations
Dim alteredFileLocation, tempFileLocation, excelFileLocation
alteredFileLocation = fileDirectory & fileName & ".txt"
tempFileLocation = tempDirectory & tempName & ".txt"
excelFileLocation = fileDirectory & fileName & ".txt"

' Call to the main program
main()



sub main
	' Open SAP
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")

	' Not yet completed
	If not(WshShell.AppActivate("SAP Logon")) then
		WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")	'add the ' (x86)' to the program files directory dma02 3/5/2015
		Wscript.Sleep 500
		
		Dim i : i = 0
		Do While not(WshShell.AppActivate("SAP Logon"))
			WScript.Sleep 250
			timeoutCheck i, 400, "SAP Logon Timeout"		' Loop a max of 10 seconds
		Loop
	End if
	
	' Run GUI Script
	Dim application, SapGuiAuto, connection, session, isNewConn
	If Not IsObject(application) Then
	   Set SapGuiAuto  = GetObject("SAPGUI")
	   Set application = SapGuiAuto.GetScriptingEngine
	End If
	If Not IsObject(connection) Then
		If application.Children.Count > 0 then				' If it has connections
			Set connection = application.Children(0)
			isNewConn = false
			If not connection.description = "1.1 PMx Production (PE1)" Then		'Change the connenction from 'PE1' dma02 3/5/2015
				MsgBox connection.description
				Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true) 'Change the connenction from 'PE1' dma02 3/5/2015
				isNewConn = true
			End if
		Else
			Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true) 'Change the connenction from 'PE1' dma02 3/5/2015
			isNewConn = true
		End if
	End If
	If Not IsObject(session) Then
	   Set session = connection.Children(0)
	End If
	If IsObject(WScript) Then
	   WScript.ConnectObject session,     "on"
	   WScript.ConnectObject application, "on"
	End If
	session.findById("wnd[0]").maximize
	
	' Login
	If isNewConn then
		session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = userName
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
		session.findById("wnd[0]").sendVKey 0

		' If any messages come up clear them
		Dim messageCount, logonOption
		messageCount = 0
		Do while session.Children.Count > 1
			if messageCount > 5 then
				MsgBox "Error, too many message boxes detected"
				Wscript.quit
				exit do
			else
				Set logonOption = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1", false)
				' Check for message to bump off another person logged on
				if TypeName(logonOption) <> "Nothing" then
					logonOption.select
				End if
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End if
			messageCount = messageCount + 1
		Loop
		
		
	Else
		Dim sessionCount
		sessionCount = connection.Children.Count
		
		session.CreateSession
		do while connection.Children.Count <= sessionCount
			WScript.Sleep 250
		loop
		Set session = connection.Children(connection.Children.Count - 1)
	End If
	'session.LockSessionUI
	
	' Open Transaction
	session.StartTransaction("SQVI")
	'session.findById("wnd[0]/tbar[1]/btn[17]").press
	'session.findById("wnd[1]/tbar[0]/btn[8]").press
	



	'session.findById("wnd[0]/tbar[1]/btn[8]").press		' Execute query
	session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "RESB"
	session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 4
	session.findById("wnd[0]/usr/btnP1").press
	session.findById("wnd[0]/tbar[1]/btn[17]").press
	session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
	session.findById("wnd[1]/tbar[0]/btn[2]").press
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[4]").select

	
	' Send to local file
	session.findById("wnd[0]/tbar[1]/btn[45]").press
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = tempDirectory
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = tempName & ".txt"
	session.findById("wnd[1]/tbar[0]/btn[11]").press
	

	' Send to local file
	session.findById("wnd[0]/tbar[1]/btn[45]").press
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = fileDirectory
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fileName & ".txt"
	session.findById("wnd[1]/tbar[0]/btn[11]").press

	' Back up 
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	connection.CloseSession(session.Id)
			

	' Edit text file
	Dim objFS, objInS, objOutS, lineNum, lineString, objFS2, objFS3
	alteredFileLocation = fileDirectory & fileName & ".txt"
	tempFileLocation = tempDirectory & tempName & ".txt"
	
	' File Objects
	Set objFS = CreateObject("Scripting.FileSystemObject")			' Create temporary copy to read from
	Set objInS = objFS.OpenTextFile(tempFileLocation, FOR_READING, false, -1)
	Set objOutS = objFS.OpenTextFile(alteredFileLocation, FOR_WRITING, false)	' Path not found error
	
	lineNum = 1
	' Trim header lines and remove tab at start of line
	Do While objInS.AtEndOfStream <> TRUE
		lineString = objInS.ReadLine
		If lineNum < 6 Then								' Skip some of the first few lines
			If lineNum = 3 then
				objOutS.WriteLine Mid(lineString, 2)	' Copy column names line
			End If
			lineNum = lineNum + 1
		Else
			objOutS.WriteLine Mid(lineString, 2)		' Copy line over skipping starting tab
	   End If
	Loop
	
	objInS.Close
	objOutS.Close
	'objFS.DeleteFile tempFileLocation
	
	' Convert to .xls file
	Dim objExcel, objWorkbook
	

	Set objExcel = CreateObject("Excel.Application")			' Start Application
	objExcel.Application.Visible = showWindow
	Set objWorkbook = objExcel.Workbooks.Open(alteredFileLocation)		' Open File
	
	Dim fileOpen
	fileOpen = true
	On Error Resume Next
	objExcel.ActiveSheet.Columns("A").Delete
    	'objExcel.ActiveSheet.Columns("B").Delete
    	'objExcel.ActiveSheet.Rows("2").Delete
		Do while fileOpen
			objExcel.DisplayAlerts = False	' Ignore file type alert
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close ' closes the active workbook and saves any changes
			'objWorkbook.SaveAs excelFileLocation	' Save
			objExcel.DisplayAlerts = True
			If Err <> 0 Then		' File must already be open
				objExcel.Application.Visible = true
				selection = msgbox("The file at '" & excelFileLocation &"' is already open." & vbCrLf & _
					"Select Abort to Cancel, Retry to try again once the file is closed, or Ignore to enter new file name", _
					vbAbortRetryIgnore, "File already open")
				Select Case selection
				Case vbAbort
					fileOpen = false
					Wscript.Quit
				' Case vbRetry
					' fileOpen = true ' (continue loop)
				Case vbIgnore
					' change excelFileLocation
					excelFileLocation = InputBox("Enter new file name", "Save Excel File", excelFileLocation)
					If Right(excelFileLocation, 4) <> ".txt" Then excelFileLocation = excelFileLocation & ".txt"
				End Select
			Else
				fileOpen = false
			End If
			Err.Clear
		Loop
	On Error GoTo 0
	
	objExcel.Application.Quit			' Exit


 End sub

' Prompts user and then clears loopVar if user doesn't cancel
Sub timeoutCheck(loopVar, maxVal, title)
	loopVar = loopVar + 1		'Timeout var increase'
	if loopVar > maxVal then
		if isNull(title) then
			title = "Loop Timeout"
		End if
		OkCancelMsg "A loop has timed out. Press OK to continue or Cancel to exit", title
		loopVar = 0
	End if
End Sub
' File:			IT Tickets BO Export.vbs
' Author:		Anthony Ciccarello
' Change Date:	08/05/2013
Option Explicit
' Global Variables
Dim objIE, objExcel

'''''''''''''''''''' Main '''''''''''''''''''''''''
Call main()

Sub main
	' Variable definitions
	Dim WshShell, objWindow, objFrame, element, selection, showWindow, loopIterator
	Dim downloadFolder, savePath, fileIterator, fileElementID, fileSaveName, teamNames, colNames
	Const BO3ShortURL = "http://po1aps.ad.sys:8080/BOE/BI"
	Const BO3LoginURL = "http://po1aps.ad.sys:8080/BOE/BI"
	Const BO3DocumentsURL = "http://po1aps.ad.sys:8080/BOE/BI"
	Const username = "dma02"
	Const password = "Morgan13"		' Needs to be set to Jesse's BO 3 password
	showWindow = true		' Only show applications if running from console
		 ' "CSCRIPT.EXE" = UCase( Right( WScript.Fullname, 11 ) )	'Can use if running from command line
	
	set WshShell = WScript.CreateObject("WScript.Shell")
	downloadFolder = WshShell.SpecialFolders("mydocuments") & "\"							' Where the file is saved from IE
	fileSaveName = Array("ASC - KPI Reporting.xls", "IT Incidents.xls", "GI6 - Problem General Information.xls")						' Desired file name
	fileElementID = Array("ListingURE_listColumn_0_0_1", "ListingURE_listColumn_2_0_1", "ListingURE_listColumn_3_0_1")		' Web page element IDs
	
	'Define Excel file format
	savePath = downloadFolder & fileSaveName(1)		' Have excel write over the original	' Where excel doc is saved to
	colNames = Array("Incident Number", "Service Request Number", "Company", "Requester First Name", "Requester Last Name", _
		"Summary", "Notes", "Status", "Priority", "Assigned Group", "Assignee", "Service Level", "Submit Date", _
		"Responded Date", "Last Modified Date", "Last Modified By", "Next Action Date", "Pending Date", "Last Work Info Summary", _
		"Last Work Info Notes")
	teamNames = Array("BELFORTI", "CICCARELLO", "DIMAURO", "ST-JOHN", "TANGUAY-COLUCCI", "WRIGHT")
	
	

	' Get IE Object for BO 3 Page
	ConsoleOut "Starting IE"
	If not(getIE(objIE, BO3ShortURL)) then
		' Create a new window '
		Set objIE = WScript.CreateObject("InternetExplorer.Application")
		objIE.Navigate BO3LoginURL		' Navigate to login
		objIE.Visible = showWindow		' Set visibility
		PageWait objIE, 500
		consoleOut "IE Created"
	Else	' If found a BO3 Window already open
		consoleOut "IE window found"
		objIE.Visible = True
		WshShell.AppActivate("Windows Internet Explorer")
		On Error Resume Next
		objIE.Navigate BO3DocumentsURL	' Start at documents page
		If Err Then		' Unable to get resource because browser has pop up
			ConsoleOut "ERROR: " & Err & " when trying to navigate to webpage"
			WshShell.AppActivate("BusinessObjects InfoView - Windows Internet Explorer")
			Msgbox "Please close the browser window and restart script"
			Wscript.Quit
		End if
		On Error Goto 0
		PageWait objIE, 250
	End If
		
	' Check if logged in already (will go directly to docs)
	If getIE(objIE, BO3DocumentsURL) then
		consoleOut "Already logged in"
	Else 
		getIE objIE, BO3LoginURL
		
		' Enter login credentials
		Set objFrame = getSubFrame(objIE.Document.parentWindow, 0,	"infoView_home")
		MsgBox(objFrame)
		getElement(objFrame, "usernameTextEdit").value = username
		getElement(objFrame, "passwordTextEdit").value = password
		getElement(objFrame, "buttonTable").firstChild.click
		consoleOut "Logging In"
	End if
	objIE.Visible = showWindow		' Set visibility
	'objIE.silent = true
	PageWait objIE, 1000
	
	' Open Files and download
	for fileIterator = 0 to Ubound(fileSaveName)
		' Get page frames
		Set objFrame = getSubFrame(objIE.Document.parentWindow, 0,	"headerPlusFrame")
		Set objFrame = getSubFrame(objFrame, 0,	"dataFrame")
		Set objFrame = getSubFrame(objFrame, 0,	"workspaceFrame")
		Set objFrame = getSubFrame(objFrame, 1,	"workspaceBodyFrame")
		' Double click document to open
		Set element = getElement(objFrame, fileElementID(fileIterator)).parentNode
		element.fireevent "onmousedown"
		element.fireevent "ondblclick"
		PageWait objIE, 250
		consoleOut "Document Opened"
		
		' Refresh the data
		Set objFrame = getSubFrame(objFrame, 0,	"webiViewFrame")
		getElement(objFrame, "iconMenu_icon_refreshAll").click	' Click refresh button
		PageWait objIE, 250
		consoleOut "Query Options Loaded"
		getElement(objFrame, "promptsOKButton").click 	' Select "Run Query"
		PageWait objIE, 1000
		
		' Wait for query to run (loop while the wait Dialogue box is visible)
		loopIterator = 0
		Do while getElement(objFrame, "waitDlg").currentStyle.visibility = "visible"
			Wscript.sleep 500
			timeoutCheck loopIterator, 200, "Query load timeout"	' Check for loop timeout
		Loop
		Wscript.Sleep 250
		consoleOut "Query Run"

		' Select save to excel file
		getElement(objFrame, "iconMenu_icon_docMenu").click							' Select Document Actions
		getElement(objFrame, "iconMenu_menu_docMenu_item_saveDocComputerAs").click	' Hover on "Save to my computer as"
		getElement(objFrame, "saveDocComputerMenu_item_saveXLS").click				' Select Excel
		consoleOut "Saving File"
		
		'''''''' Save it locally '''''''''''''
		' Wait for right download status
		Do while objIE.busy = False or objIE.ReadyState <> 1
			Wscript.Sleep 100
			timeoutCheck loopIterator, 200, "Cannot find the download window"	' Check for loop timeout
		Loop
		
		' Loop while File Download window not active
		loopIterator = 0
		Do While Not WshShell.AppActivate("File Download")
			consoleOut VBTab & "Busy: " & objIE.Busy & ";" & VBTab & "State: " & objIE.ReadyState
			
			Wscript.Sleep 100												' Wait
			timeoutCheck loopIterator, 200, "Cannot find 'File Download' window"	' Check for loop timeout
		Loop
		consoleOut VBTab & "Window Found: " & VBTab & "File Download"
		
		' Select Save
		loopIterator = 0
		Do While objIE.ReadyState = 1 and WshShell.AppActivate("File Download")
			Wscript.Sleep 500
			WshShell.SendKeys "%{s}"
			consoleOut VBTab & "Sent Alt + S"
			Wscript.Sleep 250
			timeoutCheck loopIterator, 10, "File download window open window"	' Check for loop timeout
		Loop
		consoleOut VBTab & "Busy: " & objIE.Busy & ";" & VBTab & "State: " & objIE.ReadyState
		
		' Wait for "Save As" to become active
		loopIterator = 0
		Do While Not WshShell.AppActivate("Save As")
			Wscript.Sleep 100									' Wait
			ConsoleOut VBTab & "Waiting for 'Save As' Box"
			timeoutCheck loopIterator, 100, "Cannot find 'Save As' window"	' Check for loop timeout
		Loop
		
		' Save Document
		loopIterator = 0
		consoleOut VBTab & "Window Found: " & VBTab & "downloadPDForXLS"
		Do While WshShell.AppActivate("downloadPDForXLS.jsp from reporting.itc.alstom.com Completed")
			If WshShell.AppActivate("Save As") Then
				consoleOut VBTab & "Selecting save location"
				Wscript.Sleep 250
				
				' Set file location
				WshShell.SendKeys "%n"
				Wscript.Sleep 250
				WshShell.SendKeys downloadFolder & fileSaveName(fileIterator) & "{ENTER}"
				Wscript.Sleep 1000
				
				' Clear "Do you want to replace?" dialogue
				If WshShell.AppActivate("0% of downloadPDForXLS.jsp") Then
					Wscript.Sleep 250
					WshShell.SendKeys "%y"
					Wscript.Sleep 500
				End If
			Else
				' Test if still downloading
				consoleOut VBTab & "Waiting for download to complete"
				Wscript.Sleep 250									' Wait
			End If
			timeoutCheck loopIterator, 5, "Download window still open"	' Check for loop timeout
			
			' Check for file overwrite error box
			Do While WshShell.AppActivate("Error Copying File or Folder") <> 0
				objIE.Visible = true
				consoleOut "Error window while trying to save file"
				' Give user an option of how to proceed
				selection = msgbox("The file at '" & downloadFolder & fileSaveName(fileIterator) &"' is already open." & vbCrLf & _
					"Please close the file or select cancel", vbRetryCancel, "File already open")
					
				WshShell.AppActivate("Error Copying File or Folder")
				Wscript.Sleep 250
				WshShell.SendKeys "{ENTER}"
				Wscript.Sleep 250
				If selection = vbCancel Then
					WshShell.SendKeys "%y"	' Cancel download
					Wscript.Sleep 500
					Wscript.Echo "File download cancelled. You will need to rerun the script once you have closed the excel file."
					ConsoleOut "Script Exited"
					Wscript.Quit
				Else
					WshShell.SendKeys "%n"
					Wscript.Sleep 500
				End If
			Loop
		Loop
		

		Wscript.Sleep 250
		
		' Clear "Download Complete"
		If WshShell.AppActivate("Download complete") Then
			Wscript.Sleep 250
			WshShell.SendKeys "{ENTER}"
		End if
		
		' Close the document
		consoleOut "Closing the document"
		objFrame.execScript "_dontCloseDoc = true;", "javascript"		' Disable page change prompt box
		objIE.Document.parentWindow.execScript "window.findFrame('workspaceHeaderFrame').onClose();", "javascript"
		
	Next 	' Repeat for second report
	
	' Exit Browser
	objIE.Quit
	ConsoleOut "Closing Browser"
	WScript.Sleep 500
	
	consoleOut "Browser Closed"
	
	' Clear "Download Complete"
	If WshShell.AppActivate("Download complete") Then
		Wscript.Sleep 250
		WshShell.SendKeys "{ENTER}"
		Wscript.Sleep 250
	End if
	
	Set objIE = Nothing
	
	
	
	''''''''''''''''''''''''''''''''''''' Edit Excel File '''''''''''''''''''''''''''''''''''''''''
	' Define Constants
	Dim objWorkbook, objSheets, objWorksheet
	Const xlOr = 2					' Define argument for AutoFilter
	Const xlFilterValues = 7		' Define argument for AutoFilter multiple select
	Const xlShiftToRight = -4161	' Define argument for insert 
	' Const xlWorkbookNormal = -4143	' Define file format argument
	
	consoleOut vbNewLine & "Open Excel"
	Set objExcel = CreateObject("Excel.Application")			' Start Application
	objExcel.Application.Visible = showWindow
	objExcel.Workbooks.Open(downloadFolder & fileSaveName(0))
	objExcel.Workbooks.Open(downloadFolder & fileSaveName(2))
	Set objWorkbook = objExcel.Workbooks.Open(downloadFolder & fileSaveName(1))		' Open File
	Set objSheets = objWorkbook.Sheets
	consoleOut "File Opened"
	
	' Delete Sheets
	objExcel.DisplayAlerts = False			' Ignore loss of data alert
	objSheets("Presentation Page").delete	' Delete title page
	objSheets("Presentation").delete		' Delete page outlining query
	objExcel.DisplayAlerts = True
	
	' Rearrange Columns
	Set objWorksheet = objSheets("General Information")
	Dim iRow, iCol, index, colTitle, TargetCol
	'Constant values
	Set objWorksheet = objSheets("General Information")
	'target_sheet = "Final Report" 'Specify the sheet to store the results
	iRow = objWorksheet.UsedRange.Rows.Count 'Determine how many rows are in use
	'Create a new sheet to store the results
	'Worksheets.Add.Name = "Final Report"
	'Start organizing columns
	For index = 0 to UBound(colNames) - 1
		colTitle = colNames(index)
		For iCol = 1 To objWorksheet.Columns.Count
			'Sets the TargetCol to zero in order to prevent overwriting existing target columns
			TargetCol = 0
			'Read the header of the sheet to determine the column order
			If colTitle = objWorksheet.Cells(1, iCol).Value Then TargetCol = index + 1
			
			'If a TargetColumn was determined (based upon the header information) then copy the column to the right spot
			If TargetCol <> 0 Then
				'colWidth = objWorksheet.Columns(iCol).ColumnWidth
				objWorksheet.Columns(iCol).Cut
				objWorksheet.Columns(TargetCol).Insert xlShiftToRight
				ConsoleOut VBTab & "Moved Column: " & colNames(TargetCol - 1)
				Exit For		' Move on to next column name
			End If
		Next
	Next	'Move to the next column until all columns are read
	ConsoleOut "Columns Rearranged"
	
	' Find colomn Positions
	Dim colName, lastNameCol, summaryCol, priorityCol, statusCol, targetStatusCol
	For iCol = 1 To objWorksheet.Columns.Count
		colName = objWorksheet.Cells(1, iCol).Value
		If colName = "Requester Last Name" Then lastNameCol = iCol
		If colName = "Summary" Then summaryCol = iCol
		If colName = "Priority" Then priorityCol = iCol
		If colName = "Status" Then statusCol = iCol
		If colName = "SLA Fix 1 SLT Status" Then targetStatusCol = iCol
	Next
	ConsoleOut "Filter Columns Found"
	
	' Filter to team
	objWorksheet.Range("A1").AutoFilter lastNameCol, teamNames, xlFilterValues
	consoleOut "Team incidents filtered"
	
	' Create Data Quality Tab
	objSheets("General Information").Copy ,objWorkSheet 					' After objSheets("General Information")
	Set objWorksheet = objWorksheet.next 									' Get new sheet
	objWorksheet.name = "Data Quality"										' Rename
	objWorksheet.Range("A1").AutoFilter summaryCol, "*BW:Data Quality*"		' Filter on Data quality issues
	objWorksheet.Range("A1").AutoFilter priorityCol, "Low"					' Filter on low priority
	objWorksheet.Range("A1").AutoFilter statusCol, "<>Resolved"				' Filter out resolved issues
	consoleOut "Data Quality Tab Created"
	
	' Create "Service target breached" tab
	objSheets("General Information").Copy ,objWorksheet						' After objSheets("Data Quality")
	Set objWorksheet = objWorksheet.next 									' Get new sheet
	objWorksheet.name = "Service Target Breached"							' Rename
	objWorksheet.Range("A1").AutoFilter statusCol, "<>Resolved"							' Filter out resolved issues
	objWorksheet.Range("A1").AutoFilter targetStatusCol, "Missed", xlOr, "Missed Goal"	' Filter on Target Breached
	consoleOut "Service Target Breached Tab Created"
	
	' Show finished spreadsheet
	objExcel.Application.Visible = true
	
	Dim fileOpen
	fileOpen = true
	On Error Resume Next
		Do while fileOpen
			objExcel.DisplayAlerts = False	' Ignore file type alert
			objWorkbook.SaveAs savePath		' Unable to update file format , xlWorkbookNormal	' Save
			objExcel.DisplayAlerts = True
			If Err <> 0 Then
				consoleOut "ERROR: " & Err & " While trying to save file"
				selection = msgbox("The file at '" & savePath &"' is already open." & vbCrLf & _
					"Select Abort to Cancel, Retry to try again once the file is closed, or Ignore to enter new file name", _
					vbAbortRetryIgnore, "File already open")
				Select Case selection
				Case vbAbort
					fileOpen = false
					consoleOut "Save Cancelled. Script will exit"
					Wscript.Quit
				' Case vbRetry
					' fileOpen = true ' (continue loop)
				Case vbIgnore
					' change savePath
					savePath = WshShell.SpecialFolders("mydocuments") & "\" & InputBox("Enter new file name", _ 
						"Save IT Ticket data", "IT Incidents2")
					If Right(savePath, 4) <> ".xls" Then savePath = savePath & ".xls"
				End Select
			Else
				fileOpen = false
				consoleOut "File Saved"
			End If
			Err.Clear
		Loop
	On Error GoTo 0
	MsgBox "IT Ticket Export Complete"
	consoleOut "Script Finished"
End Sub






''''''''''''''''''''' Methods ''''''''''''''''''''
' Prompt with an option to cancel '
Sub OkCancelMsg(msg, title)
	Dim result
	result = MsgBox(msg, vbOKCancel, title)	' Give user option to cancel
	If result = vbCancel then
		objIE.Visible = True					' Close Internet Explorer
		consoleOut "Script Terminated"
		Wscript.Quit
	End If
	ConsoleOut VBTab & "Continuing script"
End Sub

' Wait for a page to load '
Sub PageWait(IE, waitTime)
	Wscript.Sleep 250
	consoleOut VBTab & "Busy: " & IE.Busy & ";" & VBTab & "State: " & IE.ReadyState & " (Initial)"
	
	Do While IE.Busy Or IE.ReadyState <> 4
		WScript.Sleep waitTime
		consoleOut VBTab & "Busy: " & IE.Busy & ";" & VBTab & "State: " & IE.ReadyState
	Loop
	Wscript.Sleep 250
End Sub

' Wait for sub frame to appear and return the frame object selected by index among sibling windows
Function getSubFrame(objFrame, index, frameName)
	Dim loopIterator : loopIterator = 0
	' Loop while no frames found
	On Error Resume Next
	Do while objFrame.frames.length <= index	
		If Err Then																' ERR "Permission denied: 'objFrame.frames.length'"
			ConsoleOut "ERROR: " & Err & " While trying to find " & objFrame.name & " sub-fames"
			Err.Clear															' Clear Error and continue
		End If	
		Wscript.Sleep 100														' Wait
		timeoutCheck loopIterator, 80, "No frames in " & objFrame.name						' Check for loop timeout
	Loop
	If Err Then ConsoleOut "ERROR: " & Err & " While trying to find " & frameName
	
	' Loop while name isn't found
	loopIterator = 0
	Do while objFrame.frames(index).name <> frameName
		If Err Then															' ERR "Member not found." x5
			ConsoleOut "ERROR: " & Err & " While trying to find " & frameName
			Err.Clear																' Clear Error and continue
		End If
		Wscript.Sleep 100													' Wait
		timeoutCheck loopIterator, 50, frameName & " not found in " & objFrame.name	' Check for loop timeout
	Loop
	If Err Then ConsoleOut "ERROR: " & Err & " While trying to find " & frameName
	On Error Goto 0
	consoleOut VBTab & "Found frame: " & VBTab & frameName
	'Return Value
	Set getSubFrame = objFrame.frames(index)
End Function

' Wait for element to load and return (found by elemID)
Function getElement(objFrame, elemID)
	Dim loopIterator : loopIterator = 0
	
	' Loop while element isn't found
	On Error Resume Next
	Do While Not isValidElement(objFrame.Document.getElementById(elemID))
		If error <> 0 Then
			ConsoleOut "ERROR: " & Err & " While trying to find " & elmID
			Err.Clear																' Clear Error and continue
		End if
		Wscript.Sleep 100								' Wait
		timeoutCheck loopIterator, 100, elemID & " not found"		' Check for loop timeout
	Loop
	If Err <> 0 Then ConsoleOut "ERROR: " & Err & " While trying to find " & frameName
	On Error Goto 0
	consoleOut VBTab & "Found element: " & VBTab & elemID
	'Return Value
	Set getElement = objFrame.Document.getElementById(elemID)
End Function

' Get IE object based on URL '
Function getIE(objIE, url)
	getIE = false

	For each objIE In CreateObject("Shell.Application").Windows
		If InStr(objIE.LocationURL, url) Then
			getIE = true
			Exit For
		End If
	Next
	if not(getIE) then
		objIE = Null
	end if
End Function

' Prompts user and then clears loopVar if user doesn't cancel
Sub timeoutCheck(loopVar, maxVal, title)
	loopVar = loopVar + 1		'Timeout var increase'
	if loopVar > maxVal then
		objIE.Visible = True	' Allow the page to be seen
		if isNull(title) then
			title = "Loop Timeout"
		End if
		consoleOut "Loop Timeout: " & title
		OkCancelMsg "A loop has timed out. Press OK to continue or Cancel to exit", title
		loopVar = 0
	End if
End Sub

Function isValidElement(element)
	If IsNull(element) Or TypeName(element)="Nothing" Then
		isValidElement = false
	Else
		isValidElement = true
	End if
End function

Sub consoleOut(message)
	if "CSCRIPT.EXE" = UCase( Right( WScript.Fullname, 11 ) ) then
		Wscript.Echo message
	End if
End sub

' Close IE and script '
Sub ExitScript(IE)
	If not IsEmpty(IE) Then
		IE.Quit
	End If
	Set IE = Nothing
	WScript.Quit
End Sub
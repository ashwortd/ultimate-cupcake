Const fileName = "zmb52_export"
Const fileDirectory = "C:\Users\dma02\Desktop\Daily Reports\"
Const tempDirectory = "C:\Temp\"
Const tempName = "Data_Export_Temp"
Const userName = "rbelfort"
Const password = "alstomsummer2015"
Const showWindow = true		' Show excel window?



Const FOR_READING = 1		' File IO function inputs
Const FOR_WRITING = 2
Const xlAddIn = 18			' Excel save as file format input

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
		If lineNum < 7 Then								' Skip some of the first few lines
			If lineNum = 6 then
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
    	'objExcel.ActiveSheet.Columns("A").Delete
    	objExcel.ActiveSheet.Columns("B").Delete
    	objExcel.ActiveSheet.Rows("2").Delete
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
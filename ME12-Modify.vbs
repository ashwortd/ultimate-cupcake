'===============================================================================================================
'ME12 Modify Purchasing Info Record
'Derek Ashworth
'Creation Date: 6-9-2016	
'Data File: Purchase_Info_Record_Modify_Form.xlsx and modifies existing purchasing info records.

'===============================================================================================================
Dim minQty,x, xlValue
'My Functions ----------------------------------
Function valCheck(x)
 If ExcelSheet.Cells(Row,x).Value<>"" Then
 	valCheck=True
 	xlValue=ExcelSheet.Cells(Row,x).Value
  Else
   valCheck=False
  End If
 End Function
 
 Function wndTitle(r)
 	wndTitle=session.findbyid("wnd["&r&"]").text
 End Function
 
Function saveRecord()
	session.findbyid("wnd[0]/tbar[0]/btn[11]").press
End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function
'End of My Functions -----------------------------------

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If

If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If

If Not IsObject(session) Then
	Set session = connection.Children(0)
	Dim instance
	If connection.Children.Count < 6 Then
		session.createSession
		WScript.sleep 3000
		instance = connection.Children.Count
		instance = instance - 1
		Set session = connection.Children(CInt(instance))
	
	Else Do
		 ans = InputBox("You have too many SAP sessions open to create another. Which session should I use?" &vbNewLine&_
					"Remember the first session is 0, second 1 etc.")
	 	 ans = CInt(ans)
			If ans < 6 Then
				If ans > -1 Then
					ansCheck = 1
					Set session = connection.Children(CInt(ans))
				End If
			End	If
		Loop While ansCheck <> 1
	End If
End If
 
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
Set objShell = CreateObject("Wscript.Shell")
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,StatusBar,i,vendorCheck,vendorNumber,materialNumber,StatusBar2,lastRow,remaining
Dim validFrom,validFromCheck,validTo,validToCheck,purchOrg,purchOrgCheck,plant,plantCheck,priceCheck,leadTime

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ Don't change anything above ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^

' 1. - Ask user to choose file, then open file.
		file = ChooseFile(defaultLocalDir)
		MsgBox file
		Function ChooseFile (ByVal initialDir)
		    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
		    Set shell = CreateObject( "WScript.Shell" )
		    Set ex = shell.Exec( "mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
		    ChooseFile = Replace( ex.StdOut.ReadAll, vbCRLF, "" )
		        Set ex = Nothing
		        Set shell = Nothing
		End Function
		Set ExcelApp = CreateObject("Excel.Application")
		ExcelApp.Visible=True
		Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
		
' 2. - Tell script which row to start with.
		Row=InputBox("Which row would you like to start with?","PIR Modification")
			If Row = False Then
				Call Endscript
			End If
		Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
' 3. - Open transaction ME12 in the first instance of SAP
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nme12"
		session.findById("wnd[0]").sendVKey 0

' 4. - Initiate main function loop. The script checks first column of the current row (row changes after every loop) for a value, if cell is blank then the script stops.
				Do
				Call Main
				Row=Row+1
				Loop While ExcelSheet.Cells(Row,1).Value <>""
				Call Endscript

' 5. - Enter values from the spreadsheet into the first screen of ME11 and press enter.
	Sub Main
		session.findById("wnd[0]/usr/ctxtEINA-LIFNR").text = ExcelSheet.Cells(Row,1).Value 	'Vendor Number   "10014015"
		session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = ExcelSheet.Cells(Row,2).Value 	'Material Number "GRUV-LUb"
		session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ExcelSheet.Cells(Row,3).Value 	'Purchasing Org  "US36"
		session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ExcelSheet.Cells(Row,4).Value 	'Plant		  "500b"
		session.findById("wnd[0]/usr/ctxtEINA-INFNR").text=""
		session.findbyid("wnd[0]/usr/radRM06I-NORMB").select
		session.findById("wnd[0]").sendVKey 0									   	'Press Enter

' 5E. - Check for and log any errors.
			StatusBar = session.findById("wnd[0]/sbar").text					'Copies the statusbar text to evaluate
			
				If session.findById("wnd[0]/sbar").MessageType = "E" Then		'If the error message is type "E - Error"
					ExcelSheet.Cells(Row,19).Value = StatusBar				'Copy error message to the spreadsheet
					Call errorSub										'Cancel all input, go to the next row and start over
					Exit Sub
				End If
				If session.findById("wnd[0]/sbar").MessageType = "W" Then
					ExcelSheet.Cells(Row,19).Value = StatusBar
					Call errorSub
					Exit Sub
				End If



'Modify
	If Right(wndTitle(0),12)="General Data" Then
		session.findById("wnd[0]/tbar[1]/btn[7]").press	
	End If
		   For i=5 To 10
		   If valCheck(i)=True Then
		   	Select Case i
		   		Case 5
		   			session.findById("wnd[0]/usr/txtEINE-APLFZ").text = xlValue
		   		Case 6
		   			session.findById("wnd[0]/usr/txtEINE-UNTTO").text = xlValue
		   		Case 7
		   			session.findById("wnd[0]/usr/txtEINE-UEBTO").text = xlValue
		   		Case 8
		   			session.findById("wnd[0]/usr/ctxtEINE-EKGRP").text = xlValue
		   		Case 9
		   			session.findById("wnd[0]/usr/txtEINE-NORBM").text = xlValue
		   		Case 10
		   			session.findById("wnd[0]/usr/txtEINE-MINBM").text = xlValue
		   	End Select
		   				
		   End If
		   Next  
	If Right(wndTitle(0),12)<>"General Data" Then
		session.findById("wnd[0]/tbar[1]/btn[6]").press
	End If
		For z =11 To 13
		 If valCheck(z)=True Then
		  Select Case z
		  	Case 11
		  		session.findById("wnd[0]/usr/txtEINA-MAHN1").text = xlValue
		  	Case 12
		  		session.findById("wnd[0]/usr/txtEINA-MAHN2").text = xlValue
		  	Case 13
		  		session.findById("wnd[0]/usr/txtEINA-MAHN3").text = xlValue
		  	end Select
		  End If
		 Next
saveRecord
ExcelSheet.Cells(Row,14).Value= sbarStatus
	If session.findById("wnd[0]/sbar").MessageType = "E" Then 		
		Call errorSub
		Exit Sub
	End If	

' 8E. - After saving check for, and log any errors to the spreadsheet.
		StatusBar = session.findById("wnd[0]/sbar").text
		ExcelSheet.Cells(Row,19).Value= StatusBar
	
		If session.findById("wnd[0]/sbar").MessageType = "E" Then 		'If there is a serious error, cancel out and start again on the next row
			Call errorSub
			Exit Sub
		End If
End Sub


' HELPER SUBS - Do not modify anything below.
			Sub errorSub
				session.findById("wnd[0]/tbar[0]/btn[15]").press 		 'Press exit. No changes are saved.
				session.findById("wnd[0]/tbar[0]/okcd").text = "/nme12" 'Reopen ME11
				session.findById("wnd[0]").sendVKey 0 				 'Press Enter
			End Sub

			Sub Endscript
			MsgBox("Script completed")
				ExcelWorkbook.Close(True)
				ExcelApp.Quit
				Set ExcelApp=Nothing
				Set ExcelWorkbook=Nothing
				Set ExcelSheet=Nothing
				WScript.ConnectObject session,     "off"
				WScript.ConnectObject application, "off"
				WScript.Quit
			End Sub
'----------EOF
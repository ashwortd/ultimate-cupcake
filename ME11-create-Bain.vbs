'===============================================================================================================
'                                                                             ME11 Create Purchasing Info Record
'                                                                                             Eric Bain 06-02-15
'
'Script takes data from Create PIR spreadsheet and creates new purchasing info records.
'It has error checking to verify that the price, lead time and validity period are all accurate.
'Please read instructions included in the ME11 PIR template.xlsx
'===============================================================================================================
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
' - Use this to set a specific file and workbook
		'Set ExcelWorkbook = ExcelApp.Workbooks.Open("C:\Users\264711\Documents\PIR Updates\Fonderie Saguenay\Updates\Create PIR2.xlsx")
		'Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")



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
		Row=InputBox("Which row would you like to start with?","PIR Creation")
			If Row = False Then
				Call Endscript
			End If
		Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
' 3. - Open transaction ME11 in the first instance of SAP
		session.findById("wnd[0]").maximize
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11"
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
		session.findById("wnd[0]/usr/ctxtEINA-MATNR").text = ExcelSheet.Cells(Row,6).Value 	'Material Number "GRUV-LUb"
		session.findById("wnd[0]/usr/ctxtEINE-EKORG").text = ExcelSheet.Cells(Row,3).Value 	'Purchasing Org  "US36"
		session.findById("wnd[0]/usr/ctxtEINE-WERKS").text = ExcelSheet.Cells(Row,2).Value 	'Plant		  "500b"
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

' 6. - Open info Purchase Org data tab and set the price and minimum order quantity.
	priceCheck = ExcelSheet.Cells(Row,11).Value
		If priceCheck <> "NA" Then
		   session.findById("wnd[0]/tbar[1]/btn[7]").press									'Click Purchase Org tab
		   session.findById("wnd[0]/usr/txtEINE-NORBM").text = ExcelSheet.Cells(Row,13).Value 		'Enter minimum order quantity
		   
' 6E. - Sometimes the price field is grayed out. This will attempt to enter the price and then double check that it is correct.
	  'If it doesn't pass the second check it will enter it again.
		   On Error Resume Next															
		   session.findById("wnd[0]/usr/txtEINE-NETPR").text = ExcelSheet.Cells(Row,11).Value 	'Enter price from spreadsheet
		   priceCheck = session.findById("wnd[0]/usr/txtEINE-NETPR").text					'Copy price field in SAP to make sure price went in
		   If priceCheck = "" Then													'If priceCheck is blank then enter price again
			   session.findById("wnd[0]/usr/tblSAPMV13ATCTRL_D0201/txtKONP-KBETR[2,0]").text = ExcelSheet.Cells(Row,11).Value	
		   End If
		   
' 7. - Verify that there is a lead time. If no lead time, enter default 7 weeks.		  
		      leadTime = session.findById("wnd[0]/usr/txtEINE-APLFZ").text 'Read field where lead time should be.
		   If leadTime < 1 Then
		      session.findById("wnd[0]/usr/txtEINE-APLFZ").text = 49	  'If it's empty or 0, enter 7 weeks.
		   End If
		   If leadTime ="?" Then
		      session.findById("wnd[0]/usr/txtEINE-APLFZ").text = 49
		   End If
		   
' 8. - Set the purchasing group and the condition's validity period.
		   session.findById("wnd[0]/usr/ctxtEINE-EKGRP").text = "ELV"						'Enter default Purchasing Group
		   session.findById("wnd[0]/tbar[1]/btn[8]").press								'Open Conditions
		   session.findById("wnd[0]/usr/ctxtRV13A-DATAB").text = ExcelSheet.Cells(Row,15).Value   '"07/01/2014"
		   session.findById("wnd[0]/usr/ctxtRV13A-DATBI").text = ExcelSheet.Cells(Row,16).Value   '"07/01/2015"
		   session.findById("wnd[0]/tbar[0]/btn[11]").press								'Click save. Saving the new info record
		End If 'End of the priceCheck <> "NA" statement

' 8B. - If we don't have pricing but need to create the master data
		If priceCheck = "NA" Then													
		   On Error Resume Next	
		   session.findById("wnd[0]/tbar[0]/btn[11]").press
		   On Error Resume Next															
		   session.findById("wnd[0]/usr/txtEINE-NORBM").text = ExcelSheet.Cells(Row,13).Value 		'Enter minimum order quantity
		   On Error Resume Next															
		   session.findById("wnd[0]/tbar[0]/btn[11]").press
		   On Error Resume Next	
		   StatusBar = session.findById("wnd[0]/sbar").text
	
			If StatusBar = "Fill in all required entry fields" Then 		'Price required
			session.findById("wnd[0]/usr/txtEINE-NETPR").text = "0.01"
			ExcelSheet.Cells(Row,18).Value= "Flag for del"
			session.findById("wnd[0]/tbar[0]/btn[11]").press
			End If
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
				session.findById("wnd[0]/tbar[0]/okcd").text = "/nme11" 'Reopen ME11
				session.findById("wnd[0]").sendVKey 0 				 'Press Enter
					StatusBar=""								 'Reset the statusbar checking variable
					Row=Row+1									 'Tell script to go to the next row in the spreadsheet
				Call Main										 'Restart the main sub
				wscript.sleep 500
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
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
Dim ex,wb,ws,partFile
Dim vrc,row,workingRow,sapRow
Dim part,plant,status,continue,result,lotSize,altBom
Dim setupTotal,laborTotal,maxLead,wrkCtr,setupHours

Call Main

Sub Main
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nCS03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	Do
		continue = "Yes"
		Do
			continue = "Yes"
			part = UCase(InputBox("Enter part number below:", "Part Number"))
			If Len(Trim(part)) > 18 Or Len(Trim(part)) = 0 Then
				result = MsgBox("Script only supports part numbers greater than 0 and less than 18 characters."&Chr(13)&Chr(13)&"Try again?", 20)
				Select Case result
				Case vbYes 
					continue = "No"
				Case vbNo
					MsgBox("The script is now terminating.")
					WScript.ConnectObject session,     "off"
					WScript.ConnectObject application, "off"
					WScript.Quit
				End Select
			End If
		Loop While continue = "No"
		
		Session.findById("wnd[0]/usr/ctxtRC29N-MATNR").text = part
		
		Do
			continue = "Yes"
			plant = InputBox("Enter plant below:", "Plant",Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text)
			plant = UCase(Trim(plant))
			Select Case plant
			Case "500C"
				Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "500C"
			Case "500D"
				Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "500D"
			Case "500E"
				Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "500E"
			Case "500F"
				Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "500F"
			Case "500G"
				Session.findById("wnd[0]/usr/ctxtRC29N-WERKS").text = "500G"
			Case Else
				result = MsgBox("Script only supports plants 500C-500G"&Chr(13)&Chr(13)&"OK to try again."&Chr(13)&"Cancel to terminate script.", 17)
				Select Case result
				Case vbOK 
					continue = "No"
				Case vbCancel
					MsgBox("The script is now terminating.")
					WScript.ConnectObject session,     "off"
					WScript.ConnectObject application, "off"
					WScript.Quit
				End Select
			End Select
		Loop While continue = "No"
		
		Do
			lotSize = InputBox("Enter the number of assemblies below:", "Lot Size", 1)
			'*************TO DO***********TO DO****
			'***TO DO******************************
			'**********************TO DO***********   need to implement error checking to check if user enters integer
			'***********TO DO**********************
			'******************TO DO***************
			'***TO DO*********************TO DO****
		Loop While continue = "No"
	
		session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		
		If session.findById("wnd[0]/titl").text = "Display material BOM: Alternative Overview" Then
			row = 0
			While session.findById("wnd[0]/usr/tblSAPLCSDITCALT/txtRC29K-STLAL[0,"&row&"]").text <> "__"
				row = row + 1
			Wend
			altBom = InputBox("Select the BoM you wish to use from the available BoMs 1 -  "&row,"Alternative BoMs")
			'*************TO DO***********TO DO****
			'***TO DO******************************
			'**********************TO DO***********   need to implement error checking to check if user enters integer
			'***********TO DO**********************   & if said integer is a valid alternate BoM number
			'******************TO DO***************
			'***TO DO*********************TO DO****
			session.findById("wnd[0]/tbar[0]/btn[3]").press
			session.findById("wnd[0]/usr/txtRC29N-STLAL").text = altBom
			session.findById("wnd[0]/tbar[0]/btn[0]").press
		End If
		status = session.findById("wnd[0]/sbar").Text
		If status <> "" Then
			result = MsgBox("ERROR:"&Chr(13)&status&Chr(13)&Chr(13)&"Try again?", 20)
			Select Case result
				Case vbYes 
					continue = "No"
				Case vbNo
					MsgBox("The script is now terminating.")
					WScript.ConnectObject session,     "off"
					WScript.ConnectObject application, "off"
					WScript.Quit
				End Select
		End If
	
	Loop While continue = "No"
	Call extractBomWithCosts
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	MsgBox("Done.")
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End Sub



Sub extractBomWithCosts
	

	Set ex = WScript.CreateObject("Excel.application")
	ex.Visible = False
	
	Set wb = ex.Workbooks.Add
	partFile = part
	If InStr(part,"/") <> 0 Then
		partFile = Replace(partFile,"/","-") & ".xlsx"
	ElseIf InStr(part,"\") <> 0 Then
		partFile = Replace(partFile,"\","-") & ".xlsx"
	Else
		partFile = partFile & ".xlsx"
	End If
	oFile = ex.GetSaveAsFilename(partFile,"Excel Workbook (*.xlsx), *.xlsx")
	If oFile = "False" Then
		MsgBox("An output file must be selected. The script is now terminating.")
		WScript.ConnectObject session,     "off"
		WScript.ConnectObject application, "off"
		WScript.Quit
	End If
	wb.SaveAs oFile, 51
	
	Set ws = wb.Sheets(wb.ActiveSheet.Name)

	'wb.SaveAs "O:\CustSvc\Parts\SRVCTRS\Ops\allcost\Costing Sheet Repository\"&partFile,51

	vrc = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").VisibleRowCount
	row = 0
	sapRow = 0
	ws.Cells(1,1).Value = "'"&part
	ex.Cells(1,1).Font.Size = 14
	ws.Cells(1,5).Value = "Created: " & Now()
	ws.Cells(2,1).Value = "Item"
	ws.Cells(2,2).Value = "Qty"
	ws.Cells(2,3).Value = "UoM"
	ws.Cells(2,4).Value = "Material"
	ws.Cells(2,5).Value = "Description"
	ws.Cells(2,6).Value = "Cost"
	ws.Cells(2,7).Value = "Total Cost"
	ws.Cells(2,8).Value = "Stock"
	ws.Cells(2,9).Value = "Source"
	ws.Cells(2,10).Value = "Lead Days"
	ex.Range("A2:J2").Font.Bold = True
	ex.Range("A2:J2").HorizontalAlignment = -4108
	While session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,"&sapRow&"]").text <> "__________________"

		If sapRow + 1 = vrc Then
			ws.Cells(row+3,1).Value = row + 1
			ws.Cells(row+3,2).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,"&sapRow&"]").text
			ws.Cells(row+3,3).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-MEINS[5,"&sapRow&"]").text
			ws.Cells(row+3,4).Value = "'"&session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,"&sapRow&"]").text
			ws.Cells(row+3,5).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-KTEXT[3,"&sapRow&"]").text
			ws.Cells(row+3,7).Formula = "=B"&row+3&"*F"&row+3
			
			session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").verticalScrollbar.position = row + 1
			sapRow = - 1
			If ws.Cells(row+3,4).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,"&sapRow + 1&"]").text Then
				row = row - 1
			End If
		Else

			ws.Cells(row+3,1).Value = row + 1
			ws.Cells(row+3,2).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-MENGE[4,"&sapRow&"]").text
			ws.Cells(row+3,3).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-MEINS[5,"&sapRow&"]").text
			ws.Cells(row+3,4).Value = "'"&session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,"&sapRow&"]").text
			ws.Cells(row+3,5).Value = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/txtRC29P-KTEXT[3,"&sapRow&"]").text
			ws.Cells(row+3,7).Formula = "=B"&row+3&"*F"&row+3
			
		End If
		
		row = row + 1
		sapRow = sapRow + 1
	Wend
	ws.Cells(row+3,1).Value = row + 1
	ws.Cells(row+3,3).Value = "Hr"
	ws.Cells(row+3,4).Value = "LABOR"
	ws.Cells(row+3,6).Value = "46"
	ws.Cells(row+3,7).Formula = "=B"&row+3&"*F"&row+3
	
	ws.Cells(row+4,6).Value = "Total:"
	ws.Cells(row+5,6).Value = "5%"
	ws.Cells(row+6,6).Value = "Lot Total:"
	ws.Cells(row+4,7).Formula = "=Sum(G3:G"&row+3&")"
	ws.Cells(row+5,7).Formula = "=((Sum(G3:G"&row+2&"))*.05+G"&row+4&")"
	ws.Cells(row+6,7).Formula = "=G"&row+5&"*"&lotSize
	
	ex.Range("G"&row+5).select
	With ex.Selection.Borders(9)
        .LineStyle = 1
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = 2
	End With
	
	Call costOutBom
	
End Sub

Sub costOutBom
	workingRow = 3
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	While ws.Cells(workingRow,4).Value <> "LABOR"
		
		Call GetComponentCosts
		workingRow = workingRow + 1
		
	Wend
	
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = part
	session.findById("wnd[0]/tbar[0]/btn[0]").press

	If Right(session.findById("wnd[1]").text,7) = "View(s)" Then
		session.findById("wnd[1]/tbar[0]/btn[20]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = plant
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
		Session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		
	Else 	
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = plant
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
		Session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		
	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "V" Then
		ws.Cells(1,3).Value = "Variable??? " & session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text
	Else
		ws.Cells(1,3).Value =  "Standard: " & session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text
	End If
	ws.Cells(1,4).Value = "On Hand: " & session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-LBKUM").text

	session.findById("wnd[0]/tbar[0]/btn[3]").press	
	
	
	ws.Cells(workingRow+1,10).Value = "Longest:" 
	ws.Cells(workingRow+2,10).Formula = "=Max(J3:J"&workingRow-1&")"
	maxLead = ws.Cells(workingRow+2,10).Value
	ws.Cells(workingRow+2,10).Value = maxLead & " Days"
	Call GetVendors
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nCA03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/ctxtRC27M-MATNR").text = part
	session.findById("wnd[0]/usr/ctxtRC27M-WERKS").text = plant
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	If Right(session.findById("wnd[0]").text,14) = "Initial Screen" Then
		session.findById("wnd[1]").close
		ws.Cells(workingRow,5).Value = "NO ROUTING EXISTS!"
	ElseIf Right(session.findById("wnd[0]").text,14) <> "Initial Screen" And plant = "500E" Then
		ws.Cells(workingRow, 1).Value = ""
		ws.Cells(workingRow, 3).Value = ""
		ws.Cells(workingRow, 4).Value = ""
		ws.Cells(workingRow, 5).Value = "SEE BELOW FOR LABOR BREAKDOWN"
		ex.Range("E"&workingRow).Font.Bold = True
		ws.Cells(workingRow, 6).Value = ""
		ws.Cells(workingRow, 7).Value = ""
		ws.Cells(workingRow + 3, 1).Value = "LABOR:"
		ws.Cells(workingRow + 4, 1).Value = "LOT SIZE OF"
		ws.Cells(workingRow + 4, 2).Value = lotSize
		ws.Cells(workingRow + 5, 6).Value = "WkCtr Totals"
		ws.Cells(workingRow + 5, 1).Value = "Work Ctr"
		ws.Cells(workingRow + 5, 2).Value = "Setup"
		ws.Cells(workingRow + 5, 3).Value = "Hours"
		ws.Cells(workingRow + 5, 4).Value = "Rate"
		ws.Cells(workingRow + 5, 5).Value = "Description"
		ex.Range("A"&workingRow+3&":F"&workingRow+5).Font.Bold = True
		
		ws.Range("A"&workingRow+5&":F"& workingRow + (session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text + 5)).Select
		ex.Selection.Borders(5).LineStyle = -4142
	    ex.Selection.Borders(6).LineStyle = -4142
	    With ex.Selection.Borders(7)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
	    With ex.Selection.Borders(8)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
	    With ex.Selection.Borders(9)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
	    With ex.Selection.Borders(10)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
	    With ex.Selection.Borders(11)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
	    With ex.Selection.Borders(12)
	        .LineStyle = 1
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 2
	    End With
		
		For i = 6 To session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text + 5
			wrkCtr = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/ctxtPLPOD-ARBPL[2,"&i-6&"]").text
			Select Case wrkCtr
				Case "150MC"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "95"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "202BM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "95"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
							
				Case "203BM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "150"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "204GL"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "95"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "205CM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "95"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "401DP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
							
				Case "402DP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
							
				Case "Blast"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "CNCB"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "95"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "CNCP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
										
				Case "ECOST"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "50"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "FINSP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "85"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "FORM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "90"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
						
				Case "FSA"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "100"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "HWELD"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "L192"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "90"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LASSY"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LCNC"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "100"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LGVBM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "150"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LGWP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "70"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LO"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "LWELD"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "MASSY"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "MINSP"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "85"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "MMACH"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "NID/MIXER"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "60"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "NCVBM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "140"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "ODGR"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "85"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "Paint"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "PWELD"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "REBUILD"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "46"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "RIFF"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "60"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "SAW"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "65"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "SHRC"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "50"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "STEAM"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "75"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "TMD16"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "135"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "VCNC"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "130"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "VTL36"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "80"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
					
				Case "VTL54"
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 2).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i-6&"]").text
					ws.Cells(workingRow + i, 3).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i-6&"]").text
					ws.Cells(workingRow + i, 4).Value = "85"
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = (ws.Cells(workingRow + i, 2).Value * ws.Cells(workingRow + i, 4).Value)+(lotSize * ws.Cells(workingRow + i, 3).Value * ws.Cells(workingRow + i, 4).Value)
				
				Case Else
					ws.Cells(workingRow + i, 1).Value = wrkCtr
					ws.Cells(workingRow + i, 5).Value = session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-LTXA1[6,"&i-6&"]").text
					ws.Cells(workingRow + i, 6).Value = "NEEDS REVIEW"
			End Select
		Next
		ws.Cells(workingRow+i+1,5).Value = "Total Labor Cost For Lot of: "&lotSize
		ws.Cells(workingRow+i+2,5).Value = "TOTAL COST"
		ws.Cells(workingRow+i+2,5).Font.Bold = True
		ws.Cells(workingRow+i+1,6).Formula = "=sum(F"&workingRow+6&":F"&workingRow+i-1&")"
		ws.Cells(workingRow+i+1,7).Formula = "=F"&workingRow+i+1
		ex.Range("G"&workingRow+i+1).select
		With ex.Selection.Borders(9)
	        .LineStyle = -4119
	        .ColorIndex = 0
	        .TintAndShade = 0
	        .Weight = 4
    	End With
		ws.Cells(workingRow+i+2,7).Formula = "=sum(G"&workingRow+3&":G"&workingRow+i+1&")"
		
		For i = 6 To session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text + 5
			If InStr(ws.Cells(workingRow + i, 5).Value,"PSTR") <> 0 Then
				ws.Rows(workingRow + i&":"&workingRow + i).Select
    			ex.Selection.Delete -4162
    		End If
		Next		
	Else
		setupHours = 0
		laborTotal = 0
		For i = 0 To session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text - 1
			If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[13,"&i&"]").text <> "" Then
			setupHours = setupHours + session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i&"]").text
			End If
			
			If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[16,"&i&"]").text <> "" Then
			laborTotal = laborTotal + session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i&"]").text
			End If 
		Next
	
		If setupTotal > "0.0" Then
			ws.Cells(row+6,5).Value = "NOTE: Assembly requires "&setupHours&" hours of setup."
			ws.Cells(row+7,5).Value = "$"&setupHours*46&" needs to be added to the cost of the lot."
		End If
		ws.Cells(row+3,2).Value = laborTotal
		ws.Cells(row+6,5).Value = "TOTAL LOT SIZE OF "&lotSize&" ASSEMBLIES"
		ex.Cells(row+6,6).Font.Bold = True
	End If	
	
	ws.Range("F3:G"&workingRow).Select
	ex.Selection.Style = "Currency"
	ex.Cells(workingRow+1,7).Style = "Currency"
	ex.Cells(workingRow+2,7).Style = "Currency"
	ex.Cells(workingRow+3,7).Style = "Currency"
	ws.Range("D"&workingRow+6&":G"&(i + workingRow)+2).Select
	ex.Selection.Style = "Currency"

	With ws.PageSetup
		.Orientation = 2
		.Zoom = False
		.FitToPagesWide = 1
		.FitToPagesTall = False
		.PaperSize = 1
		.PrintArea = "A:K"
	End With
	ws.Columns.AutoFit
	ws.Cells(1,1).Select
	wb.ActiveSheet.name = ws.cells(1,1).value
	wb.Close(True)
	ex.Quit 
	
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
End Sub

Sub GetComponentCosts
	
	If Len(ws.Cells(workingRow,4).Value) < 1 Then
		Exit sub
	End If
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ws.Cells(workingRow,4).Value
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	If session.findById("wnd[0]/sbar").text <> "" Then
		ws.Cells(workingRow,9).Value = session.findById("wnd[0]/sbar").text
		Exit Sub
	End If
	
	If Right(session.findById("wnd[1]").text,7) = "View(s)" Then
		session.findById("wnd[1]/tbar[0]/btn[20]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = plant
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
		Session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		
	Else 	
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = plant
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
		Session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		
	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "F" Then
		ws.Cells(workingRow,9).Value = "BUY"
	ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "E" Then
		ws.Cells(workingRow,9).Value = "MAKE"
	ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "X" Then
		ws.Cells(workingRow,9).Value = "Make or Buy"
	Else
		ws.Cells(workingRow,9).Value = "UNKNOWN"
	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text <> 0 Then
		ws.Cells(workingRow,10).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text
	Else
		ws.Cells(workingRow,10).Value = "UNKNOWN"
	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "V" Then
		ws.Cells(workingRow,6).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text
	Else
		ws.Cells(workingRow,6).Value =  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text
	End If
	ws.Cells(workingRow,8).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-LBKUM").text
	session.findById("wnd[0]/tbar[0]/btn[3]").press	

End Sub

Sub GetVendors
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nME03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/ctxtEORD-WERKS").text = plant
	For i = 3 To workingRow - 1
		If ws.Cells(i,9).Value = "BUY" Then
			session.findById("wnd[0]/usr/ctxtEORD-MATNR").text = ws.Cells(i,4).Value
			session.findById("wnd[0]/tbar[0]/btn[0]").press
			If session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text <> "" Then
				ws.Cells(i,9).Value = "'"&session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text
			Else
				ws.Cells(i,9).Value = "No Source List Record"
			End If 
			session.findById("wnd[0]/tbar[0]/btn[3]").press
		End If
	Next
End Sub
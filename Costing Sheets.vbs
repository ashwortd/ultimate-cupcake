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
Dim part,plant,status,continue,result
Dim setupTotal,laborTotal,maxLead

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
			plant = InputBox("Enter plant below:", "Plant")
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
	
		session.findById("wnd[0]/usr/ctxtRC29N-STLAN").text = "1"
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		status = session.findById("wnd[0]/sbar").Text
		
		If status <> "" Then
			MsgBox("ERROR:"&Chr(13)&status&Chr(13)&Chr(13)&"Script will start over.")
			continue = "No"
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
	wb.SaveAs "D:\Documents and Settings\dma02\Desktop\"&partFile,51
	
	Set ws = wb.Sheets(wb.ActiveSheet.Name)

	vrc = session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT").VisibleRowCount
	row = 0
	sapRow = 0
	ws.Cells(1,1).Value = "'"&part
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
	While session.findById("wnd[0]/usr/tabsTS_ITOV/tabpTCMA/ssubSUBPAGE:SAPLCSDI:0152/tblSAPLCSDITCMAT/ctxtRC29P-IDNRK[2,"&sapRow&"]").text <> "__________________"
	MsgBox("sapRow="&sapRow&" row="&row&" vrc="&vrc)
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
		ws.Cells(row+5,6).Value = "3%"
		ws.Cells(row+4,7).Formula = "=Sum(G3:G"&row+3&")"
		ws.Cells(row+5,7).Formula = "=G"&row+4&"*1.03"
		
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
	If session.findById("wnd[0]/titl").text = "Display Routing: Initial Screen" Then
		session.findById("wnd[1]").close
		ws.Cells(workingRow,5).Value = "NO ROUTING EXISTS!"
	Else
		setupTotal = 0
		laborTotal = 0
		For i = 0 To session.findById("wnd[0]/usr/txtRC27X-ENTRIES").text - 1
			If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i&"]").text <> "" Then
			setupTotal = setupTotal + session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW01[11,"&i&"]").text
			End If
			
			If session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i&"]").text <> "" Then
			laborTotal = laborTotal + session.findById("wnd[0]/usr/tblSAPLCPDITCTRL_1400/txtPLPOD-VGW02[14,"&i&"]").text
			End If 
		Next
	
		If setupTotal > "0.0" Then
			ws.Cells(row+6,5).Value = "NOTE: Assembly requires "&setupTotal&" hours of setup."
			ws.Cells(row+7,5).Value = "$"&setupTotal*46&" needs to be added to the cost of the lot."
		End If
		ws.Cells(row+3,2).Value = laborTotal
		
	End If	
	
	ws.Columns.AutoFit
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
	
	If session.findById("wnd[1]").text = "Select View(s)" Then
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
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press
	Else 	
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = plant
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
		Session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
		If Err.Number <> 0 Then
			Err.Clear
		End If
		On Error Goto 0
		session.findById("wnd[1]/tbar[0]/btn[0]").press
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
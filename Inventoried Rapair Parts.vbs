'******************************************
'******************************************
'** Author: Adam Damke		 
'** 									 
'** Description: This is where you would
'** put the description of what your
'** script does.
'** 
'** 
'** 
'** 									 
'** File: Inventoried Repair Parts.vbs
'** Date Last Updated: 12/05/2013
'*********************************************************
'*********************************************************

On Error Resume Next
If Not IsObject(application) Then
	Set SapGuiAuto = GetObject("SAPGUI")
	Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
	Set connection = application.Children(0)
	If Err.Number <> 0 Then
		MsgBox("You are not properly logged into SAP."& chr(13) &"Please login and try again."& chr(13) & chr(13) &"Script terminating...")
		WScript.Quit
	End If
On Error Goto 0
End If
If Not IsObject(session) Then
	Set session = connection.Children(0)
End If
If IsObject(WScript) Then
	WScript.ConnectObject session, "on"
	WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

Dim objShell, strPath
Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory


Dim fso, ex, wb, ws
Dim status, lastRow, sLoc, row
Const xlUp = -4162

Call Main



Sub Main

	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZWMLIST"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "U03"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELEK/tabpPUSH1/ssub%_SUBSCREEN_SELEK:ZWM_OVERVIEW:0100/btn%_S_MATNR_%_APP_%-VALU_PUSH").press
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "*REB*"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "*RBLD*"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "*RBLT*"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "*REBUILT*"
	session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "*REP*"
	session.findById("wnd[1]/tbar[0]/btn[8]").press
	
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = strPath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SESC Repair Inventory.txt"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").text
	If InStr(status, "transmitted") = 0 Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
'	--------------------------------------------------------------------------------------------------------
	
	session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "U04"
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = strPath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "SWSC Repair Inventory.txt"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").text
	If InStr(status, "transmitted") = 0 Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
'	--------------------------------------------------------------------------------------------------------
	
	session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "U05"
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = strPath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "MWSC Repair Inventory.txt"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").text
	If InStr(status, "transmitted") = 0 Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
'	--------------------------------------------------------------------------------------------------------
	
	session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "U06"
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = strPath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "RMSC Repair Inventory.txt"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").text
	If InStr(status, "transmitted") = 0 Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	
'	--------------------------------------------------------------------------------------------------------

	session.findById("wnd[0]/usr/ctxtP_LGNUM").text = "U07"
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
	session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/ctxtDY_PATH").text = strPath
	session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "NESC Repair Inventory.txt"
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	WScript.Sleep(1000)
	status = session.findById("wnd[0]/sbar").text
	If InStr(status, "transmitted") = 0 Then
		session.findById("wnd[1]/tbar[0]/btn[11]").press
		WScript.Sleep(500)
	End If	
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	Call CreateReport
	
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End Sub



Sub CreateReport

	Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
	Set ex = WScript.CreateObject("Excel.application")
	ex.Visible = True
	
	Set wb = ex.Workbooks.Open(strPath & "\500X Inventoried Repair Parts.xlsx")
	
'	Set ws = wb.Sheets(1)
	For Each sheet In wb.Sheets
		sheet.Cells.ClearContents
		Set wkbk = ex.Workbooks.Open(strPath & "\" & sheet.name & " Repair Inventory.txt")
		wkbk.Sheets(1).Cells.Copy
		'ex.Workbooks("500X Inventoried Repair Parts.xlsx").Activate
		wb.Activate
		sheet.select
		sheet.Cells(1,1).select
		sheet.paste
		ex.CutCopyMode = False
		sheet.Cells(1,1).select
		wkbk.Close(True)
		Set wkbk = Nothing	
		
	Next

		
	For Each sheet In wb.Sheets
		sheet.select
		
		If Left(sheet.name,1) = "M" Or Left(sheet.name,1) = "N" Then
			sheet.Columns("Q:R").Select
			ex.Selection.Delete -4131
			
			sheet.Columns("K:N").Select
			ex.Selection.Delete -4131
			
			sheet.Columns("H:H").Select
			ex.Selection.Delete -4131
			
			sheet.Columns("A:C").Select
			ex.Selection.Delete -4131
		Else
			sheet.Columns("Q:R").Select
			ex.Selection.Delete -4131
			
			sheet.Columns("J:M").Select
			ex.Selection.Delete -4131
			
			sheet.Columns("A:C").Select
			ex.Selection.Delete -4131
		End If
		
		sheet.Rows("6:6").Select
	    ex.Selection.Delete -4162
	    
	    sheet.Rows("1:4").Select
	    ex.Selection.Delete -4162
		    
	
		sheet.Cells(1,9).Value = "SLoc Description"
		
		lastRow = sheet.range("C" & sheet.Rows.Count).End(xlUp).Row
		
		For Each cell In sheet.Range("F2:F" & lastRow)
		
			If cell.Value = "1" Then
				cell.Value = "'0001"
			End If
		
		Next
		
		For Each cell In sheet.Range("A1:I1")
		
			cell.Font.Size = 12
			cell.Font.Bold = True
		Next
		
		For Each cell In sheet.Range("A2:A" & lastRow)
		
			cell.HorizontalAlignment = -4152
			
		Next
		
	'	row = 2
	'	While row <= lastRow
		
	'		sLoc = sheet.Cells(row,6).Value 
			
	'		Select Case sLoc
	'		Case "D001"
	'			sheet.Cells(row,9).Value = "AES-SOMERSET"
	'		Case "D002"
	'			sheet.Cells(row,9).Value = "ALL ALBRIGHT-AAS"
	'		Case "D003"
	'			sheet.Cells(row,9).Value = "ALL FTMARTIN-AFM"
	'		Case "D004"
	'			sheet.Cells(row,9).Value = "ALL MITCHELL-AMS"
	'		Case "D005"
	'			sheet.Cells(row,9).Value = "DET RIVERRGE-DRR"
	'		Case "D006"
	'			sheet.Cells(row,9).Value = "DETROIT ED-DTC"
	'		Case "D007"
	'			sheet.Cells(row,9).Value = "DETROIT ED-DET"
	'		Case "D008"
	'			sheet.Cells(row,9).Value = "DOM CHESAPK-DCP"
	'		Case "D009"
	'			sheet.Cells(row,9).Value = "DOM CLOVER-DCL"
	'		Case "D010"
	'			sheet.Cells(row,9).Value = "DOM ENERGY-NEP"
	'		Case "D011"
	'			sheet.Cells(row,9).Value = "DOM MT.STORM-DMS"
	'		Case "D012"
	'			sheet.Cells(row,9).Value = "DOM YORKTOWN-DYO"
	'		Case "D013"
	'			sheet.Cells(row,9).Value = "DOMCHESTRFLD-DCF"
	'		Case "D014"
	'			sheet.Cells(row,9).Value = "DYDANSKAMMER-DYD"
	'		Case "D015"
	'			sheet.Cells(row,9).Value = "EME HOMERCTY-HMR"
	'		Case "D016"
	'			sheet.Cells(row,9).Value = "EXELON-PEC"
	'		Case "D017"
	'			sheet.Cells(row,9).Value = "FE ASHTABUL-FEA"
	'		Case "D018"
	'			sheet.Cells(row,9).Value = "FE E LAKE-CSC"
	'		Case "D019"
	'			sheet.Cells(row,9).Value = "FE SAMMIS-OES"
	'		Case "D020"
	'			sheet.Cells(row,9).Value = "HORSEHEAD-HHC"
	'		Case "D021"
	'			sheet.Cells(row,9).Value = "NRG DUNKIRK-NMP"
	'		Case "D022"
	'			sheet.Cells(row,9).Value = "NRG HUNTLEY-NEH"
	'		Case "D023"
	'			sheet.Cells(row,9).Value = "NRG SOMERSET-NES"
	'		Case "D024"
	'			sheet.Cells(row,9).Value = "PPL GEN TUBE-PP2"
	'		Case "D025"
	'			sheet.Cells(row,9).Value = "PPL BRUNNER-PGB"
	'		Case "D026"
	'			sheet.Cells(row,9).Value = "PPL MONTOUR-PGM"
	'		Case "D027"
	'			sheet.Cells(row,9).Value = "PSEG FOSSIL-UIL"
	'		Case "D028"
	'			sheet.Cells(row,9).Value = "PUBLC SVC NH-PNH"
	'		Case "D029"
	'			sheet.Cells(row,9).Value = "RE CONEMAUG-REC"
	'		Case "D030"
	'			sheet.Cells(row,9).Value = "RE KEYSTN-REK"
	'		Case "D031"
	'			sheet.Cells(row,9).Value = "RE PORTLAND-REP"
	'		Case "D032"
	'			sheet.Cells(row,9).Value = "RE SEWRDTUBE-SEW"
	'		Case "D033"
	'			sheet.Cells(row,9).Value = "RE SHAWVILLE-RES"
	'		Case "D034"
	'			sheet.Cells(row,9).Value = "RE TITUS-RET"
	'		Case "D035"
	'			sheet.Cells(row,9).Value = "RELIANT NE-PEL"
	'		Case "D036"
	'			sheet.Cells(row,9).Value = "RPAULSMITH-ARP"
	'		Case "D039"
	'			sheet.Cells(row,9).Value = "TRIGEN CINGY-TCS"
	'		Case "D040"
	'			sheet.Cells(row,9).Value = "ALL HATFIELD-AHF"
	'		Case "D041"
	'			sheet.Cells(row,9).Value = "GENON PWR MW-GCH"
	'		Case "D042"
	'			sheet.Cells(row,9).Value = "GENON DICKERSON"
	'		Case "D043"
	'			sheet.Cells(row,9).Value = "RE SEWARD-REW"
	'		Case "DE01"
	'			sheet.Cells(row,9).Value = "DETROIT DET@MW"
	'		Case "DE02"
	'			sheet.Cells(row,9).Value = "FE ASHTAB-FEA@MW"
	'		Case "DE03"
	'			sheet.Cells(row,9).Value = "FE ELAKE-CS3@MW"
	'		Case "DE04"
	'			sheet.Cells(row,9).Value = "FE SAMMIS-OE3@MW"
	'		Case "DE05"
	'			sheet.Cells(row,9).Value = "PPL GEN PGB@CTW"
	'		Case "DE06"
	'			sheet.Cells(row,9).Value = "PUBLC SVC-PNH@MW"
	'		Case "DE07"
	'			sheet.Cells(row,9).Value = "RE KEYST-REK@CTW"
	'		Case "DE08"
	'			sheet.Cells(row,9).Value = "RE SHAWVIL-RES@M"
	'		Case "DE09"
	'			sheet.Cells(row,9).Value = "SDW @ ALLEN"
	'		Case "DS01"
	'			sheet.Cells(row,9).Value = "SDW SHARED"
	'		Case "DZ01"
	'			sheet.Cells(row,9).Value = "NESC OPT STK-NEG"
	'		Case "0001"
	'			sheet.Cells(row,9).Value = "NESC SUPPLY AREA"
	'		Case "GD01"   
	'			sheet.Cells(row,9).Value = "NE DEPLETION STK"
	'		Case Else
	'			sheet.Cells(row,9).Value = "SLoc Description not known by Script"
	'		End Select
			
	'		row = row + 1
	'	Wend
		
		row = 2
		While row <= lastRow
			If sheet.Cells(row,3).Value = "RECV-RBLD" Then
				sheet.Rows(row & ":" & row).Select
	    		ex.Selection.Delete -4162
				row = row - 1
				lastRow = lastRow - 1
			End If
			row = row + 1
		Wend
		
		lastRow = sheet.range("C" & sheet.Rows.Count).End(xlUp).Row
		sheet.Range("A2:I" & lastRow).Sort sheet.Range("C2:C" & lastRow), 1
		sheet.Cells(1,1).select
		sheet.Columns.AutoFit
		With sheet.PageSetup
			.Orientation = 1
			.Zoom = False
			.FitToPagesWide = 1
			.FitToPagesTall = False
			.PaperSize = 1
			.PrintArea = "A:I"
		End With
	
	Next
	
	MsgBox("Done.")
	
	wb.Close(True)
	ex.Quit
	Set fso = Nothing
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing

End Sub
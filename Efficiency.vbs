   On Error Resume Next
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
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
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize

Dim ex,wb,ws,lastRow
Dim vrc,rc,n,i
Const xlUp = -4162
Const xlColorIndexNone = -4142

Call Main


Sub Main

	Set ex = GetObject( , "Excel.Application")
	Set wb = ex.Workbooks("Copy of Efficiency Report.xlsm")
	
	Set ws = wb.Sheets("REPORT")
	ws.rows("2:" & ws.rows.count).Interior.ColorIndex = xlColorIndexNone
	ws.rows("2:" & ws.rows.count).clearcontents
	Set ws = wb.Sheets("Sheet2")
	ws.Cells(1,1).CurrentRegion.Offset(1,0).clearcontents
	Set ws = wb.Sheets("Sheet3")
	ws.Cells(1,1).CurrentRegion.Offset(1,0).clearcontents
	Set ws = wb.Sheets("SETUP")

	Call ZPP_006
	Call ZPP_TIMEBOOKING
	Call ClearRawData
	Call Cleanup

	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
	
	MsgBox("The requested report has been created." & chr(13) & chr(13) & "Thank you.")					
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit	

End Sub



Sub ZPP_006
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_006"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/ctxtZSTARTD").text = ws.Cells(7,6).Value 
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/ctxtZSTARTT").text = "00:00:01"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/ctxtZENDD").text = ws.Cells(10,6).Value
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/ctxtZENDT").text = "23:59:59"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/radC_ALLCON").select
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_CONF:1200/ctxtS_WERKS-LOW").text = ws.Cells(4,6).Value

	session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	Call ZPP_006_LAYOUT
	
	Call CopyPasteReport("Sheet2")
	Call FixUpSheet2
	ws.Cells(1,1).select
End Sub



Sub ZPP_TIMEBOOKING
	
	ws.Range("D2:D"&lastRow).copy
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nZPP_TIMEBOOKING"
	session.findById("wnd[0]/tbar[0]/btn[0]").press

	session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:ZPPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/PP_500E_004"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:ZPPIO_ENTRY:1200/btn%_S_AUFNR_%_APP_%-VALU_PUSH").press	
	session.findById("wnd[1]/tbar[0]/btn[24]").press
	session.findById("wnd[1]/tbar[0]/btn[8]").press


	session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	Call CopyPasteReport("Sheet3")
	Call FixUpSheet3
	
	ws.Cells(2,2).CurrentRegion.Offset(1,0).Select
	
	ex.Selection.copy
	ws.Cells(1,1).select
	Set ws = wb.Sheets("REPORT")
	ws.select
	ws.Range("A2").select
	ws.paste
	ex.CutCopyMode = False
	
	session.findById("wnd[0]/tbar[0]/btn[15]").press
	session.findById("wnd[0]/tbar[0]/btn[15]").press

End Sub



Sub ZPP_006_LAYOUT
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_VARIANT"
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&COL0"
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER2_LAYO/shellcont/shell").selectAll
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_FL_SING").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressToolbarButton "&FIND"
	session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "Order"
	session.findById("wnd[2]/tbar[0]/btn[0]").press
	session.findById("wnd[2]/tbar[0]/btn[12]").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/cntlCONTAINER1_LAYO/shellcont/shell").pressToolbarButton "&FIND"
	session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "Operation/activity"
	session.findById("wnd[2]/tbar[0]/btn[0]").press
	session.findById("wnd[2]/tbar[0]/btn[12]").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2").select
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R2/ssubSUB_DYN0510:SAPLSKBH:0610/btnAPP_FL_SING").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R3").select
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R3/ssubSUB_DYN0510:SAPLSKBH:0600/cntlCONTAINER1_FILT/shellcont/shell").pressToolbarButton "&FIND"
	session.findById("wnd[2]/usr/txtGS_SEARCH-VALUE").text = "order type"
	session.findById("wnd[2]/tbar[0]/btn[0]").press
	session.findById("wnd[2]/tbar[0]/btn[12]").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R3/ssubSUB_DYN0510:SAPLSKBH:0600/btnAPP_WL_SING").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R3/ssubSUB_DYN0510:SAPLSKBH:0600/cntlCONTAINER2_FILT/shellcont/shell").selectedRows = "0"
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R3/ssubSUB_DYN0510:SAPLSKBH:0600/btn600_BUTTON").press
	session.findById("wnd[2]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/btn%_%%DYN001_%_APP_%-VALU_PUSH").press
	session.findById("wnd[3]/usr/tabsTAB_STRIP/tabpNOSV").select
	session.findById("wnd[3]/usr/tabsTAB_STRIP/tabpNOSV/ssubSCREEN_HEADER:SAPLALDB:3030/tblSAPLALDBSINGLE_E/ctxtRSCSEL_255-SLOW_E[1,0]").text = "ZNPH"
	session.findById("wnd[3]/tbar[0]/btn[8]").press
	session.findById("wnd[2]/tbar[0]/btn[0]").press
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R5").select
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R5/ssubSUB_DYN0510:SAPLSKBH:0503/chkGS52_SCREEN-LAYOUT-CWIDTH_OPT").selected = false
	session.findById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R5/ssubSUB_DYN0510:SAPLSKBH:0503/chkGS52_SCREEN-LAYOUT-CWIDTH_OPT").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
End Sub



Sub CopyPasteReport(reportName)
	
	vrc=session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").visiblerowcount
	rc=session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").rowcount
	n = Round(rc/vrc,0)+1
	
	For i = 1 To n
		On Error Resume Next
		session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = vrc * i
		If Err.Number <> 0 Then
			Err.Clear
		End If
		WScript.Sleep(250)
	Next
	On Error Goto 0
	
	
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectAll
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItemByPosition "0"
	
	Set ws = wb.Sheets(reportName)
	ws.select
	ws.Cells(1,1).CurrentRegion.Offset(1,0).clearcontents
	ws.Range("A2").select
	ws.PasteSpecial , , , , , ,true
	ex.CutCopyMode = False
	
End Sub



Sub FixUpSheet2

	lastRow = ws.range("A" & ws.Rows.Count).End(xlUp).Row
	ws.Range("$A$1:$B$"&lastRow).RemoveDuplicates Array(1, 2), 1
	ws.Cells(2,3).Formula = "=A2&"&Chr(34)&"/"&Chr(34)&"&B2"
	ws.Range("C2").autofill ws.Range("C2:C"&lastRow)
	ws.Range("A1").entireColumn.copy
	ws.Range("D1").select
	ws.paste
	ex.CutCopyMode = False
	lastRow = ws.range("D" & ws.Rows.Count).End(xlUp).Row
	ws.Range("$D$1:$D$"&lastRow).RemoveDuplicates 1, 1
	
End Sub



Sub FixUpSheet3
	
	lastRow = ws.range("A" & ws.Rows.Count).End(xlUp).Row
	ws.Cells(2,13).Formula = "=I2+K2"
	ws.Cells(2,14).Formula = "=H2+J2"
	ws.Cells(2,15).Formula = "=A2&"&Chr(34)&"/"&Chr(34)&"&F2"
	ws.Cells(2,16).Formula = "=MATCH(O2,Sheet2!C:C,0)"
	ws.Range("M2:P2").autofill ws.Range("M2:P"&lastRow)
	ws.Range("M2:N" & lastRow).select
	ex.Selection.Copy
	ws.Range("M2:N" & lastRow).PasteSpecial -4163
	ex.CutCopyMode = False
	ws.Cells(1,1).CurrentRegion.Select
	ex.Selection.AutoFilter 16,"=#N/A"
	ws.Cells(1,1).CurrentRegion.Offset(1,0).Select
	ex.Selection.Delete -4162
	ex.Selection.AutoFilter
	ws.Columns("O:P").Delete -4131
	ws.Columns("H:K").Delete -4131
	
	
End Sub



Sub ClearRawData

	Set ws = wb.Sheets("Sheet2")
	ws.Cells(1,1).CurrentRegion.Offset(1,0).clearcontents
	Set ws = wb.Sheets("Sheet3")
	ws.Cells(1,1).CurrentRegion.Offset(1,0).clearcontents
		
End Sub



Sub Cleanup

ex.Application.Run "'Efficiency Report.xlsm'!Cleanup"
Set ws = wb.Sheets("Report")
ws.select
ws.Columns("H").Hidden = True
ws.Columns("K").Style = "Percent"
ws.Cells(1,12).select
End Sub
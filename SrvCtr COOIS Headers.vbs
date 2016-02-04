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

Dim objShell
Set objShell = CreateObject("Wscript.Shell")
Dim ex, wb, ws, strPath
Dim vrc, rc, n, i, report
Dim fName, plant, profile

strPath = objShell.CurrentDirectory

Call Main

Sub Main
	report = 1
	While report <= 6
		
		Select Case report
		
		Case 1
			fName = "500C_Orders.xlsx"
			plant = "500C"
		
		Case 2
			fName = "500D_Orders.xlsx"
			plant = "500D"
		
		Case 3
			fName = "500F_Orders.xlsx"
			plant = "500F"
		
		Case 4
			fName = "500G_Orders.xlsx"
			plant = "500G"
		
		Case 5
			fName = "500H_Orders.xlsx"
			plant = "500H"
		
		Case 6
			fName = "500I_Orders.xlsx"
			plant = "500I"
		
		End Select
		
		Call PrepReport
	
		Call RunReport
	
		Call CloseReport
		
		report = report + 1
	
	Wend
	
	Call SendEmail
	Set objShell = Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End Sub


Sub PrepReport
	Set ex = WScript.CreateObject("Excel.Application")
	ex.Visible = True

	Set wb = ex.Workbooks.Open(strPath&"\"&fName)
	
End Sub


	
Sub RunReport

	i = 1
	While i <= 3
		
		Select Case i
		
		Case 1
			profile = "ZPP0040"
		
		Case 2
			profile = "ZPP0060"
		
		Case 3
			profile = "ZPP0070"
		
		End Select
	
		Call COOIS_ORDER_HEADERS
		
		i = i + 1
	
	Wend
	
	

End Sub


	
Sub CloseReport
	MsgBox(wb.Name)
	wb.Close(True)
	ex.Quit
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
End Sub




Sub COOIS_ORDER_HEADERS

	session.findById("wnd[0]/tbar[0]/okcd").text = "/nCOOIS"
	session.findById("wnd[0]/tbar[0]/btn[0]").press

	session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/cmbPPIO_ENTRY_SC1100-PPIO_LISTTYP").key = "PPIOH000"
	
	session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").text = "/PP_0000_001"
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_WERKS-LOW").text = plant
	session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtP_SELID").text = profile

	session.findById("wnd[0]/tbar[1]/btn[8]").press
	
	Set ws = wb.Sheets(wb.Sheets(i).Name)
	ws.select
	ws.Cells(1,1).CurrentRegion.Offset(2,0).clearcontents
	If session.findById("wnd[0]/titl").text = "SAP" Then
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		
		ws.Cells(4,2).Value = "NO DATA FOR SYSTEM STATUS PROFILE: " & profile
		ws.Cells(1,2).Value = "Created: " & Now()
		Exit Sub
	End If

	Call CopyPasteReport

End Sub


Sub SendEmail

	Dim objOutl,objMailItem,strEmailAddr
	
	Set objOutl = CreateObject("Outlook.Application")
	
	Set objMailItem = objOutl.CreateItem(0)
	'comment the next line if you do not want to see the outlook window
	objMailItem.Display
	strEmailAddr  = "adam.damke@power.alstom.com"
	objMailItem.Recipients.Add strEmailAddr
	objMailItem.Subject = "Testing Automatic Email Sending w/ Attachments"
	objMailItem.Body = "See Attached"
	objMailItem.Attachments.Add "O:\CustSvc\Parts\SRVCTRS\Ops\Service Center Orders\500C_Orders.xlsx"
	objMailItem.Attachments.Add "O:\CustSvc\Parts\SRVCTRS\Ops\Service Center Orders\500D_Orders.xlsx"
	objMailItem.Attachments.Add "O:\CustSvc\Parts\SRVCTRS\Ops\Service Center Orders\500I_Orders.xlsx"
	objMailItem.Send
	Set objMailItem = Nothing
'	WScript.Sleep(2500)
'	objOutl.Quit
	Set objOutl = Nothing
	
End Sub



Sub CopyPasteReport
	
	vrc=session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").visiblerowcount
	rc=session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").rowcount
	n = Round(rc/vrc,0)+1
	
	If rc > vrc Then	
		For i = 1 To n
			On Error Resume Next
			session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").firstVisibleRow = vrc * i
			If Err.Number <> 0 Then
				Err.Clear
			End If
			WScript.Sleep(250)
		Next
		On Error Goto 0
	End If
	
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectAll
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").contextMenu
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItemByPosition "0"
	
'	Set ws = wb.Sheets(reportName)
'	ws.select
	ws.Range("A3").select
	ws.PasteSpecial , , , , , ,true
	ex.CutCopyMode = False
	ws.Cells(1,2).Value = "Created: " & Now()
	ws.Cells(1,1).select
	ws.Columns.autofit
End Sub
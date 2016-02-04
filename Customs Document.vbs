If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
On Error Resume Next
	Set connection =application.Children(0)
 If Err.Number <> 0 Then
	MsgBox("You are not connected to PMx,please connect and try again")
	On Error Goto 0
	WScript.Quit
 End if
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,vrc,vrc2
Dim messtxt,z,Row,SDRow,Itemno,SDNum,SAPRow
Dim workingRow,i,strRejectionReason
'************Ask for data file
'Set objDialog = CreateObject("UserAccounts.CommonDialog")

'objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx|All Files|*.*"
'objDialog.FilterIndex = 1
'objDialog.InitialDir = "C:\Scripts"
'intResult = objDialog.ShowOpen
 
'If intResult = 0 Then
'    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
'End If
'****************

Set wshShell = WScript.CreateObject( "WScript.Shell" )
strUserName = wshShell.ExpandEnvironmentStrings( "%USERNAME%" )

SDNum=InputBox("Enter Sales Document Number:","Document Number")
SDRow=0
Row=11
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open ("O:\CustSvc\Parts\Pmx Scripting\Script Data\Customs Document.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = SDNum
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
vrc=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").VisibleRowCount
'vrc=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").VisibleRowCount
SAPRow=0
ExcelSheet.Cells(14,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBKD-KURSK[85,0]").text
ExcelSheet.Cells(4,2).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/txtKUAGV-TXTPA").text
ExcelSheet.Cells(1,7).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
ExcelSheet.Cells(8,7).Value = Session.findbyid("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text

'MsgBox(vrc)
While Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text <>"__________________"
'While Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,0]
'While Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text <>"__________________"
If Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text="SUM" Then
	SDRow=SDRow+1
	SAPRow=SAPRow+1
End if
If (SDRow+1) = vrc Then
	Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG").verticalScrollbar.position=SDRow+1
   	SDRow=0
   	SAPRow=-1
End If
If ExcelSheet.Cells(Row,1).Value=ExcelSheet.cells(Row-1,1).value Then
   ExcelSheet.rows(Row).delete
End If

strRejectionReason=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/cmbVBAP-ABGRU[27,"&(SDRow)&"]").text
'MsgBox(strRejectionReason)
If strRejectionReason <> "Item captured in error" Then
	ExcelSheet.Cells(Row,1).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-POSNR[0,"&(SDRow)&"]").text
	ExcelSheet.Cells(100,100).Value= Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,3).Value = "'"&Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,2).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtRV45A-KWMENG[2,"&(SDRow)&"]").text
	'ExcelSheet.Cells(Row,3).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-KDMAT[6,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-NETPR[18,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,4).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-ARKTX[5,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,8).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WAERK[22,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,6).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/ctxtVBAP-WAERK[22,"&(SDRow)&"]").text
	ExcelSheet.Cells(Row,7).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4401/subSUBSCREEN_TC:SAPMV45A:4900/tblSAPMV45ATCTRL_U_ERF_AUFTRAG/txtVBAP-NETWR[21,"&(SDRow)&"]").text
	ExcelSheet.Rows(Row+1).EntireRow.Insert
		
	Row=Row+1
	'SDRow=SDRow+1
	'SAPRow=SAPRow+1
End if
'	MsgBox("vrc="&vrc&" Row="&Row&" SDRow="&SDRow&" SAPRow="&SAPRow)
'End If
'Row=Row+1
SDRow=SDRow+1
SAPRow=SAPRow+1
Wend


ExcelSheet.SaveAs "D:\Documents and Settings\"&strUserName&"\Desktop\"&SDNum&"-Customs Document.xlsx",51
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit


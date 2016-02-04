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
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,vrc,vrc2
Dim messtxt,z,Row,SDRow,Itemno,SDNum,SAPRow
Dim workingRow,i
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
End If
'****************

SDNum=InputBox("Enter Sales Document Number:","Document Number")
SDRow=0
Row=35
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVA23"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = SDNum
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
vrc=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").VisibleRowCount
SAPRow=0
'MsgBox(vrc)
While Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text <>"__________________"

ExcelSheet.Cells(Row,1).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtVBAP-POSNR[0,"&(SDRow)&"]").text
ExcelSheet.Cells(100,100).Value= Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
ExcelSheet.Cells(Row,2).Value = "'"&Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,"&(SDRow)&"]").text
ExcelSheet.Cells(Row,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtRV45A-KWMENG[2,"&(SDRow)&"]").text
ExcelSheet.Cells(Row,3).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtVBAP-KDMAT[6,"&(SDRow)&"]").text
ExcelSheet.Cells(Row,8).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtVBAP-NETWR[9,"&(SDRow)&"]").text
ExcelSheet.Cells(Row,4).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/txtVBAP-ARKTX[5,"&(SDRow)&"]").text

If SDRow+1=vrc Then
	Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT").verticalScrollbar.position=Row-35
    SDRow=0
    SAPRow=-1
End If
If ExcelSheet.Cells(Row,1).Value=ExcelSheet.cells(Row-1,1).value Then
	ExcelSheet.rows("Row:Row").delete()
End If
Row=Row+1
SDRow=SDRow+1
SAPRow=SAPRow+1
'MsgBox("vrc="&vrc&" Row="&Row&" SDRow="&SDRow&" SAPRow="&SAPRow)
Wend

ExcelSheet.Cells(5,5).Value = session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/ctxtVBAK-ANGDT").text
ExcelSheet.Cells(3,5).Value = session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/ctxtRV45A-KETDAT").text
ExcelSheet.Cells(2,2).Value = Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/ctxtVBAK-VBELN").text
ExcelSheet.Cells(3,2).Value = Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/txtKUAGV-TXTPA").text
ExcelSheet.Cells(4,2).Value = Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/subPART-SUB:SAPMV45A:4701/txtKUWEV-TXTPA").text
ExcelSheet.Cells(2,8).Value = Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/txtVBKD-BSTKD").text
Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
ExcelSheet.Cells(3,8).Value = Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4012/ctxtKUAGV-KUNNR").text
ExcelSheet.SaveAs "D:\Documents and Settings\dma02\Desktop\"&SDNum&"-HR.xlsx",51
Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07").select
vrc2=Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW").VisibleRowCount
'MsgBox("vrc2="&vrc2)
For i = 0 To vrc2-1
'MsgBox("i="&i)
	If Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&(i)&"]").text ="Customer ServiceRep" Then
		ExcelSheet.Cells(8,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[2,"&(i)&"]").text
	ElseIf Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&(i)&"]").text ="Customer ServiceMgr" Then
		ExcelSheet.Cells(9,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[2,"&(i)&"]").text
	ElseIf Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,"&(i)&"]").text ="Employee respons." Then
		ExcelSheet.Cells(10,5).Value = Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\07/ssubSUBSCREEN_BODY:SAPMV45A:4352/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/txtGVS_TC_DATA-REC-NAME1[2,"&(i)&"]").text
    End If
Next


'<<order description>>/app/con[0]/ses[0]/wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\12/ssubSUBSCREEN_BODY:SAPMV45A:4312/sub8309:SAPMV45A:8309/tabsZSD_ADDL_B_HEAD/tabpZSD_ADDL_B_HEAD_FC1/ssubZSD_ADDL_B_HEAD_SCA:ZSD_ADDL_DATA_B:8309/txtVBAK-KTEXT

Row=35
Do While ExcelSheet.Cells(Row,1).Value<>""
Call GetVendors
Loop

Call costOutBom
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit

Sub GetVendors
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nME03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	session.findById("wnd[0]/usr/ctxtEORD-WERKS").text ="500B"
	'For i = 3 To workingRow - 1
		'If ExcelSheet.Cells(Row,9).Value = "BUY" Then
			session.findById("wnd[0]/usr/ctxtEORD-MATNR").text = ExcelSheet.Cells(Row,2).Value
				If ExcelSheet.cells(Row,2).value="SUM" Then
				Row = Row+1
				Exit Sub
				End If
			session.findById("wnd[0]/tbar[0]/btn[0]").press
			If session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text <> "" Then
				ExcelSheet.Cells(Row,11).Value = "'"&session.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text
			Else
				ExcelSheet.Cells(Row,11).Value = "No Source List Record"
			End If 
			session.findById("wnd[0]/tbar[0]/btn[3]").press
		'End If
	Row = Row+1
	'Next
	
End Sub

Sub costOutBom
	workingRow = 35
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM03"
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	While ExcelSheet.Cells(workingRow,2).Value <> ""
		
		Call GetComponentCosts
		
		
	Wend
End Sub
	
	Sub GetComponentCosts
	
	If Len(Excelsheet.Cells(workingRow,2).Value) < 1 Then
		Exit Sub
	End If	
	If ExcelSheet.cells(workingRow,2).value="SUM" Then
		workingRow=workingRow+1
		Exit Sub
	End If
	Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(workingRow,2).Value
	session.findById("wnd[0]/tbar[0]/btn[0]").press
	If session.findById("wnd[0]/sbar").text <> "" Then
		ExcelSheet.Cells(workingRow,11).Value = session.findById("wnd[0]/sbar").text
		Exit Sub
	End If
	
	If session.findById("wnd[1]").text = "Select View(s)" Then
		session.findById("wnd[1]/tbar[0]/btn[20]").press
		session.findById("wnd[1]/tbar[0]/btn[0]").press
		On Error Resume Next
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500B"
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "G001"
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
		session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500B"
		Session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "G001"
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
'	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "F" Then
'		ExcelSheet.Cells(workingRow,10).Value = "BUY"
'	ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "E" Then
'		ExcelSheet.Cells(workingRow,10).Value = "MAKE"
'	ElseIf session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = "X" Then
'		ExcelSheet.Cells(workingRow,10).Value = "Make or Buy"
'	Else
'		ExcelSheet.Cells(workingRow,10).Value = "UNKNOWN"
'	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text <> 0 Then
		ExcelSheet.Cells(workingRow,12).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text/7
	Else
		ExcelSheet.Cells(workingRow,12).Value = "UNKNOWN"
	End If
	
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
	If session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = "V" Then
		ExcelSheet.Cells(workingRow,9).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text
	Else
		ExcelSheet.Cells(workingRow,9).Value =  session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text
	End If
	ExcelSheet.Cells(workingRow,13).Value = session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-LBKUM").text
	session.findById("wnd[0]/tbar[0]/btn[3]").press	
	workingRow = workingRow + 1
End Sub
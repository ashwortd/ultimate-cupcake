If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   On Error Resume next
   Set connection = application.Children(0)
   If Err.Number<>0 Then
   	MsgBox("You are not connected to PMx, please connect and try again")
   	On Error Goto 0
   	WScript.Quit
   	End If
   	
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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,currentTab
Dim messtxt,z,Row,WshShell,rowcheck,QMStatus
'************Ask for data file
'Set objDialog = CreateObject("UserAccounts.CommonDialog")

'objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
'objDialog.FilterIndex = 1
'objDialog.InitialDir = "C:\Scripts"
'intResult = objDialog.ShowOpen
 
'If intResult = 0 Then
'    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
'End If
'****************
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open ("O:\CustSvc\Parts\Pmx Scripting\Script Data\Material Master Creation for Plant 500Ba.xlsm")
Set ExcelSheet = ExcelWorkbook.Worksheets("Script Sheet")
Row=InputBox("Row to start at")-1
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
Do
rowcheck = ExcelSheet.Cells(Row+2,1).Value 
Call MMextend500B
Call ext99
Call extsubc
Call TaxClassCheck
'MsgBox (rowcheck)
Loop until rowcheck = "Blank"

ExcelApp.DisplayAlerts(False)
ExcelWorkbook.Close(True)
ExcelApp.Quit
MsgBox("Parts added")
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
		
Sub MMextend500B
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
row=row+1
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
Session.findbyid("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text=""
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
Session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 6
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = true

Session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(20).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = true
Session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ExcelSheet.Cells(Row,3).Value
Session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'messtxt=0
		Exit Sub
	End If
On Error Goto 0
On Error Resume next
	session.findById("wnd[1]").sendVKey 0
On Error Goto 0
currentTab=Session.activewindow.guifocus.ID
currentTab= Left(currentTab,50)
currentTab= Right(currentTab,8)
'MsgBox(currentTab)
If currentTab="tabpBABA" Then
	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
					session.findById("wnd[0]").sendVKey 0
					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					Call MMextend500B
				ElseIf z=2 Then 
					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					ExcelApp.Quit
					Set ExcelApp=Nothing
					Set ExcelWorkbook=Nothing
					Set ExcelSheet=Nothing
					WScript.Quit
				End if
	End If
End if	
If currentTab <>"tabpSP04" then
	session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").text = "c001_0000_tax_us"
	Session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").caretPosition = 16
	Session.findById("wnd[0]").sendVKey 0
	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = "gbo"
	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/tbar[1]/btn[8]").press	
End If

	

Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]").sendVKey 0

	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
					session.findById("wnd[0]").sendVKey 0
					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					Call MMextend500B
				ElseIf z=2 Then
					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3 
					ExcelApp.Quit
					Set ExcelApp=Nothing
					Set ExcelWorkbook=Nothing
					Set ExcelSheet=Nothing
					WScript.Quit
				End if
	End If	
		
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = ExcelSheet.Cells(Row,9).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = ExcelSheet.Cells(Row,10).Value
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMARA-MTPOS_MARA").text =""
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = ExcelSheet.Cells(Row,11).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = ExcelSheet.Cells(Row,8).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = ExcelSheet.Cells(Row,12).Value
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = ExcelSheet.Cells(Row,13).Value
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = ExcelSheet.Cells(Row,14).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = ExcelSheet.Cells(Row,15).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").text = ExcelSheet.Cells(Row,16).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text = ExcelSheet.Cells(Row,17).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,18).Value
Session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN")= ExcelSheet.Cells(Row,16).Value
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").setFocus
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").text = ExcelSheet.Cells(Row,19).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = ExcelSheet.Cells(Row,20).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,21).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = ExcelSheet.Cells(Row,22).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = ExcelSheet.Cells(Row,23).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,25).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = ExcelSheet.Cells(Row,24).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = ExcelSheet.Cells(Row,26).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = ExcelSheet.Cells(Row,27).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "d"
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
Session.findById("wnd[0]").sendVKey 0
QMStatus=session.findById("wnd[0]/sbar").Text
	If QMStatus ="Plants exist in which you have not specified a control key" Then
		Session.findById("wnd[0]").sendVKey 0
		QMStatus=""
	End If	
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = ExcelSheet.Cells(Row,28).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = ExcelSheet.Cells(Row,29).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = ExcelSheet.Cells(Row,30).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-EKALR").selected = True
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = ExcelSheet.Cells(Row,31).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,37).Value = session.findById("wnd[0]/sbar").Text

End Sub

Sub ext99
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
'extend material to distribution channel 99
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "99"
session.findById("wnd[1]").sendVKey 0
On Error Resume next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		Exit Sub
	End If
On Error Goto 0	
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = ExcelSheet.Cells(Row,9).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = ExcelSheet.Cells(Row,10).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = ExcelSheet.Cells(Row,11).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = ExcelSheet.Cells(Row,8).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").caretPosition = 15
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,34).Value = session.findById("wnd[0]/sbar").Text
End sub

Sub extsubc
'extend to SUBC 01 and 99
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 6
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(20).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = true
Session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "subc"
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume Next

	ExcelSheet.Cells(Row,56).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,56).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=46
		Exit Sub
	End If
On Error Goto 0	
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,35).Value = session.findById("wnd[0]/sbar").Text
ExcelSheet.Cells(Row,36).Value = session.findById("wnd[0]/sbar").Text
ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=6
End Sub

Sub TaxClassCheck
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Goto 0
session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = "gbo"
session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
End Sub

Sub endscript
WScript.Sleep(1000)
'MsgBox("The end has come")
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing

End sub
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
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim messtxt,z,Row
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
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Script Sheet")
Row=InputBox("Row to start at")-1
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
'Do Until ExcelSheet.Cells(Row,1).Value = "0"
'messtxt=ExcelSheet.cells(row,55)
Call MMextend500B
Sub MMextend500B
row=row+1
'messtxt= ExcelSheet.cells(row,55)
If ExcelSheet.Cells(Row,1).Value=("0") Then
		'Call endscript
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
	End If
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").setFocus
'session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
'Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
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
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 2
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 3
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 4
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 5
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'messtxt=0
		Call ext99
	End If
On Error Goto 0	
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = ExcelSheet.Cells(Row,7).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").caretPosition = 1
session.findById("wnd[0]").sendVKey 0

	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
					session.findById("wnd[0]").sendVKey 0
					Call MMextend500B
				ElseIf z=2 Then 
					ExcelApp.Quit
					Set ExcelApp=Nothing
					Set ExcelWoorkbook=Nothing
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
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").caretPosition = 15
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = ExcelSheet.Cells(Row,12).Value
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = ExcelSheet.Cells(Row,13).Value
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = ExcelSheet.Cells(Row,14).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = ExcelSheet.Cells(Row,15).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").text = ExcelSheet.Cells(Row,16).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").caretPosition = 1
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text = ExcelSheet.Cells(Row,17).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,18).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
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
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = ExcelSheet.Cells(Row,23).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,25).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = ExcelSheet.Cells(Row,24).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = ExcelSheet.Cells(Row,26).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = ExcelSheet.Cells(Row,27).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").caretPosition = 2
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = ExcelSheet.Cells(Row,28).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = ExcelSheet.Cells(Row,29).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").text = ExcelSheet.Cells(Row,30).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-VERPR").caretPosition = 6
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-EKALR").selected = True
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = ExcelSheet.Cells(Row,31).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").setFocus
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,37).Value = session.findById("wnd[0]/sbar").Text
Call ext99
End Sub

Sub ext99
'extend material to distribution channel 99
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
Session.findById("wnd[0]/usr/cmbRMMG1-MTART").setFocus
'session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "99"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'messtxt=0
		Call extsubc
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
Call extsubc
End sub

Sub extsubc
'extend to SUBC 01 and 99
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").setFocus
'session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 9
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
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
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,11]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 2
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 3
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 4
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 5
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).selected = true
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").caretPosition = 0
session.findById("wnd[1]/tbar[0]/btn[0]").press
'Session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "subc"
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'messtxt=0
		Call MMextend500B
	End If
On Error Goto 0	
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,35).Value = session.findById("wnd[0]/sbar").Text
ExcelSheet.Cells(Row,36).Value = session.findById("wnd[0]/sbar").Text

Call MMextend500B
End sub
'Move on to next row and startover
'Row=Row+1
'Loop

Sub endscript
'MsgBox("The end has come")
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing

End sub
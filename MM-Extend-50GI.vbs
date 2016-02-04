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
Dim messtxt,z,Row,iRow,strCurrentTab
Dim mrp3status,qmStatus,WndTTL

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open ("O:\CustSvc\Parts\Pmx Scripting\Script Data\MM 50GI-GE01 extend.xlsm")
Set ExcelSheet = ExcelWorkbook.Worksheets("50GI-Script Sheet")
Row=InputBox("Row to start at")
If TypeName(Row)="Empty" Then
	MsgBox("Goodbye!")
	ExcelWorkbook.Close(True)
	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
End If

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,1).Value <> ("0")
	Call MMextend50GI
	Call SelectTab
	WndTTL="nope"
	Row=Row+1
Loop

MsgBox("The end has come")
ExcelWorkbook.Close(True)
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub MMextend50GI
Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text=""
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").setFocus
Session.findById("wnd[0]").sendVKey 0
'Session.findById("wnd[0]").sendVKey 0
If session.findById("wnd[0]/sbar").Text<>"" Then
	ExcelSheet.Cells(Row,56)=session.findById("wnd[0]/sbar").Text
	Exit Sub
End If
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = ExcelSheet.Cells(Row,3).Value
Session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text="U03"
Session.findbyid("wnd[1]/usr/ctxtRMMG1-LGTYP").text="001"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").caretPosition = 2
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	ExcelSheet.Cells(Row,55).Value=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If ExcelSheet.Cells(Row,55).Value="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		Exit sub
	End If
On Error Goto 0	
On Error Resume next
	session.findById("wnd[1]").sendVKey 0
On Error Goto 0
'If strCurrentTab <>"tabpSP04" then
'	session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").text = "c001_0000_tax_us"
'	Session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").caretPosition = 16
'	Session.findById("wnd[0]").sendVKey 0
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = "gpe"
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
'	session.findById("wnd[0]").sendVKey 0
'	session.findById("wnd[0]/tbar[1]/btn[8]").press	
'End If

On Error Resume next
	session.findById("wnd[1]").sendVKey 0
	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
					session.findById("wnd[0]").sendVKey 0
					Exit Sub
				Else
					If z=2 Then
						ExcelWorkbook.Close(True) 
						ExcelApp.Quit
						Set ExcelApp=Nothing
						Set ExcelWoorkbook=Nothing
						Set ExcelSheet=Nothing
						WScript.Quit
				End If
			End if
	End If
End Sub
Sub SelectTab		
On Error Goto 0


Do
	strCurrentTab=Session.activewindow.guifocus.id
	strCurrentTab=Left(strCurrentTab,50)
	strCurrentTab=Right(strCurrentTab,8)
	'MsgBox(strCurrentTab)
	On Error Resume Next
	WndTTL=Session.findbyid("wnd[1]/usr/txtSPOP-DIAGNOSE1").text
	On Error Goto 0

If strCurrentTab="G1-MATNR" Then
	Exit Sub
End if

If strCurrentTab="tabpBABA" then
'If strCurrentTab="tabpBABA" Then
	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
					session.findById("wnd[0]").sendVKey 0
					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					Exit Sub
				Else
					If z=2 Then 
						ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
						ExcelWorkbook.Close(True)
						ExcelApp.Quit
						Set ExcelApp=Nothing
						Set ExcelWoorkbook=Nothing
						Set ExcelSheet=Nothing
						WScript.Quit
					End If
				End If
	End If
End If

'Session.findById("wnd[1]").sendVKey 0	
'If strCurrentTab <>"tabpSP04" then
'	session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").text = "c001_0000_tax_us"
'	Session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").caretPosition = 16
'	Session.findById("wnd[0]").sendVKey 0
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = "gpe"
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
'	session.findById("wnd[0]").sendVKey 0
'	session.findById("wnd[0]/tbar[1]/btn[8]").press	
'End If

'On Error Resume next
'	session.findById("wnd[1]").sendVKey 0
'	If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
'			z=MsgBox (ExcelSheet.Cells(Row,1).Value &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
'				If z=1 Then
'					session.findById("wnd[0]/tbar[0]/btn[15]").press
'					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
'					session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
'					session.findById("wnd[0]").sendVKey 0
'					Exit Sub
'				ElseIf z=2 Then
'					ExcelWorkbook.Close(True) 
'					ExcelApp.Quit
'					Set ExcelApp=Nothing
'					Set ExcelWoorkbook=Nothing
'					Set ExcelSheet=Nothing
'					WScript.Quit
'				End if
'	End If	
'On Error Goto 0

If strCurrentTab= "tabpSP04" Then	
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
					Exit Sub
				Else
					If z=2 Then
						ExcelWorkbook.Close(True) 
						ExcelApp.Quit
						Set ExcelApp=Nothing
						Set ExcelWoorkbook=Nothing
						Set ExcelSheet=Nothing
						WScript.ConnectObject session,     "off"
   	    				WScript.ConnectObject application, "off"
						WScript.Quit
					End If
				End If	
	End If
End If
if strCurrentTab="tabpSP05" Then
On Error Resume next
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMARA-MTPOS_MARA").text =ExcelSheet.Cells(Row,38).Value		
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = ExcelSheet.Cells(Row,9).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = ExcelSheet.Cells(Row,10).Value
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMARA-MTPOS_MARA").text =""
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = ExcelSheet.Cells(Row,11).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = ExcelSheet.Cells(Row,8).Value
Session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab= "tabpSP06" Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = ExcelSheet.Cells(Row,12).Value
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = ExcelSheet.Cells(Row,13).Value
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = ExcelSheet.Cells(Row,14).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = ExcelSheet.Cells(Row,15).Value
Session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP07"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").text = ExcelSheet.Cells(Row,16).Value
Session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab= "tabpSP08" Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text = ExcelSheet.Cells(Row,17).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
End If

If strCurrentTab="tabpSP09"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,18).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP11"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").text = ExcelSheet.Cells(Row,19).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
End If

If strCurrentTab="tabpSP12" Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = ExcelSheet.Cells(Row,20).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,21).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = ExcelSheet.Cells(Row,22).Value
Session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP13" Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = ExcelSheet.Cells(Row,23).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-SOBSL").text=ExcelSheet.Cells(Row,39).Value
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = ExcelSheet.Cells(Row,25).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = ExcelSheet.Cells(Row,24).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = ExcelSheet.Cells(Row,26).Value
Session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP14"Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text = ExcelSheet.Cells(Row,36).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text=ExcelSheet.Cells(Row,41).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text=ExcelSheet.Cells(Row,42).Value
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = ExcelSheet.Cells(Row,27).Value
Session.findById("wnd[0]").sendVKey 0
mrp3stat=session.findById("wnd[0]/sbar").Text
mrp3stat=Left(mrp3stat,5)
If mrp3stat="Check" Then
	session.findById("wnd[0]").sendVKey 0
	mrp3stat=""
End If
End If

If strCurrentTab="tabpSP15"Then
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-ALTSL").text=ExcelSheet.Cells(Row,43).Value
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP16"Then
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text=ExcelSheet.Cells(Row,44).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2525/txtMPOP-ANZPR").text=ExcelSheet.Cells(Row,45).Value
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP17"Then
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-FEVOR").text=ExcelSheet.Cells(Row,46).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP17/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2601/ctxtMARC-SFCPF").text=ExcelSheet.Cells(Row,47).Value
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP19" Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = "d"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP23"Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23").select
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = "PMX0003"
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = "USQP"
Session.findById("wnd[0]").sendVKey 0
qmStatus=session.findById("wnd[0]/sbar").Text
qmStatus=Left(qmStatus,6)
If qmStatus="Plants" Then
	Session.findById("wnd[0]").sendVKey 0
	qmstatus=""
End If
End If
If strCurrentTab="tabpSP24" Then
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = ExcelSheet.Cells(Row,28).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = ExcelSheet.Cells(Row,29).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = ExcelSheet.Cells(Row,30).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
End If

If strCurrentTab="tabpSP26" Then
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-EKALR").selected = True
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = true
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = ExcelSheet.Cells(Row,31).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
WndTTL="You are about to exit."
ExcelSheet.Cells(Row,37).Value = session.findById("wnd[0]/sbar").Text

'Else
'	session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").text = "c001_0000_tax_us"
'	Session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").caretPosition = 16
'	Session.findById("wnd[0]").sendVKey 0
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = "gpe"
'	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
'	session.findById("wnd[0]").sendVKey 0
'	session.findById("wnd[0]/tbar[1]/btn[8]").press	

End If
Loop Until WndTTL="You are about to exit."
End Sub
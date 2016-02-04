Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,currentTab
Dim messtxt,z,Row,WshShell,rowcheck,QMStatus,note1,StrPickle,check6,errMatMaintd,strParts
Dim connStr, objConn, getNames
Const adOpenStatic = 3
Const adLockOptimistic = 3
Const salesOrg = "5013"
Const extDistChl="01"
Const intDistChl="99"
Const salesPlant="500B"
Const storLoc="G001"
Const delPlant="500B"
Const taxClassInd="1"
Const availCheck="03"
Const transGrp="Z001"
Const loadGrp="0002"
Const profitCenter="5000000013"
Const purchGrp="ELV"
Const mrpType="PD"
Const lotSize="EX"
Const procureType="F"
Const grProcessTime=1
Const schedMargKey="000"
Const priceCont="S"
Const taxClass="c001_0000_tax_us"
Const qmControlKey="PMX0003"
Const certType="USQP"
Const forecastModel="N"
Const ccPhysInvInd="D"
Const valTaxClass="GBO"
Const ccFix=True
Const qmProcActive=True
Const wQtyStructure=True
Const matOrigin=True
Const subContChl="SUBC"
Const industrySector="A"
Const matType="ZENG"



'***********************
Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function
'***********************

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

'**************

'''''''''''''''''''''''''''''''''''''
'Define the driver and data source
'Access 2007, 2010, 2013 ACCDB:
'Provider=Microsoft.ACE.OLEDB.12.0
'Access 2000, 2002-2003 MDB:
'Provider=Microsoft.Jet.OLEDB.4.0
''''''''''''''''''''''''''''''''''''''
connStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\Users\dma02\Desktop\MaterialMasterExtension.accdb"
 
'Define object type
Set objConn = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")
'Open Connection
objConn.open connStr
 
'Define recordset and SQL query
'Set rs = objConn.execute("SELECT [Part Number],[NewValue],[AssignmentGrp],[ItmCatGrp],[CommCode],[SalesText],[PurchasingText],[MRP_Controller],[PlannedDelivery],[ValuationClass],[StandardPrice],[MaterialPricingGroup],[OriginGrp],[Extended] FROM ScriptData")
Set rs = objConn.execute("SELECT * FROM ScriptData")

'**************************
For x = 0 To 14
	strParts= strParts&" "&x&"-"&rs.Fields(x)
'	objRecordSet.Open "UPDATE [NewMaterialMasterEntrySheet] Set Extended = true WHERE PartNumber='"&rs.Fields(0)&"'",objConn,adOpenStatic, adLockOptimistic
Next
MsgBox strParts


session.findById("wnd[0]").maximize
session.StartTransaction("MM01")

Do While Not rs.EOF
 
Call MMextend500B
Call ext99
Call extsubc
Call TaxClassCheck

rs.MoveNext
Loop 
Call endscript

Sub MMextend500B

session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = rs.Fields(0)
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = industrySector
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = matType
Session.findbyid("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text=""
Session.findById("wnd[0]").sendVKey 0
	If sbarStatus ="Material type Project Materials copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End If
	If sbarStatus ="Material type Standard Components copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End if
Session.findById("wnd[1]/tbar[0]/btn[19]").press
Session.findById("wnd[1]/tbar[0]/btn[20]").press

session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(16).selected = false'14 for pe1
Session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = salesPlant
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = storLoc
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = salesOrg
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = extDistChl
Session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	errMatMaintd=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If errMatMaintd="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press
		Session.findById("wnd[1]/tbar[0]/btn[12]").press
		errMatMaintd="Clear"
		Exit Sub
	End If
On Error Goto 0
On Error Resume next
	session.findById("wnd[1]").sendVKey 0
On Error Goto 0
currentTab=Session.activewindow.guifocus.ID
currentTab= Left(currentTab,50)
currentTab= Right(currentTab,8)

If currentTab="tabpBABA" Then
	If sbarStatus="Fill in all required entry fields" Then
			z=MsgBox (rs.Fields(0) &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.StartTransaction("MM01")
					'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					StrPickle="Yes"
					z=0
					Exit Sub
				ElseIf z=2 Then 
					'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
'					ExcelApp.Quit
'					Set ExcelApp=Nothing
'					Set ExcelWorkbook=Nothing
'					Set ExcelSheet=Nothing
					Call endscript
				End if
	End If
End if	
If currentTab <>"tabpSP04" then
	session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").text = taxClass
	Session.findById("wnd[0]/usr/subSUBSCR_ZUORD:SAPLCLFM:1600/tblSAPLCLFMTC_OBJ_CLASS/ctxtRMCLF-CLASS[0,0]").caretPosition = 16
	Session.findById("wnd[0]").sendVKey 0
	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = valTaxClass
	Session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/tbar[1]/btn[8]").press	
End If

	

Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = delPlant
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2184/tblSAPLMGD1TC_STEUERN/ctxtMG03STEUER-TAXKM[4,0]").text = taxClassInd
session.findById("wnd[0]").sendVKey 0

	If sbarStatus="Fill in all required entry fields" Then
			z=MsgBox (rs.Fields(0) &" Basic data has not been entered into PMx",17,"PMx Material Extension Issue")
				If z=1 Then
					session.findById("wnd[0]/tbar[0]/btn[15]").press
					session.findById("wnd[1]/usr/btnSPOP-OPTION2").press
					session.StartTransaction("MM01")
					'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3
					StrPickle="Yes"
					Exit Sub
					
				ElseIf z=2 Then
'					ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=3 
'					ExcelApp.Quit
'					Set ExcelApp=Nothing
'					Set ExcelWorkbook=Nothing
'					Set ExcelSheet=Nothing
					Call endscript
				End if
	End If	
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select		
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = rs.Fields(11)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = rs.Fields(2)
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMARA-MTPOS_MARA").text =""
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = rs.Fields(3)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = rs.Fields(1)
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text = availCheck
On Error Resume Next
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARA-TRAGR").text = transGrp
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2162/ctxtMARC-LADGR").text = loadGrp
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = profitCenter
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07/ssubTABFRA1:SAPLMGMM:2004/subSUB2:SAPLMGD1:2205/ctxtMARC-STAWN").text = rs.Fields(4)
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").text = rs.Fields(5)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP08/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2121/cntlLONGTEXT_VERTRIEBS/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = purchGrp
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").text = rs.Fields(6)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP11/ssubTABFRA1:SAPLMGMM:2010/subSUB2:SAPLMGD1:2321/cntlLONGTEXT_BESTELL/shellcont/shell").setSelectionIndexes 68,68
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
'Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2481/ctxtMARC-MAABC").text = rs.fields(?)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISMM").text = mrpType
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = rs.Fields(7)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2483/ctxtMARC-DISLS").text = lotSize
session.findById("wnd[0]").sendVKey 0
On Error Resume Next
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-BESKZ").text = procureType
On Error Goto 0
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-PLIFZ").text = rs.Fields(8)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/txtMARC-WEBAZ").text = grProcessTime
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2485/ctxtMARC-FHORI").text = schedMargKey
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB4:SAPLMGD1:2493/txtMARC-WZEIT").text = rs.Fields(14)
Session.findById("wnd[0]").sendVKey 0
On Error Resume Next
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16").select
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP16/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2524/ctxtMPOP-PRMOD").text =forecastModel
On Error Goto 0
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/chkMARC-CCFIX").selected = ccFix
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2701/ctxtMARC-ABCIN").text = ccPhysInvInd
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/chkMARA-QMPUR").selected = qmProcActive
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-SSQSS").text = qmControlKey
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP23/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2752/ctxtMARC-QZGTP").text = certType
Session.findById("wnd[0]").sendVKey 0

	If sbarStatus ="Plants exist in which you have not specified a control key" Then
		Session.findById("wnd[0]").sendVKey 0
		
	End If	
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text = rs.Fields(9)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-VPRSV").text = priceCont
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text = rs.Fields(10)
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-EKALR").selected = wQtyStructure
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/chkMBEW-HKMAT").selected = matOrigin
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2904/ctxtMBEW-HRKFT").text = rs.Fields(12)
session.findById("wnd[0]/tbar[0]/btn[11]").press
End Sub

Sub ext99
If StrPickle="Yes" Then
	Exit Sub
End if
session.StartTransaction("MM01")
'extend material to distribution channel 99
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = rs.Fields(0)
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = industrySector
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = matType
session.findById("wnd[0]").sendVKey 0
If sbarStatus="Material type Project Materials copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End If
If sbarStatus ="Material type Standard Components copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End if
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(16).selected = false
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = salesPlant
Session.findbyid("/app/con[0]/ses[0]/wnd[1]/usr/ctxtRMMG1-VKORG").text=salesOrg
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = intDistChl
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
	note1=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If note1="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press
		Session.findById("wnd[1]/tbar[0]/btn[12]").press
		note1="None"
		Exit Sub
	End If
On Error Goto 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = delPlant
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").setFocus
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").caretPosition = 4
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KONDM").text = rs.Fields(11)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-KTGRM").text = rs.Fields(2)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text = rs.Fields(3)
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-PRODH").text = rs.Fields(1)
session.findById("wnd[0]/tbar[0]/btn[11]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
'ExcelSheet.Cells(Row,34).Value = session.findById("wnd[0]/sbar").Text
End sub

Sub extsubc
If StrPickle="Yes" Then
	Exit Sub
End if

'extend to SUBC 01 and 99
session.StartTransaction("MM01")
WScript.Sleep(2000)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = rs.Fields(0)
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = industrySector
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = matType
session.findById("wnd[0]").sendVKey 0
If sbarStatus ="Material type Project Materials copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End If
If sbarStatus ="Material type Standard Components copied from master record" Then
		Session.findById("wnd[0]").sendVKey 0
	End if
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(16).selected = false
Session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = salesPlant
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = subContChl
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = salesOrg
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = extDistChl
session.findById("wnd[1]").sendVKey 0
On Error Resume Next

	errMatMaintd=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If errMatMaintd="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press:
		Session.findById("wnd[1]/tbar[0]/btn[12]").press:
		'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=46
		errMatMaintd="Clear"
		Exit Sub
	End If
On Error Goto 0
For i =1 To 5	
session.findById("wnd[0]").sendVKey 0
Next

On Error Resume Next
check6=Session.findbyid("wnd[1]").text
On Error Goto 0
If check6="Last data screen reached" Then
	Session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").press
	'ExcelSheet.Cells(Row,35).Value = session.findById("wnd[0]/sbar").Text
	'ExcelSheet.Cells(Row,36).Value = session.findById("wnd[0]/sbar").Text
	'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=6
	check6="None"
	Exit Sub
End If

session.findById("wnd[0]/tbar[0]/btn[11]").press
'ExcelSheet.Cells(Row,35).Value = session.findById("wnd[0]/sbar").Text
'ExcelSheet.Cells(Row,36).Value = session.findById("wnd[0]/sbar").Text
'ExcelWorkbook.Worksheets("New Material Master Entry Sheet").cells(Row,1).entirerow.interior.colorindex=6
End Sub

Sub TaxClassCheck
If StrPickle="Yes" Then
	StrPickle="No"
	Exit Sub
End if

session.StartTransaction("MM02")
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = rs.Fields(0)
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Goto 0
session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").text = valTaxClass
session.findById("wnd[0]/usr/subSUBSCR_BEWERT:SAPLCTMS:5000/tabsTABSTRIP_CHAR/tabpTAB1/ssubTABSTRIP_CHAR_GR:SAPLCTMS:5100/tblSAPLCTMSCHARS_S/ctxtRCTMS-MWERT[1,0]").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[11]").press
objRecordSet.Open "UPDATE [NewMaterialMasterEntrySheet] Set Extended = true WHERE PartNumber='"&rs.Fields(0)&"'",objConn,adOpenStatic, adLockOptimistic
End Sub

Sub endscript
'Close connection and release objects
objConn.Close
Set rs = Nothing
Set objConn = Nothing
WScript.Sleep(1000)
WScript.Quit
'ExcelApp.Quit
'Set ExcelApp=Nothing
'Set ExcelWorkbook=Nothing
'Set ExcelSheet=Nothing

End sub
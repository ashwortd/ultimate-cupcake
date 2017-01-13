'File: SC_Material_Extend.vbs
'Author: Derek Ashworth
'Edit Date: 05/17/2016
Option Explicit

'File Definitions
Const fileName="SC_Material_Data.xlsm"
Const fileDirectory="\\winfile02\data\CustSvc\Parts\Pmx Scripting\Script Data\"
Const showWindow = True
Dim excelFileLocation,excelApp,excelWorkbook,excelWorksheet
Dim intRow,userName,password,window
Dim strMaterialNum
Dim strIndSect
Dim strMatType
Dim strOrgPlant
Dim strOrgSLoc
Dim strOrgSOrg
Dim strDistCh
Dim strWarehouseNo
Dim strStorType
Dim strTaxClass
Dim strValClass
Dim strSO1DelPlant
Dim strSO1TaxClass
Dim strSO2ProdHier
Dim strSO2MatPrGrp
Dim strSO2ActAssGrp
Dim strSO2ItmCatGrp
Dim strSalesAvChk
Dim strSalesTransGrp
Dim strSalesLoadGrp
Dim strSalesProfitCtr
Dim strSalesMatSerial
Dim strFTECommCd
Dim strSalesTxt
Dim strPurchPgrp
Dim strPOText
Dim strMrp1MrpTp
Dim strMrp1MrpCont
Dim strMrp1LotSz
Dim strMinLotSz
Dim strMrp2ProcType
Dim strMrp2SpecProc
Dim strMrp2InHsProd
Dim strMrp2PDT
Dim strMrp2GRProcTm
Dim strMrp2SchedMargKey
Dim strMrp2ProdStoeLoc
Dim strMrp2StorLocEP
Dim strMrp2SafetyStk
Dim strMrp3Dim strGrp
Dim strMrp3ConsMode
Dim strMrp3BwdConsPer
Dim strMrp3MxdMrp
Dim strMrp3TRLT
Dim strMrp4IndCol
Dim strMrp4MatMemo
Dim strMrp4SpecStkTp
Dim strMrp4SlocMrpInd
Dim strMrp4ROP
Dim strMrp4RepQty
Dim strForecastFModel
Dim strForecastFPeriod
Dim strWSUntIss
Dim strWSProdSched
Dim strWSPrdSchedProf
Dim strWSBatMan
Dim strQMContKey
Dim strQMCertTyp
Dim strPltDataS1CCPhysInvInd
Dim strStorBin
Dim strAcct1ValClass
Dim strAcct1Price
Dim strAcctPrCont
Dim strCst1CtrGroup
'Functions

Function selViewFirstPos()
selViewFirstPos = session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,0]").text
End Function
Function SetView(x)
 session.findById("wnd[1]/tbar[0]/btn[19]").press
	If x=1 Then
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
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 10
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(17).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(24).selected = True
		
	ElseIf x=2 Then
		
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
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 6
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(20).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = true
	End If
End Function

Function strMsgBox(window)
	strMsgBox=session.findById("wnd["&window&"]/usr/txtMESSTXT1").text
End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function

Function currentTab()
	currentTab=Session.activewindow.guifocus.ID
	currentTab= Left(currentTab,50)
	currentTab= Right(currentTab,8)
End Function
'File locations

intRow=InputBox("What is the starting row to extend?")
Set excelApp=CreateObject("Excel.Application")
excelFileLocation=fileDirectory&fileName
Set excelWorkbook=excelApp.workbooks.open(excelFileLocation)
excelApp.visible=True
OpenSAP()

sub openSAP
	' Open SAP
	Dim WshShell
	set WshShell = WScript.CreateObject("WScript.Shell")

	' Not yet completed
	If not(WshShell.AppActivate("SAP Logon")) then
		WshShell.Exec("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")
		Wscript.Sleep 500
		
		Dim i : i = 0
		Do While not(WshShell.AppActivate("SAP Logon"))
			WScript.Sleep 250
			timeoutCheck i, 400, "SAP Logon Timeout"		' Loop a max of 10 seconds
		Loop
	End if
	
	' Run GUI Script
	Dim application, SapGuiAuto, connection, session, isNewConn
	If Not IsObject(application) Then
	   Set SapGuiAuto  = GetObject("SAPGUI")
	   Set application = SapGuiAuto.GetScriptingEngine
	End If
	If Not IsObject(connection) Then
		If application.Children.Count > 0 then				' If it has connections
			Set connection = application.Children(0)
			isNewConn = false
			If not connection.description = "1.1 PMx Production (PE1)" then
				Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true)
				isNewConn = true
			End if
		Else
			Set connection = application.OpenConnection("1.1 PMx Production (PE1)", true)
			isNewConn = true
		End if
	End If
	If Not IsObject(session) Then
	   Set session = connection.Children(0)
	End If
	If IsObject(WScript) Then
	   WScript.ConnectObject session,     "on"
	   WScript.ConnectObject application, "on"
	End If
	session.findById("wnd[0]").maximize
	
	' Login
	If isNewConn Then
		userName=InputBox("SAP PE1 Username:")
		password=InputBox("SAP PE1 Password:")
		session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = userName
		session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
		session.findById("wnd[0]").sendVKey 0

		' If any messages come up clear them
		Dim messageCount, logonOption
		messageCount = 0
		Do while session.Children.Count > 1
			if messageCount > 5 then
				MsgBox "Error, too many message boxes detected"
				Wscript.quit
				exit do
			else
				Set logonOption = session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1", false)
				' Check for message to bump off another person logged on
				if TypeName(logonOption) <> "Nothing" then
					logonOption.select
				End if
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End if
			messageCount = messageCount + 1
		Loop
		
		
	Else
		Dim sessionCount
		sessionCount = connection.Children.Count
		
		session.CreateSession
		do while connection.Children.Count <= sessionCount
			WScript.Sleep 250
		loop
		Set session = connection.Children(connection.Children.Count - 1)
	End If
End Sub

'Warehouse manage facility *******

'*********************************

'Set Values **********************
strMaterialNum = excelWorksheet.Cells(Row,1).Value
strIndSect = "A"
strMatType = "ZENG"
strOrgPlant = excelWorksheet.Cells(Row,4).Value
strOrgSLoc = excelWorksheet.Cells(Row,5).Value
strOrgSOrg	= excelWorksheet.Cells(Row,6).Value
strDistCh = excelWorksheet.Cells(Row,7).Value
strWarehouseNo = excelWorksheet.Cells(Row,8).Value
strStorType = excelWorksheet.Cells(Row,9).Value
strTaxClass = "C001_0000_TAX_US"
strValClass = excelWorksheet.Cells(Row,11).Value
strSO1DelPlant = excelWorksheet.Cells(Row,12).Value
strSO1TaxClass = "1"
strSO2ProdHier = excelWorksheet.Cells(Row,14).Value
strSO2MatPrGrp = excelWorksheet.Cells(Row,15).Value
strSO2ActAssGrp = excelWorksheet.Cells(Row,16).Value
strSO2ItmCatGrp = excelWorksheet.Cells(Row,17).Value
strSalesAvChk = "03"
strSalesTransGrp = "Z001"
strSalesLoadGrp = "0002"
strSalesProfitCtr = excelWorksheet.Cells(Row,21).Value
strSalesMatSerial = excelWorksheet.Cells(Row,22).Value
strFTECommCd = excelWorksheet.Cells(Row,23).Value
strSalesTxt = excelWorksheet.Cells(Row,24).Value
strPurchPgrp = excelWorksheet.Cells(Row,25).Value
strPOText = excelWorksheet.Cells(Row,26).Value
strMrp1MrpTp = "PD"
strMrp1MrpCont = excelWorksheet.Cells(Row,28).Value
strMrp1LotSz = "EX"
strMinLotSz = "ROQ"
strMrp2ProcType = excelWorksheet.Cells(Row,31).Value
strMrp2SpecProc = excelWorksheet.Cells(Row,32).Value
strMrp2InHsProd = excelWorksheet.Cells(Row,33).Value
strMrp2PDT = excelWorksheet.Cells(Row,34).Value
strMrp2GRProcTm = excelWorksheet.Cells(Row,35).Value
strMrp2SchedMargKey = excelWorksheet.Cells(Row,36).Value
strMrp2ProdStoeLoc = excelWorksheet.Cells(Row,37).Value
strMrp2StorLocEP = excelWorksheet.Cells(Row,38).Value
strMrp2SafetyStk = "ROP"
strMrp3StrGrp = excelWorksheet.Cells(Row,40).Value
strMrp3ConsMode = excelWorksheet.Cells(Row,41).Value
strMrp3BwdConsPer = excelWorksheet.Cells(Row,42).Value
strMrp3MxdMrp = excelWorksheet.Cells(Row,43).Value	
strMrp3TRLT = excelWorksheet.Cells(Row,44).Value
strMrp4IndCol = excelWorksheet.Cells(Row,45).Value
strMrp4MatMemo = excelWorksheet.Cells(Row,46).Value
strMrp4SpecStkTp = excelWorksheet.Cells(Row,47).Value
strMrp4SlocMrpInd = excelWorksheet.Cells(Row,48).Value
strMrp4ROP = excelWorksheet.Cells(Row,49).Value
strMrp4RepQty = excelWorksheet.Cells(Row,50).Value
strForecastFModel ="N"
strForecastFPeriod = "12"
strWSUntIss = excelWorksheet.Cells(Row,53).Value
strWSProdSched = excelWorksheet.Cells(Row,54).Value
strWSPrdSchedProf="Z00002"
strWSBatMan = excelWorksheet.Cells(Row,56).Value
strQMContKey = "PMX0003"
strQMCertTyp = "USQP"
strPltDataS1CCPhysInvInd ="D"
strStorBin = excelWorksheet.Cells(Row,60).Value
strAcct1ValClass =excelWorksheet.Cells(Row,61).Value
strAcct1Price= excelWorksheet.Cells(Row,62).Value
strAcctPrCont = excelWorksheet.Cells(Row,63).Value
strCst1CtrGroup = excelWorksheet.Cells(Row,64).Value
'end of values *************


Sub extendMaterial
	session.StartTransaction("MM01")
	session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = strMaterialNum
	session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = strIndSect
	session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = strMatType
	session.findById("wnd[0]").sendVKey 0
		If sbarStatus ="Material type Project Materials copied from master record" Then
			Session.findById("wnd[0]").sendVKey 0
		End If
		If sbarStatus ="Material type Standard Components copied from master record" Then
			Session.findById("wnd[0]").sendVKey 0
		End if
	session.findById("wnd[1]/tbar[0]/btn[20]").press
	
	If selViewFirstPos="Basic Data 1" Then 'Set proper views for user rights
 		SetView(1)
	Else
 		SetView(2)
 	End if
End Sub





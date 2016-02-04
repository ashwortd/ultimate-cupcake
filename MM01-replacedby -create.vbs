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
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")

objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
 
If intResult = 0 Then
    Wscript.Quit
'Else
'    Wscript.Echo objDialog.FileName
End If
'****************
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,wndstatus,sapstatus
Dim test,startplace,sapstat2
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Open(objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet2")
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,4).Value<>""
Call addpart
Loop

Sub addpart()

Session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text =ExcelSheet.Cells(Row,4).Value
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]").sendVKey 0
On Error Resume next
Session.findById("wnd[1]/tbar[0]/btn[20]").press
Session.findbyid("wnd[0]/sbar").text=sapstatus
sapstatus=Left(sapstatus,9)
'MsgBox(sapstatus)
	If sapstatus="The group" Then
		WScript.Sleep(500)
		Exit Sub
	End if
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,14]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,14]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 1
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 2
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 3
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 4
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 5
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 6
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,15]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,15]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 7
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(23).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").setFocus
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,16]").caretPosition = 0
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 8
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 9
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = ExcelSheet.Cells(Row,6).Value
session.findById("wnd[1]").sendVKey 0
On Error Resume Next
Startplace= Right(Left(Session.ActiveWindow.GuiFocus.ID,50),8)
'MsgBox startplace
'/app/con[0]/ses[0]/wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB1:SAPLMGD1:1001/txtMAKT-MAKTX
If startplace= "tabpSP09" Then
		Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,9).Value
		Session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]/tbar[0]/btn[0]").press
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,8).Value
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").setFocus
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").caretPosition = 3
		'session.findById("wnd[0]").sendVKey 4
		'session.findById("wnd[1]").close
		session.findById("wnd[0]").sendVKey 0
		Session.FINDBYID("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").TEXT=ExcelSheet.Cells(Row,10).Value
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text=ExcelSheet.Cells(Row,11).Value
		session.findById("wnd[0]").sendVKey 0
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text=ExcelSheet.Cells(Row,12).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = ExcelSheet.Cells(Row,13).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = ExcelSheet.Cells(Row,14).Value
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT2").text = ExcelSheet.Cells(Row,15).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = ExcelSheet.Cells(Row,16).Value
		session.findById("wnd[0]").sendVKey 0
		If Session.findbyid("wnd[0]/sbar").text <>"" Then
			Session.findById("wnd[0]").sendVKey 0
		End If
		Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").text = ExcelSheet.Cells(Row,17).Value
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
		
		ExcelSheet.Cells(Row,18).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
		Exit Sub	
		End If	
If startplace="tabpSP12" Then
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,8).Value
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").setFocus
		'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").caretPosition = 3
		'session.findById("wnd[0]").sendVKey 4
		'session.findById("wnd[1]").close
		session.findById("wnd[0]").sendVKey 0
		Session.FINDBYID("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").TEXT=ExcelSheet.Cells(Row,10).Value
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text=ExcelSheet.Cells(Row,11).Value
		session.findById("wnd[0]").sendVKey 0
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text=ExcelSheet.Cells(Row,12).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = ExcelSheet.Cells(Row,13).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = ExcelSheet.Cells(Row,14).Value
		Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT2").text = ExcelSheet.Cells(Row,15).Value
		session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = ExcelSheet.Cells(Row,16).Value
		session.findById("wnd[0]").sendVKey 0
		If Session.findbyid("wnd[0]/sbar").text <>"" Then
			Session.findById("wnd[0]").sendVKey 0
		End If
		Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").text = ExcelSheet.Cells(Row,17).Value
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[0]").sendVKey 0
		session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
		
		ExcelSheet.Cells(Row,18).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
		Exit Sub
End if		
wndstatus = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
'MsgBox wndstatus
If wndstatus = "Material already maintained for this" Then
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Session.findbyid("wnd[1]").close
	ExcelSheet.Cells(Row,18).Value = (wndstatus)
	Row=Row+1
	wndstatus=0
	Exit sub
End If
wndstatus = Session.findbyid("wnd[2]/usr/txtMESSTXT1").text
'MsgBox wndstatus
If wndstatus = "Entry 50GC 0001  does not exist in T001L (check" Then
	Session.findbyid("wnd[2]/tbar[0]/btn[0]").press
	Session.findbyid("wnd[1]").close
	ExcelSheet.Cells(Row,18).Value = (wndstatus)
	Row=Row+1
	wndstatus=0
	Exit sub
End If
On Error Goto 0
test = Right(Left(Session.ActiveWindow.GuiFocus.ID,50),8)
If test="tabpSP21" Then
	session.findById("wnd[0]").sendVKey 0
	Session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
	ExcelSheet.Cells(Row,18).Value = session.findById("wnd[0]/sbar").Text
	Row=Row+1
	test="0"
	Exit Sub
End if
Session.findById("wnd[0]").sendVKey 0
If Session.findbyid("wnd[0]/sbar").text="Material not yet created in supplying plant" Then
	Session.findbyid("wnd[0]").sendVKey 0
	End if
session.findById("wnd[0]").sendVKey 0


session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = ExcelSheet.Cells(Row,7).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/ctxtMARC-EKGRP").text = ExcelSheet.Cells(Row,9).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[0]/btn[0]").press
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = ExcelSheet.Cells(Row,8).Value
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").setFocus
'session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").caretPosition = 3
'session.findById("wnd[0]").sendVKey 4
'session.findById("wnd[1]").close
On Error Resume Next
session.findById("wnd[0]").sendVKey 0
sapstat2=session.findById("wnd[0]/sbar").Text
sapstat2=Left(sapstat2,7)
'MsgBox (sapstat2)
If sapstat2 = "The MRP" then
	ExcelSheet.Cells(Row,18).Value = session.findById("wnd[0]/sbar").Text
	Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text ="001"
	session.findById("wnd[0]").sendVKey 0
End If
On Error Goto 0	
Session.FINDBYID("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").TEXT=ExcelSheet.Cells(Row,10).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGPRO").text=ExcelSheet.Cells(Row,11).Value
session.findById("wnd[0]").sendVKey 0
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-STRGR").text=ExcelSheet.Cells(Row,12).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-VRMOD").text = ExcelSheet.Cells(Row,13).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT1").text = ExcelSheet.Cells(Row,14).Value
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/txtMARC-VINT2").text = ExcelSheet.Cells(Row,15).Value
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2492/ctxtMARC-MISKZ").text = ExcelSheet.Cells(Row,16).Value
session.findById("wnd[0]").sendVKey 0
If Session.findbyid("wnd[0]/sbar").text <>"" Then
	Session.findById("wnd[0]").sendVKey 0
End If
Session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2495/ctxtMARC-SBDKZ").text = ExcelSheet.Cells(Row,17).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0

If Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text="" then
   Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/ctxtMBEW-BKLAS").text="4200"
   Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP24/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2800/subSUB2:SAPLMGD1:2802/txtMBEW-STPRS").text="10.00"	
End If
session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press

ExcelSheet.Cells(Row,18).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub

'****Close SAP Connection And Excel
MsgBox("The end has come")
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
	
'File: DelPlant_ItemCatGrp.vbs
'Author: Derek Ashworth
'Creation Date: 03/07/2014

Dim oExcel,oWorkbook,oSheet
Dim Row,Status,PopupStatus,elementID,elementLeft,ElementFinal
Dim PopupStatus1

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
'************Ask for data file
Set objDialog = CreateObject("UserAccounts.CommonDialog")
objDialog.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
objDialog.FilterIndex = 1
objDialog.InitialDir = "C:\Scripts"
intResult = objDialog.ShowOpen
If intResult = 0 Then
    Wscript.Quit
End If
'****************
Set oExcel = CreateObject("Excel.Application")
Set oWorkbook = oExcel.Workbooks.Open (objDialog.FileName)
Set oSheet = oWorkbook.Worksheets(1)
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
session.findById("wnd[0]").sendVKey 0
Row=InputBox("Which row would you like to start?","Excel Sheet Starting Point")
Do
PopupStatus="none"
PopupStatus1="None"
Call Main
Row=Row+1
Loop While oSheet.cells(Row,1).value<>""
oWorkbook.Close(True)
oExcel.Quit
Set oExcel=Nothing
Set oWorkbook=Nothing
Set oSheet=Nothing
MsgBox("Changes Completed - Check Excel File for processing errors")
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub Main
session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text=oSheet.cells(Row,1).value
session.findbyid("wnd[0]/tbar[1]/btn[5]").press
session.findbyid("wnd[1]/tbar[0]/btn[20]").press
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = false
'session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = false
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text=oSheet.cells(Row,2).value
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text=oSheet.cells(Row,3).value
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text=""
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text=""
On Error Goto 0
session.findbyid("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
PopupStatus1 = session.findbyid("wnd[1]").text
PopupStatus = session.findbyid("wnd[2]").text
'MsgBox(PopupStatus)
'MsgBox(PopupStatus1)
	If PopupStatus = "Warning" Then
			oSheet.cells(Row,11)=session.findbyid("wnd[2]/usr/txtMESSTXT1").text
			PopupStatus="none"
		ElseIf PopupStatus <>"none" Then
			oSheet.Cells(Row,11)=session.findbyid("wnd[2]/usr/txtMESSTXT1").text
			oSheet.Cells(Row,12)=session.findById("wnd[0]/sbar").Text
			session.findbyid("wnd[2]/tbar[0]/btn[0]").press
On Error Goto 0
			
		
	End If
	If PopupStatus1="Class Type 4 Entries" Then
		session.findbyid("wnd[1]/tbar[0]/btn[12]").press
		PopupStatus1="None"
	ElseIf PopupStatus1<>"None" Then
		oSheet.Cells(Row,2)=session.findbyid("wnd[1]").text
		PopupStatus1="None"
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nMM02"
		session.findById("wnd[0]").sendVKey 0
		Exit sub
	End If
	
'elementID = session.ActiveWindow.GuiFocus.ID
'elementLeft = Left(elementID, 50)
'elementFinal = Right(elementLeft, 8)
'MsgBox (elementFinal)
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text=oSheet.cells(Row,4).value
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
Status=session.findById("wnd[0]/sbar").Text
	If Status ="Material not yet created in supplying plant" Then
		oSheet.cells(Row,9).value = Status
		Status = "none"
		session.findById("wnd[0]").sendVKey 0
	End if
'session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2157/ctxtMVKE-MTPOS").text=oSheet.cells(Row,8).value
session.findById("wnd[0]").sendVKey 0
session.findbyid("wnd[0]/tbar[0]/btn[11]").press
oSheet.cells(Row,10).value=session.findById("wnd[0]/sbar").Text
End Sub

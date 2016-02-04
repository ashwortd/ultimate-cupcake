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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,strPARVW,Row,strStartPos

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

Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=true
Set ExcelWorkbook = ExcelApp.Workbooks.Open(objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")
 If Row=""Then 
 	ExcelApp.Quit
	Set ExcelApp=Nothing
	Set ExcelWorkbook=Nothing
	Set ExcelSheet=Nothing
	WScript.ConnectObject session,     "off"
	WScript.ConnectObject application, "off"
	WScript.Quit
 End if
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Do Until ExcelSheet.Cells(Row,1).Value = ""
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
Session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = ExcelSheet.Cells(Row,2).Value
Session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = ExcelSheet.Cells(Row,3).Value
Session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/tbar[0]/btn[0]").press
strStartPos=Session.findbyid("wnd[0]").guifocus.ID
strStartPos=Right(strStartPos,10)
If strStartPos="TITLE_MEDI" Then
	Session.findbyid("wnd[0]/tbar[1]/btn[27]").press
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if
If strStartPos="KNVV-BZIRK" Then
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if

vrc=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").visiblerowcount
For i = 0 To 14
strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
 'If strPARVW ="Z0" Then
 '	ExcelSheet.Cells(Row,6).Value=Session.findbyid("/app/con[0]/ses[0]/wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text
 'End if
 If strPARVW ="Z9" Then
 	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").setFocus
 	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/btnDELETE_ROW").press
 	'ExcelSheet.Cells(Row,7).Value=Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text
 
 End if
Next
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = vrc
For i = 0 To 14
strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
' If strPARVW ="Z0" Then
' 	ExcelSheet.Cells(Row,6).Value=Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text
' End if
 If strPARVW ="Z9" Then
 	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").setFocus
 	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/btnDELETE_ROW").press
 	'ExcelSheet.Cells(Row,5).Value=Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/txtRF02D-VTXTM[3,"&(i)&"]").text
 End If
Next

session.findById("wnd[0]/tbar[0]/btn[11]").press
status1=session.findById("wnd[0]/sbar").Text
status1 = Right(status1,6)
 If status1 = "member" Then
 	 Session.findbyid("wnd[0]/tbar[0]/btn[0]").press
 End if
ExcelSheet.Cells(Row,6).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Loop 
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

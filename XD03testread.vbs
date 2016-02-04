If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,strPARVW,c,d,strVTXTM1,Row,strStartPos
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\ScriptData\XD03testread.xls")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
Row=InputBox("Row to start at")

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Do Until ExcelSheet.Cells(Row,1).Value = ""
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = ExcelSheet.Cells(Row,4).Value
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = ExcelSheet.Cells(Row,5).Value
session.findById("wnd[1]/tbar[0]/btn[0]").press
strStartPos=Session.findbyid("wnd[0]").guifocus.ID
strStartPos=Right(strStartPos,10)
'MsgBox(strStartPos)
If strStartPos="TITLE_MEDI" Then
	Session.findbyid("wnd[0]/tbar[1]/btn[27]").press
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if
If strStartPos="KNVV-BZIRK" Then
	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
End if

vrc=session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").visiblerowcount
'MsgBox(vrc)
'session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = 3
'session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = 6
'If strPARVW <>"______________________________" then
'MsgBox(strPARVW)
	c=6
	d=7
For i = 0 To 14
 	'MsgBox(i)
	
strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
'MsgBox(strPARVW)

 If strPARVW ="Z0" Then
 	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").setFocus
 	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/btnDELETE_ROW").press
 End if
c=c+2
d=d+2

Next
session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN").verticalScrollbar.position = vrc
For i = 0 To 14
 	'MsgBox(i)

strPARVW = session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").text 
'MsgBox(strPARVW)
 If strPARVW ="Z0" Then
 	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/tblSAPMF02DTCTRL_PARTNERROLLEN/ctxtKNVP-PARVW[0,"&(i)&"]").setFocus
 	Session.findbyid("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7324/btnDELETE_ROW").press
 End if
c=c+2
d=d+2
Next
'End if
session.findById("wnd[0]/tbar[0]/btn[11]").press
Row=Row+1
Loop 
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing

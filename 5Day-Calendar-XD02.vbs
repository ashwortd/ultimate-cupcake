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
Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("G:\ScriptData\Calendarupdate-xd02.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Row to start at")-1
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd02"
session.findById("wnd[0]").sendVKey 0
Call custslct
Sub custslct()
	row=row+1
	If ExcelSheet.Cells(Row,1).Value=("") Then
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
	End If
	Session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,4).Value
	Session.findById("wnd[1]").sendVKey 0
	Call Pickcell
End Sub

Sub Pickcell()
	Session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05").select
	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7340/tblSAPMF02DTCTRL_ABLADESTELLEN/txtKNVA-ABLAD[0,0]").text = "Factory"
	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7340/tblSAPMF02DTCTRL_ABLADESTELLEN/ctxtKNVA-KNFAK[2,0]").text = "u8"
	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7340/tblSAPMF02DTCTRL_ABLADESTELLEN/ctxtKNVA-KNFAK[2,0]").setFocus
	session.findById("wnd[0]/usr/subSUBTAB:SAPLATAB:0100/tabsTABSTRIP100/tabpTAB05/ssubSUBSC:SAPLATAB:0201/subAREA1:SAPMF02D:7340/tblSAPMF02DTCTRL_ABLADESTELLEN/ctxtKNVA-KNFAK[2,0]").caretPosition = 2
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	ExcelSheet.Cells(Row,5).Value = session.findById("wnd[0]/sbar").Text
	Call custslct
End sub

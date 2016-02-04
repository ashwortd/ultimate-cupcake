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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Set ExcelApp = CreateObject("Excel.Application")
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\dma02\Desktop\order-output.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva03"
session.findById("wnd[0]").sendVKey 0
While ExcelSheet.Cells(Row,1).Value <>""
Session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/mbar/menu[3]/menu[11]/menu[0]/menu[0]").select
session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3").columns.elementAt(0).width = 4
session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/txtDV70A-MSGNA[2,0]").setFocus
session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3/txtDV70A-MSGNA[2,0]").caretPosition = 6
ExcelSheet.Cells(Row,2).Value=Session.findbyid("wnd[0]/usr/tblSAPDV70ATC_NAST3/lblDV70A-STATUSICON[0,0]").text
ExcelSheet.Cells(Row,3).Value=Session.findbyid("wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1,0]").text
ExcelSheet.Cells(Row,4).Value=Session.findbyid("wnd[0]/usr/tblSAPDV70ATC_NAST3/txtNAST-DATVR[8,0]").text
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Wend
ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   	    WScript.ConnectObject application, "off"
	 	WScript.Quit


'********************************************************
'
' Delete material from ZSD_SLOC_DET_2 table
'
' created 2-20-13
' created by danielle thomas
'
'********************************************************

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
Dim Row,DelivPlant,StatusText,FindRow,materialnum,CurrDate

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\Delete Material from ZSD_SLOC_DET_2 table_Data.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

CurrDate = Date

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")

session.findById("wnd[0]").resizeWorkingPane 185,9,false
session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_sloc_det_2"
session.findById("wnd[0]").sendVKey 0


' ** Do Until Loop will execute until the last row (a blank row) is found **
Do Until ExcelSheet.Cells(Row,1).Value = ""

	session.findById("wnd[0]/usr/btnVIM_POSI_PUSH").press
	
	Session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = ExcelSheet.Cells(Row,1).Value  'sales org
	session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[1,21]").text = ExcelSheet.Cells(Row,2).Value  'plant
	session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[2,21]").text = ExcelSheet.Cells(Row,3).Value  'material
	session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[3,21]").text = ExcelSheet.Cells(Row,4).Value  'sloc
	session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[3,21]").setFocus
	session.findById("wnd[1]/tbar[0]/btn[0]").press
	
	MaterialNum = session.findById("wnd[0]/usr/tblSAPLZSD_SLOC_DETTCTRL_ZV_SD_SLOC_DET/ctxtZV_SD_SLOC_DET-MATNR[2,0]").text
	'MsgBox("matl num = " & materialnum)
	
	If MaterialNum = ExcelSheet.Cells(Row,3).Value then
		session.findById("wnd[0]/tbar[1]/btn[8]").press			'Select entry shown
		Session.findById("wnd[0]/tbar[1]/btn[14]").press		'delete button   	REMOVE COMMENT WHEN SCRIPT WORKS
		Session.findById("wnd[0]/tbar[0]/btn[11]").press		'save button		REMOVE COMMENT WHEN SCRIPT WORKS
		
		'StatusText = session.findById("wnd[0]/sbar").text		'capture status bar text REMOVE COMMENT WHEN SCRIPT WORKS
		ExcelSheet.Cells(Row,5).Value = "Deleted from SLOC_DET_2 " & CurrDate	'past status bar text in excel data file REMOVE COMMENT WHEN SCRIPT WORKS
	End If
		
	
	Row = Row + 1
	
Loop

ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
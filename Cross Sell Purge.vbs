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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,i,stopcheck,vrc,b
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
'Next line sets the location of the excel spreadsheet
Set ExcelWorkbook = ExcelApp.Workbooks.Add()
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
vrc=session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").visiblerowcount
b=1
Row=1
i=0
Do
If i=vrc Then
	session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position = 24*b
	i=0
	b=b+1
End if
Call main
Row=Row+1
i=i+1
stopcheck=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,"&(i)&"]").text
'MsgBox(stopcheck)
Loop While stopcheck <>"__________________"

		'Call endscript
		MsgBox("The end has come")
		ExcelWorkbook.SaveAs("D:\Documents and Settings\dma02\Desktop\CrossSell.xlsx")
		ExcelWorkbook.Close(True)
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit


Sub main

ExcelSheet.Cells(Row,1).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,"&(i)&"]").text
ExcelSheet.Cells(Row,2).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txtRV130-TEXTL[1,"&(i)&"]").text
ExcelSheet.Cells(Row,3).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,"&(i)&"]").text
ExcelSheet.Cells(Row,4).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txt*MAAPV-ARKTX[3,"&(i)&"]").text
ExcelSheet.Cells(Row,5).Value=Session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/chkMV13D-DETAI[6,"&(i)&"]").value
End Sub
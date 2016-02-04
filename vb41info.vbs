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
Dim objShell, ExcelApp,ExcelWorkbook,ExcelSheet,vrc
Dim exRow,Row,loopCount
Dim i

Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory


Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=true
Set ExcelWorkbook = ExcelApp.Workbooks.Add
ExcelWorkbook.SaveAs(strpath & "\Cross-Sell.xlsx")
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")

	

session.findById("wnd[0]").maximize
vrc=session.findbyid("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").visiblerowcount
'MsgBox(vrc)
i=0
exRow=1
loopcount=0
Call main
Sub main
'On Error Resume next
session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position = 0

Do 
		
'			MsgBox(i)
If i=vrc Then
	loopCount=loopCount+1
	session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position = (loopCount*vrc)
	i=0
End If		
		ExcelSheet.Cells(exRow,1)=session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,"&(i)&"]").text
		ExcelSheet.Cells(exRow,2)=session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txtRV130-TEXTL[1,"&(i)&"]").text
		ExcelSheet.Cells(exRow,3)=session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKONDD-SMATN[2,"&(i)&"]").text
		ExcelSheet.Cells(exRow,4)=session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/txt*MAAPV-ARKTX[3,"&(i)&"]").text
		
'If i=vrc Then
'	loopCount=loopCount+1
'	session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY").verticalScrollbar.position = (loopCount*vrc)
'	i=0
'End If
exRow=exRow+1
i=i+1 
'On Error Goto 0
Loop While Session.findById("wnd[0]/usr/tblSAPMV13DTCTRL_FAST_ENTRY/ctxtKOMGD-MATNR[0,"&(i-1)&"]").text <>"__________________"
End Sub

If ExcelSheet.Cells((exRow-1),1)=""Then
	WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
    MsgBox("done")
End If
Call main

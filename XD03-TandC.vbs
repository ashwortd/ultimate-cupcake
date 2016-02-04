Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim Row,SAPRow,TextDesc,TxtRow,ExColumn,TextLine
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
'Else
'    Wscript.Echo objDialog.FileName
End If
'****************

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)
SAPRow=0
Row=InputBox("Starting Row?")
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nxd03"
session.findById("wnd[0]").sendVKey 0
Do
Call Main
Row=Row+1
Loop Until ExcelSheet.Cells(Row,1).Value=""
		ExcelWorkbook.Close(True)
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWorkbook=Nothing
		Set ExcelSheet=Nothing
	    WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit

Sub Main
session.findById("wnd[1]/usr/ctxtRF02D-KUNNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRF02D-BUKRS").text = "5000"
session.findById("wnd[1]/usr/ctxtRF02D-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRF02D-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").text = "00"
session.findById("wnd[1]/usr/ctxtRF02D-SPART").setFocus
session.findById("wnd[1]/usr/ctxtRF02D-SPART").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume Next
Winstat=session.findbyid("wnd[2]").text
If Winstat="Error" Then
	ExcelSheet.Cells(Row,4).Value=session.findbyid("wnd[2]/usr/txtMESSTXT1").text
	session.findbyid("wnd[2]").close
	ExcelSheet.Cells(Row,2).Value=session.findbyid("wnd[1]/usr/txtINT_KNA1-NAME1").text
	Exit Sub
End If
On Error Goto 0
	'ExcelSheet.Cells(Row,2).Value=session.findbyid("wnd[0]/usr/subSUBKOPF:SAPMF02D:7001/txtINT_KNA1-NAME1").text
session.findById("wnd[0]/mbar/menu[3]/menu[6]").select
Do
	TextDesc= session.findById("wnd[0]/usr/subSUBTAB:SAPMF02D:3502/tblSAPMF02DTCTRL_TEXTE/txtRTEXT-TTEXT[2,"&SAPRow&"]").text
	If TextDesc="Master/Standard T’s and C’s" Then
		ExcelSheet.Cells(Row,3).Value=	session.findbyid("wnd[0]/usr/subSUBTAB:SAPMF02D:3502/tblSAPMF02DTCTRL_TEXTE/txtRTEXT-LTEXT[3,"&SAPRow&"]").setfocus
		session.findById("wnd[0]").sendVKey 2
		TxtRow=1
		ExColumn=10
		x=1
		Do
		'MsgBox(TxtRow)
		If TxtRow=22 Then
			session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA").verticalScrollbar.position = 20*x
			x=x+1
			TxtRow=2
		End if
		TextLine=session.findById("wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,"&TxtRow&"]").text
		ExcelSheet.Cells(Row,ExColumn).Value=TextLine
		TxtRow=TxtRow+1
		ExColumn=ExColumn+1
		Loop Until TextLine="________________________________________________________________________"
		TxtRow=1
		ExColumn=10
		ExcelSheet.Cells(Row,9).Value="Checked on 5/16/2014"
		
	End If
	SAPRow=SAPRow+1
Loop Until TextDesc="Master/Standard T’s and C’s"
SAPRow=0
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
End Sub



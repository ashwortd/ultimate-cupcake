Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Dim IniStat, LineNum,ItmChk,WndName,ItmRow,GRComplete
Dim Itemtest(3),ibchk,ibchk2,log1,LogStat,WndName1
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
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmigo"
session.findById("wnd[0]").sendVKey 0
Row=2
Do
If ItmChk = False Then
	ItmRow=0
End If

WndName="nope"
ibchk2=True
Call InboundChk
Call Main
Call NextItem
Call CheckNPost
Loop While ExcelSheet.Cells(Row,1).Value<>""

MsgBox("The end has come")
ExcelWorkbook.Close(True)
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit

Sub Main
If ibchk2=False Then
	Exit Sub
End if
'session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").hierarchyHeaderWidth = 154
'session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell[1]").topNode = "          1"
session.findbyid("wnd[0]/tbar[1]/btn[5]").press
On Error Resume Next
WndName1=session.findbyid("wnd[1]").text
If WndName1="Restart" Then
	session.findbyid("wnd[1]/usr/btnSPOP-OPTION2").press
End If
On Error Goto 0
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0011/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = ExcelSheet.Cells(Row,12).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/ctxtGODYNPRO-PO_NUMBER").text = ExcelSheet.Cells(Row,12).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/txtGODYNPRO-PO_ITEM").setFocus
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_FIRSTLINE:SAPLMIGO:0010/subSUB_FIRSTLINE_REFDOC:SAPLMIGO:2000/txtGODYNPRO-PO_ITEM").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
IniStat=session.findById("wnd[0]/sbar").Text
If Right(IniStat,5)="items" Then
	ExcelSheet.Cells(Row,27).Value = IniStat
	GRComplete="Yes"
	IniStat="none"
	ibchk2=False
	Exit Sub
End If
If IniStat <>"" Then
	ExcelSheet.Cells(Row,27).Value = IniStat
	IniStat="none"
	Exit Sub
End If
	
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").select
session.findbyid("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text=ExcelSheet.Cells(Row,16).Value
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = true
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BLDAT").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT").text = ExcelSheet.Cells(Row,20).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = ExcelSheet.Cells(Row,16).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/ctxtGOITEM-ERFME").text = "ea"
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = ExcelSheet.Cells(Row,18).Value

'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").setFocus
End Sub

Sub NextItem
If ibchk2=False Then
	Exit Sub
End if
Do
	ItmChk=True	
	If ExcelSheet.Cells(Row,12).Value = ExcelSheet.Cells(Row-1,12).Value Then
		ItmRow=ItmRow+1
		session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/btnOK_NEXT_ITEM").press
		LineNum=session.findbyid("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/ctxtGOITEM-EBELP[30,"&(ItmRow)&"]").text
	Else
		ItmChk=False
	End If
		
		 If LineNum=ExcelSheet.Cells(Row,13).Value then
            session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = True
			session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = ExcelSheet.Cells(Row,16).Value
			session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = ExcelSheet.Cells(Row,12).Value
		    
		  Else 
		  	ItmChk=False
		  End If
Loop until ItmChk =False
End Sub

Sub CheckNPost
If ibchk2=False Then
	If GRComplete="Yes" Then
		ExcelSheet.cells(Row,26)="No Items Available"
		GRComplete="No"
		Row=Row+1
		Exit Sub
	Else
		ExcelSheet.cells(Row,26)="inbound delivery"
		Row=Row+1
		Exit Sub
	End If
End if
session.findById("wnd[0]/tbar[1]/btn[7]").press
On Error Resume Next
 WndName=session.findbyid("wnd[1]").text 
On Error Goto 0
	If WndName="Display logs" Then
		LogStat=session.findbyid("wnd[1]/usr/lbl[10,3]").text
		LogStat=Left(LogStat,4)
		If LogStat="User" Then
			session.findbyid("wnd[1]/tbar[0]/btn[0]").press
			ExcelSheet.Cells(Row,26).Value="Being Processed by User"
			Row=Row+1
			Exit Sub
		End if
		log1=ExcelSheet.cells(Row,12)
		'ExcelSheet.Cells(Row,25).Value="Not possible for : "&log1
		session.findbyid("wnd[1]/tbar[0]/btn[0]").press
		session.findById("wnd[0]/tbar[1]/btn[23]").press
		ExcelSheet.Cells(Row,26).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
		Exit Sub
	End If

session.findById("wnd[0]/tbar[1]/btn[23]").press
ExcelSheet.Cells(Row,26).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub

Sub InboundChk
ibchk=ExcelSheet.cells(Row,22)
If ibchk="Inbound" Then
	ibchk2=False
End If
End Sub

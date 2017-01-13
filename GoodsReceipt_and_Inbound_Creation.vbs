Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row
Dim IniStat, LineNum,ItmChk,WndName,ItmRow,GRComplete
Dim Itemtest(3),ibchk,ibchk2,log1,LogStat,WndName1
Dim inboundItemNm
Dim strDate
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
file = ChooseFile(defaultLocalDir)
MsgBox file
Function wnd1Status()
	wnd1Status = Session.findbyid("wnd[1]").text
End Function

Function sbarStatus()
	sbarStatus = Session.findbyid("wnd[0]/sbar").text
End Function

Function ChooseFile (ByVal initialDir)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

    Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    Dim winVersion

    ' This collection should contain just the one item
    For Each objItem in colItems
        'Caption e.g. Microsoft Windows 7 Professional
        'Name e.g. Microsoft Windows 7 Professional |C:\windows|...
        'OSType e.g. 18 / OSArchitecture e.g 64-bit
        'Version e.g 6.1.7601 / BuildNumber e.g 7601
        winVersion = CInt(Left(objItem.version, 1))
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing

    If (winVersion <= 5) Then
        ' Then we are running XP and can use the original mechanism
        Set cd = CreateObject("UserAccounts.CommonDialog")
        cd.InitialDir = initialDir
        cd.Filter = "VBScript Data Files |*.xls;*.xlsx;*.xlsm|All Files|*.*"
        ' filter index 4 would show all files by default
        ' filter index 1 would show zip files by default
        cd.FilterIndex = 1
        If cd.ShowOpen = True Then
            ChooseFile = cd.FileName
        Else
            ChooseFile = ""
        End If
        Set cd = Nothing    

    Else
        ' We are running Windows 7 or later
        Set shell = CreateObject( "WScript.Shell" )
        Set ex = shell.Exec( "mshta.exe ""about: <input type=file id=X><script>X.click();new ActiveXObject('Scripting.FileSystemObject').GetStandardStream(1).WriteLine(X.value);close();resizeTo(0,0);</script>""" )
        ChooseFile = Replace( ex.StdOut.ReadAll, vbCRLF, "" )

        Set ex = Nothing
        Set shell = Nothing
    End If
End Function
'****************
Set ExcelApp = CreateObject("Excel.Application")
ExcelApp.Visible=True
Set ExcelWorkbook = ExcelApp.Workbooks.Open (file)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

strDate=InputBox("What is the posting date (MM/DD/YYYY) ?")

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
Call InboundCreate

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
If Right(sbarStatus,8)="released" Then
	ExcelSheet.Cells(Row,34).Value=sbarStatus
	ibchk2=False
	Exit Sub
End if
IniStat=session.findById("wnd[0]/sbar").Text
If Right(IniStat,5)="items" Then
	ExcelSheet.Cells(Row,31).Value = IniStat
	GRComplete="Yes"
	IniStat="none"
	ibchk2=False
	Exit Sub
End If
If IniStat <>"" Then
	ExcelSheet.Cells(Row,31).Value = IniStat
	IniStat="none"
	Exit Sub
End If
	
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES").select
session.findbyid("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text=ExcelSheet.Cells(Row,17).Value
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT.").select
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/subSUB_DETAIL_TAKE:SAPLMIGO:0304/chkGODYNPRO-DETAIL_TAKE").selected = true
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BUDAT").text = ExcelSheet.Cells(Row,2).Value
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/txtGOHEAD-BKTXT").text = ExcelSheet.Cells(Row,22).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/txtGOITEM-ERFMG").text = ExcelSheet.Cells(Row,16).Value
'session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_QUANTITIES/ssubSUB_TS_GOITEM_QUANTITIES:SAPLMIGO:0315/ctxtGOITEM-ERFME").text = "ea"
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = ExcelSheet.Cells(Row,20).Value

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
			session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_ITEMDETAIL:SAPLMIGO:0301/subSUB_DETAIL:SAPLMIGO:0300/tabsTS_GOITEM/tabpOK_GOITEM_DESTINAT./ssubSUB_TS_GOITEM_DESTINATION:SAPLMIGO:0325/txtGOITEM-SGTXT").text = ExcelSheet.Cells(Row,19).Value
		    
		  Else 
		  	ItmChk=False
		  End If
Loop until ItmChk =False
End Sub

Sub CheckNPost
If ibchk2=False Then
	If GRComplete="Yes" Then
		ExcelSheet.cells(Row,32)="No Items Available"
		GRComplete="No"
		Row=Row+1
		Exit Sub
	Else
		ExcelSheet.cells(Row,32)="inbound delivery"
		Row=Row+1
		Exit Sub
	End If
End if
session.findById("wnd[0]/tbar[1]/btn[7]").press
On Error Resume Next
If Right(wnd1Status,4)="logs" Then
	ExcelSheet.cells(Row,34).value=session.findbyid("wnd[1]/usr/lbl[10,3]").text
	session.findbyid("wnd[1]/tbar[0]/btn[0]").press
	session.findbyid("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BUDAT").text=strDate'edit for year when needed.
	
End If
On Error Goto 0
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
		ExcelSheet.Cells(Row,27).Value = session.findById("wnd[0]/sbar").Text
		Row=Row+1
		Exit Sub
	End If

session.findById("wnd[0]/tbar[1]/btn[23]").press
ExcelSheet.Cells(Row,31).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End Sub

Sub InboundChk
ibchk=ExcelSheet.cells(Row,28)
If ibchk="Inbound" Then
	ibchk2=False
End If
End Sub
Sub InboundCreate()
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl31n"
	session.findById("wnd[0]").sendVKey 0
	Row=2
	Do
	Call makeInbound
	Row=Row+1
	Loop While ExcelSheet.Cells(Row,1).Value<>""
End Sub
	
Sub makeInbound()
	Do
	If ExcelSheet.cells(Row,28)="Inbound" Then
		Exit Do
	 Else
	 	Row=Row+1
	End If
	Loop While ExcelSheet.Cells(Row,1).Value<>""
	If  ExcelSheet.Cells(Row,1).Value="" Then
		Exit Sub
	End If
	session.findById("wnd[0]/usr/txtLV50C-BSTNR").text = ExcelSheet.Cells(Row,12).Value
	session.findById("wnd[0]").sendVKey 0
	If Left(sbarStatus,4)="Purc" Then
		ExcelSheet.Cells(Row,31).Value = session.findById("wnd[0]/sbar").Text
		Exit Sub
	End If
	If Right(sbarStatus,3)="key" Then
		ExcelSheet.Cells(Row,31).Value = session.findById("wnd[0]/sbar").Text
		Exit Sub
	End If
		If Left(sbarStatus,4)="Item" Then
		session.findById("wnd[0]").sendVKey 0
	End If
	
	i=0
	Do
	session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
	inboundItemNm = session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPS-POSNR[0,"&i&"]").text
	'MsgBox (CInt(inboundItemNm))
	If CInt(inboundItemNm) =ExcelSheet.Cells(Row,13).Value Then
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPSD-G_LFIMG[6,"&i&"]").text = ExcelSheet.Cells(Row,16).Value
	 Else
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPSD-G_LFIMG[6,"&i&"]").text = ""
	End If
	i=i+1
	'MsgBox(session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPS-POSNR[0,"&i-1&"]").text)
	Loop Until session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPS-POSNR[0,"&i&"]").text=""	
	
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	If Left(sbarStatus,5)="Notif" Then
		session.findById("wnd[0]").sendVKey 0
		ExcelSheet.Cells(Row,31).Value = session.findById("wnd[0]/sbar").Text
		session.findById("wnd[0]/tbar[0]/okcd").text = "/nvl31n"
		session.findById("wnd[0]").sendVKey 0
		Exit Sub
	End If
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/ctxtRV50A-LFDAT_LA").text = ExcelSheet.Cells(Row,2).Value
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/txtLIKP-BOLNR").text = ExcelSheet.Cells(Row,22).Value
	session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/ctxtLIKP-TRATY").text="Y6"
	session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/txtLIKP-TRAID").text=ExcelSheet.Cells(Row,21).Value
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/txtLIKP-BOLNR").setFocus
	session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\03/ssubSUBSCREEN_BODY:SAPMV50A:2208/txtLIKP-BOLNR").caretPosition = 18
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	ExcelSheet.Cells(Row,31).Value = session.findById("wnd[0]/sbar").Text
End Sub

'This portion of the script takes the freight portal twice daily report and checks the 'Consignee Name' for internal or external shipment.
'This depends on a text file that identifies the vendor names need to keep it up to date. Names must be exact as they are on the shipping report
'script utilizes, text files, scripting dictionary, and multiple instances of the same variable.
' Derek Ashworth
' 7/9/2014

Const ForReading = 1
Dim File
Dim FileToRead
Dim strLine
Dim infoResult
Dim strDir, objFile, returnvalue
Dim ExcelSheet,ExcelApp,ExcelWorkbook
Dim Row,NumLines,test2,POField,strCount,objRange,test3,EDIResult
Dim PMxReadRow,ItmNum,ItmCat,ItmNumFlag,EDIField,completeField
Dim costCollector

'************Ask for data file
file = ChooseFile(defaultLocalDir)
MsgBox file

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
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")

strDir = "C:\Users\dma02\Inbound-Goods_Receipts\"
File1 = "Ship-to_suppliers.txt"
FileToRead = strDir & File1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(FileToRead, ForReading)
objFile.ReadAll
NumLines=objFile.Line
objFile.Close
Set objFile = objFSO.OpenTextFile(FileToRead, ForReading)
ReDim strLine(NumLines-1)
i=0
Row=1
Do Until objFile.AtEndOfStream
    strLine(i) = objFile.ReadLine
    i=i+1
Loop
Set objDictionary=CreateObject("Scripting.Dictionary")
For Each r In strLine
	If Not objdictionary.Exists(r) Then
		objdictionary.Add r,r
	End If
Next
Sub transaction_type
	test3=ExcelSheet.cells(Row,3)
	test3=CStr(test3)
	test2=ExcelSheet.cells(Row,5)
	If objdictionary.Exists(test2) Then
 		infoResult="Inbound"
 	Else
 		infoResult="Goods Receipt"
 	End If
	ExcelSheet.cells(row,28).value= infoResult
	If objdictionary.Exists(test3) Then
	  EDIResult="EDI Vendor"
	 Else
	 	EDIResult="Non-EDI Vendor"
	 End If
	 ExcelSheet.cells(Row,30)=EDIResult    
End Sub
Row=1
Do Until excelsheet.cells(row+1,1)=""
	Row = row+1
	Call transaction_type
Loop
ExcelSheet.cells(1,28).value="Shipment Type"
ExcelSheet.cells(1,29).value="Subcontract Check"
Set objRange = ExcelApp.Range("A1","W1")
objrange.Font.Bold=True
objrange.Font.ColorIndex=2
objrange.Interior.ColorIndex =41
ExcelWorkbook.Save
Call countStr
Sub countStr
Row=2
	Do Until ExcelSheet.cells(Row,1)=""
	
		POField=ExcelSheet.cells(Row,12).value
		strCount=Len(POField)
			If strCount=6 Then
				Set objRange = ExcelSheet.cells(Row,1).EntireRow
				objRange.Delete
				Row=Row-1
			End If
		Row=Row+1
	Loop
Row=2
	Do Until ExcelSheet.cells(Row,1)=""
	
	EDIField=ExcelSheet.cells(Row,30)
	If EDIField = "EDI Vendor" Then
		Set objRange = ExcelSheet.cells(Row,1).EntireRow
		objRange.Delete
		Row=Row-1
	End If
	Row=Row+1
	Loop
Row=2
	Do Until ExcelSheet.cells(Row,1)=""
	
	completeField=ExcelSheet.cells(Row,18)
	If completeField = "N/A" Then
		Set objRange = ExcelSheet.cells(Row,1).EntireRow
		objRange.Delete
		Row=Row-1
	End If
	Row=Row+1
	Loop
Row=2
	Do Until ExcelSheet.cells(Row,1)=""
	costCollector=ExcelSheet.cells(Row,24)
	If costCollector = "Project Number" Then
		Set objRange = ExcelSheet.cells(Row,1).EntireRow
		objRange.Delete
		Row=Row-1
	End If
	Row=Row+1
	Loop
		
End Sub
ExcelWorkbook.Save
'Check for PMx Connection
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
' end of check 
Call CheckPO
Sub CheckPO
	Row=2
	PMxReadRow=0
	
	session.findById("wnd[0]/tbar[0]/okcd").text = "/nME23n"
	session.findById("wnd[0]").sendVKey 0
	Do Until ExcelSheet.cells(Row,1)=""
		ItmNumFlag="No"
		PMxReadRow=0
		If ExcelSheet.cells(Row,28)="Goods Receipt" then
			session.findbyid("wnd[0]/tbar[1]/btn[17]").press
			Session.findbyid("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text=ExcelSheet.cells(Row,12)
			session.findbyid("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/radMEPO_SELECT-BSTYP_F").select
			session.findbyid("wnd[1]/tbar[0]/btn[0]").press
			Do Until ItmNumFlag="Yes"
				On Error Resume Next
				ItmNum=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,"&PMxReadRow&"]").text
				ItmNum=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,"&PMxReadRow&"]").text
				ItmNum=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,"&PMxReadRow&"]").text
				ItmNum=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,"&PMxReadRow&"]").text
				ItmNum=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-EBELP[1,"&PMxReadRow&"]").text
				
				On Error Goto 0 
				ItmNum=CInt(ItmNum)
				test2=ExcelSheet.cells(Row,13)
				If test2 ="" Then
					test2=10
				End if				
			If ItmNum=test2 Then
				If ExcelSheet.cells(Row,12)=ExcelSheet.cells(Row+1) Then
					itmNumFlag="No"
				Else
					ItmNumFlag="Yes"
				End if
				On Error Resume next
				ItmCat=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0015/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,"&PMxReadRow&"]").text
				ItmCat=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,"&PMxReadRow&"]").text
				ItmCat=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,"&PMxReadRow&"]").text
				ItmCat=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,"&PMxReadRow&"]").text
				ItmCat=session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0019/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EPSTP[3,"&PMxReadRow&"]").text
				On Error Goto 0
					If ItmCat="L" Then
						ExcelSheet.cells(Row,28).value="Inbound"
					End If
					If ItmCat<>"L" Then
						ExcelSheet.cells(Row,29).value="not Subcontracted"
					End If
				Row=Row+1
			End If
			If ItmNum <> test2 Then
				PMxReadRow=PMxReadRow+1
			End If
			Loop
			
		End if
		If ExcelSheet.cells(Row,28)="Inbound" Then
			ExcelSheet.cells(Row,29).value="checked"
			Row=Row+1
		End If
		
	Loop
End Sub


'******************************
Dim IniStat, LineNum,ItmChk,WndName,ItmRow,GRComplete
Dim Itemtest(3),ibchk,ibchk2,log1,LogStat,WndName1
Dim inboundItemNm


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
session.findById("wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0003/subSUB_HEADER:SAPLMIGO:0101/subSUB_HEADER:SAPLMIGO:0100/tabsTS_GOHEAD/tabpOK_GOHEAD_GENERAL/ssubSUB_TS_GOHEAD_GENERAL:SAPLMIGO:0110/ctxtGOHEAD-BLDAT").text = ExcelSheet.Cells(Row,2).Value
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
	i=0
	Do
	session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02").select
	inboundItemNm = session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPS-POSNR[0,"&i&"]").text
	'MsgBox (CInt(inboundItemNm))
	If CInt(inboundItemNm) =ExcelSheet.Cells(Row,13).Value Then
		session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV50A:1208/tblSAPMV50ATC_LIPS_TRAN_INB/txtLIPSD-G_LFIMG[6,"&i&"]").text = ExcelSheet.Cells(Row,17).Value
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
ExcelWorkbook.Save
Set ExcelSheet=Nothing
Set ExcelWorkbook=Nothing
Set ExcelApp=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"	
objFile.Close
Set objFSO = Nothing
Set objFile = Nothing
WScript.Quit
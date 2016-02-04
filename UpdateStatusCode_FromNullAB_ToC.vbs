'***************************************************************************
'	Purpose - Update Status Code to "C" for all general records where value 
'				Status Code value is currently null, "A", or "B"
'
'	Input - Excel file: (\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\UpdateStatusCode_FromNullAB_ToC_Data.xls)
'
'	Output - Update ZSD_CONS_OA
'
'	Variables - WERKS - Plant
'				LGORT - Storage Location
'				MATNR - Material number
'				ZZSSCODE - Stocking Status Code (aka. ABC Policy Code)
'				ZZCOMMENT - Comment
'
'	Revision		Date		Description
'
'***************************************************************************

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim CurrDate,NewComment,OldComment,OldComment2,CurrComment
Dim PolCode,NumOfRows,CommentLen,FullComment,myComment,StatusText

Set ExcelApp = CreateObject("Excel.Application")
Set ExcelWorkbook = ExcelApp.Workbooks.Open("\\WinFile02\data\CustSvc\Parts\Inventory\PMx Scripts\UpdateStatusCode_FromNullAB_ToC_Data.xls")
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

CurrDate = Date
OldComment = "MIGRATE ARPIL 2012"
OldComment2 = "MIGRATE APRIL 2012"

'User is prompted to enter first row of Excel spreadsheet to be read
Row=InputBox("Row to start at")
'Counter = 0
Counter = 70 'just for testing purposes

Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").text = "/nzsd_cons_oa"
session.findById("wnd[0]").sendVKey 0

' Do Until Loop will execute until the last row (a blank row) is found
Do Until ExcelSheet.Cells(Row,1).Value = ""
	'If ExcelSheet.cells(row,11).value <> "error" Then    'bypass records in flatfile not found in ZSD_CONS_OA table.
		'If ExcelSheet.cells(row,6).value = "C" Then
			Session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").text = ExcelSheet.Cells(Row,1).Value  'Plant
			Session.findById("wnd[0]/usr/ctxtS_LGORT-LOW").text = ExcelSheet.Cells(Row,2).Value  'Storage loc
			Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").text = ExcelSheet.Cells(Row,3).Value  'Material number
			Session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").setFocus
			session.findById("wnd[0]/usr/ctxtS_MATNR-LOW").caretPosition = 11
			session.findById("wnd[0]").sendVKey 0
			session.findById("wnd[0]/tbar[1]/btn[8]").press
			NumOfRows = Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").rowcount
			
			Do Until Counter = (NumofRows)
					PolCode = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZSSCODE")
					MsgBox("numofrows = " & numofrows & ". PolCode = " & PolCode & ". Counter = " & Counter)
					Select Case PolCode 
						Case ""
							CurrComment = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZCOMMENT")
							myComment = fnCreateComment(CurrComment,PolCode)
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifycell Counter,"ZZSSCODE","C"
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified	
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell Counter,"ZZCOMMENT",myComment
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
							Counter = Counter + 1
							
						Case "A"
							CurrComment = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZCOMMENT")
							myComment = fnCreateComment(CurrComment,PolCode)
							session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifycell Counter,"ZZSSCODE","C"
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell Counter,"ZZCOMMENT",myComment
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
							Counter = Counter + 1
							
						Case "B"
							CurrComment = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").GetCellValue(Counter,"ZZCOMMENT")
							myComment = fnCreateComment(CurrComment,PolCode)
							session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifycell Counter,"ZZSSCODE","C"
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").modifyCell Counter,"ZZCOMMENT",myComment
							Session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").triggerModified
							Counter = Counter + 1
							
						Case Else
							Counter = Counter + 1
					End Select
		
			Loop

			'Session.findById("wnd[0]/tbar[0]/btn[11]").press
			'Session.findById("wnd[1]/usr/btnBUTTON_1").press
			'Session.findById("wnd[0]/tbar[0]/btn[3]").press
			'StatusText = Session.findById("wnd[0]/sbar").text
			'ExcelSheet.Cells(Row,7).Value = StatusText

		'End If
	'End If    'end if statement. DST 1-31-13
	row=row+1
	Counter = 0
Loop 

'************************************************************************
'	Function:	fnCreateComment
'	Purpose:	Generate apporopriate comment for each record that is updated
'	Parameters:	CurrComment and PolCode
'	Result: 	Comment to be saved in the ZSD_CONS_OA table
'************************************************************************
Function fnCreateComment(CurrComment,PolCode)

	Select Case PolCode
		Case ""
			If (CurrComment = "") Then 
				fnCreateComment = "Status changed from null to C " & CurrDate
			ElseIf CurrComment = OldComment Then
				fnCreateComment = "Status changed from null to C " & CurrDate
			ElseIf CurrComment = OldComment2 Then
				fnCreateComment = "Status changed from null to C " & CurrDate
			Else
				FullComment = "Status changed from null to C " & CurrDate & "; " & CurrComment
				If Len(FullComment) > 80 then
					fnCreateComment = Left(FullComment,80)
				Else
					fnCreateComment = FullComment
				End If
			End If
		Case "A"
			If (CurrComment = "") Then 
				fnCreateComment = "Status changed from A to C " & CurrDate
			ElseIf CurrComment = OldComment Then
				fnCreateComment = "Status changed from A to C " & CurrDate
			ElseIf CurrComment = OldComment2 Then
				fnCreateComment = "Status changed from A to C " & CurrDate
			Else
				FullComment = "Status changed from A to C " & CurrDate & "; " & CurrComment
				If Len(FullComment) > 80 then
					fnCreateComment = Left(FullComment,80)
				Else
					fnCreateComment = FullComment
				End If
			End If
		Case "B"
			If (CurrComment = "") Then 
				fnCreateComment = "Status changed from B to C " & CurrDate
			ElseIf CurrComment = OldComment Then
				fnCreateComment = "Status changed from B to C " & CurrDate
			ElseIf CurrComment = OldComment2 Then
				fnCreateComment = "Status changed from B to C " & CurrDate
			Else
				FullComment = "Status changed from B to C " & CurrDate & "; " & CurrComment
				If Len(FullComment) > 80 then
					fnCreateComment = Left(FullComment,80)
				Else
					fnCreateComment = FullComment
				End If
			End If
	End Select 
	
End Function


ExcelApp.Quit

Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
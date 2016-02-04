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
Dim PMxReadRow,ItmNum,ItmCat,ItmNumFlag,EDIField

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

strDir = "D:\Documents and Settings\dma02\Desktop\"
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
	ExcelSheet.cells(row,22).value= infoResult
	If objdictionary.Exists(test3) Then
	  EDIResult="EDI Vendor"
	 Else
	 	EDIResult="Non-EDI Vendor"
	 End If
	 ExcelSheet.cells(Row,25)=EDIResult    
End Sub
Row=1
Do Until excelsheet.cells(row+1,1)=""
	Row = row+1
	Call transaction_type
Loop
ExcelSheet.cells(1,22).value="Shipment Type"
ExcelSheet.cells(1,23).value="Subcontract Check"
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
	
	EDIField=ExcelSheet.cells(Row,25)
	If EDIField = "EDI Vendor" Then
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
		If ExcelSheet.cells(Row,22)="Goods Receipt" then
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
						ExcelSheet.cells(Row,23).value="Inbound"
					End If
					If ItmCat<>"L" Then
						ExcelSheet.cells(Row,23).value="not Subcontracted"
					End If
				Row=Row+1
			End If
			If ItmNum <> test2 Then
				PMxReadRow=PMxReadRow+1
			End If
			Loop
			
		End if
		If ExcelSheet.cells(Row,22)="Inbound" Then
			ExcelSheet.cells(Row,23).value="checked"
			Row=Row+1
		End If
		
	Loop
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
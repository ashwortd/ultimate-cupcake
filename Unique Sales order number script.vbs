Dim ExcelApp,ExcelWorkbook,ExcelSheet
Dim SONumber,Row,SOCount,i
Const xlUp = -4162

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
Set ExcelSheet = ExcelWorkbook.Worksheets("Sheet1")
Row=1

With ExcelSheet
	SOCount= .range("A"& .rows.count).end(xlUp).row
End With

ReDim SONumber(SOCount-1)
i=0
For Each x In SONumber
	SONumber(i) = ExcelSheet.Cells(Row,1).value
	Row=Row+1
	i=i+1
Next
Set objDictionary=CreateObject("Scripting.Dictionary")
For Each r In SONumber
	If Not objdictionary.Exists(r) Then
		objdictionary.Add r,r
	End If
Next
intItems = objDictionary.Count -1

ReDim SONumber(intItems)
i=0
For Each strKey In objdictionary.Keys
	SONumber(i)= strKey
	i=i+1
Next

i=0
Set fso = CreateObject("Scripting.FileSystemObject")

logfile  = "D:\Documents and Settings\dma02\Desktop\Sample.log"
Set f = fso.OpenTextFile(logfile, 2, True)
For Each x In SOnumber
	message  = "Sales Order Number " & SONumber(i)
	f.WriteLine "[" & Date & " - " & Time & "] " & message 
	i=i+1
Next

f.Close
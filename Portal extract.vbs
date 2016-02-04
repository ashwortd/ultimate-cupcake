Dim objShell
Set objShell = CreateObject("Wscript.Shell")
Const olFolderInbox = 6
Const xlCellTypeLastCell = 11
Const xlUp = -4162
Dim colFilteredItems, sToDay, colItems, strPath, lastRow, row
strPath = objShell.CurrentDirectory
dtmYester = Now()
dtmOlder = Now() - 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objOutlook = CreateObject("Outlook.Application")
Set objNamespace = objOutlook.GetNamespace("MAPI")
Set objFolderIN = objNamespace.GetDefaultFolder(olFolderInbox)
Set objFolder = objFolderIN.Folders("Freight Portal")
Dim arrFiles()
intSize = 0
Set colItems = objFolder.Items
Set colFilteredItems = colItems.Restrict("[UnRead] = True")
Set colFilteredItems = colItems.Restrict("[From] = 'Alstom Freight Portal'")
For Each objMessage In colFilteredItems
    Set colAttachments = objMessage.Attachments 
    intCount = colAttachments.Count
    If intCount <> 0 Then
        For i = 1 To intCount
            strFileName = "D:\Documents and Settings\256858\Desktop\Email Attachments\" & objMessage.Attachments.Item(i).FileName
            objMessage.Attachments.Item(i).SaveAsFile strFileName
            ReDim Preserve arrFiles(intSize)
            arrFiles(intSize) = strFileName
            intSize = intSize + 1
        Next
    End If
Next
Set objExcel = WScript.CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.DisplayAlerts = False
'Set wb2 = objExcel.Workbooks.Add
Set wb2 = objExcel.Workbooks.Open(strPath&"\Time Booking.xlsm")
Set ws2 = wb2.Worksheets("Shipping Notes")
For Each strFile in arrFiles
    Set wb = objExcel.Workbooks.Open(strFile)
    Set ws = wb.Worksheets(1)
    Set objRange = ws.UsedRange.Offset(1, 0)
    objRange.Copy
    ws2.Activate
    ws2.Range("A1").Activate
    If ws2.Cells(1,1).Value <> "" Then
        Set objRange2 = ws2.UsedRange
        objRange2.SpecialCells(xlCellTypeLastCell).Activate
        intNewRow = objExcel.ActiveCell.Row + 1
        strNewCell = "A" &  intNewRow
        ws2.Range(strNewCell).Activate
    End If
    ws2.Paste
    wb.Close
    
    objFSO.DeleteFile(strFile)
Next

For i = colItems.Count to 1 Step - 1
    colItems(i).Delete
Next

    lastRow = ws2.range("A" & ws2.Rows.Count).End(xlUp).Row
    row = ws2.range("O" & ws2.Rows.Count).End(xlUp).Row + 1
    
    Do While row <= lastRow
    
        ws2.Cells(row,15).Formula = "=E"&row&"&"&Chr(34)&"/"&Chr(34)&"&"&"F"&row
        ws2.Cells(row,16).Formula = "=M"&row
        row = row + 1
    Loop
Set objShell = nothing


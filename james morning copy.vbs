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
session.findById("wnd[0]").maximize
'19032-256-01-RBLD
'Material Description
'Earliest Requirements Date
'Number of Latest Receipt Element
'Latest Receipt Date

Dim objShell, ex, wb, ws, ws2, wsOH
Dim i, n, row, trow, lastRowt, status
Dim tierProds, t, lastRow, lastRowOH
Dim failCase, initWidth, compWidth
Dim currentTier, newTier
Const xlCalculationManual = -4135
Const xlCalculationAutomatic = -4105
Const xlUp = -4162
Const x = 16

Set objShell = CreateObject("Wscript.Shell")
strPath = objShell.CurrentDirectory
failCase = "Unmodified"


Call Main

Sub Main

    MsgBox("Be sure to move PE1 (1) to Primary Monitor")
    
    Set ex = GetObject( , "Excel.Application")
    ex.Visible = True
    Set wb = ex.Workbooks("Final Report.xlsm")
    
    Set ws = wb.Sheets("MD04")
    Set ws2 = wb.Sheets("Transfers Ship")
    Set wsOH = wb.Sheets("Order Headers")
    
    Call R_Folders
    
    
'   ws.select
    ws.Cells(1,1).CurrentRegion.Offset(1,0).ClearContents
'   ws.Cells.Select
    ws.Cells.NumberFormat = "@"
'   ws.Cells(1,1).select
    
'   ws2.Cells(2,12).Formula = "=IF(F2<>" & Chr(34) & Chr(34) & ",VLOOKUP(F2&" & Chr(34) & "/" & Chr(34) & "&G2&" & Chr(34) & "/" & Chr(34) & "&D2,'Order Headers'!K:N,4,FALSE),IF(COUNTIF('Order Headers'!F:F,D2)=1,VLOOKUP(D2,'Order Headers'!F:N,9,FALSE),IF(COUNTIF('Order Headers'!L:L,D2&" & Chr(34) & "/" & Chr(34) & "&K2)=1,VLOOKUP(D2&" & Chr(34) & "/" & Chr(34) & "&K2,'Order Headers'!L:N,3,FALSE),VLOOKUP(F2&" & Chr(34) & "/" & Chr(34) & "&G2&" & Chr(34) & "/" & Chr(34) & "&D2,'Order Headers'!K:N,4,FALSE))))"
'   ws2.Range("L2").autofill ws.Range("L2:L"&lastRow)
    
'   ex.ScreenUpdating = False
    ex.Calculation = xlCalculationManual
'   ex.EnableEvents = False
'   ex.DisplayAlerts = False
    
    
    
    lastRowt = ws2.range("E" & ws2.Rows.Count).End(xlUp).Row
'   MsgBox lastRow
    trow = 2
    row = 2
    For Each cell In ws2.Range("E2:E"&lastRowt)
        Do
            Call DefineTopLevel
'           Call MakeDictionary
'           Call OpenPrintPreview(cell)
'           Call ExtractHierarchy
                If  status <> "Shipped" Then
            Call MakeDictionary
            Call ExtractHierarchy
                Else
                trow = trow + 1
                Exit Do
                End If        
            If tierProds.Item("1") = "" Then
            ws2.Cells(trow,12).Value = "In Inventory"
            ws2.Cells(trow,14).Value = "In Inventory"
            ElseIf Len(tierProds.Item("1")) = 10 Then
            ws2.Cells(trow,12).Value = tierProds.Item("1")
            ws2.Cells(trow,14).Value = "Purchased"
            Else
            ws2.Cells(trow,12).Value = tierProds.Item("1")
            End If
            trow = trow + 1
            Set tierProds = Nothing
'       MsgBox row & " out of " & lastRow
        Loop While False    
    Next
        
    ex.Visible= True
    ex.ScreenUpdating = True
    ex.Calculation = xlCalculationAutomatic
    ex.EnableEvents = True
    ex.DisplayAlerts = True
    
'   lastRow = ws.range("H" & ws.Rows.Count).End(xlUp).Row
'   wsOH.select
'   lastRowOH = wsOH.range("C" & wsOH.Rows.Count).End(xlUp).Row
'   wsOH.Range("J2:J" & lastRowOH).Copy
'   ws.Select
'   ws.Cells(lastRow+1,1).Select
'   ws.Paste
'   ex.CutCopyMode = False
'   wsOH.select
'   wsOH.Range("C2:C" & lastRowOH).EntireColumn.Copy
'   ws.Select
'   ws.Cells(lastRow+1,5).Select
'   ws.Paste
'   ex.CutCopyMode = False
    
'   ws.Cells(1,1).AutoFilter 1,"="
'   ws.Cells(1,1).CurrentRegion.Offset(1,0).Delete -4162
    ws.Cells(1,1).AutoFilter
    ws.Cells(1,1).AutoFilter 5,"="
    ws.Cells(1,1).CurrentRegion.Offset(1,0).Delete -4162
    ws.Cells(1,1).AutoFilter
    
    lastRow = ws.range("H" & ws.Rows.Count).End(xlUp).Row
'   ws.Range("A2:H" & lastRow).Sort ws.Range("A2:A" & lastRow), 1
    
    ws.Columns("I").NumberFormat = "General"
    ws.Cells(2,9).Formula = "=VLOOKUP(A2&"&Chr(34)&"/"&Chr(34)&"&B2,Components!D:E,2,FALSE)"
    ws.Range("I2").autofill ws.Range("I2:I"&lastRow)
    
    ex.Columns.AutoFit
    wb.Save
    
    WScript.ConnectObject session,     "off"
    WScript.ConnectObject application, "off"
    MsgBox("Done")
    WScript.Quit
End Sub



Sub R_Folders

'   wsOH.select
    
    lastRow = wsOH.range("E" & wsOH.Rows.Count).End(xlUp).Row
    wsOH.Range("A2:N" & lastRow).Sort wsOH.Range("E2:E" & lastRow), 1
    
    ohRow = 2
    ohFlag = 0
    Do While wsOH.Cells(ohRow, 5).Value <> ""
        If wsOH.Cells(ohRow, 5).Value = "ZNCR" Then
            ohFlag = 1
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nCO03"
            session.findById("wnd[0]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/usr/ctxtCAUFVD-AUFNR").Text = wsOH.Cells(ohRow, 3).Value
            session.findById("wnd[0]/tbar[0]/btn[0]").press
            session.findById("wnd[0]/mbar/menu[4]/menu[3]").select
            wsOH.Cells(ohRow, 10).Value = session.findById("wnd[0]/usr/tblSAPLKOBSTC_RULES/ctxtDKOBR-EMPGE[1,0]").Text
        ElseIf wsOH.Cells(ohRow, 5).Value <> "ZNCR" And ohFlag = 1 Then
            Exit Do
        End If
        ohRow = ohRow + 1
    Loop
    wsOH.Range("A2:N" & lastRow).Sort wsOH.Range("I2:I" & lastRow), 1
End Sub



Sub OpenPrintPreview(cell)

    session.findById("wnd[0]/tbar[0]/okcd").text = "/nco46"
    session.findById("wnd[0]/tbar[0]/btn[0]").press
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB40").select
    session.findById("wnd[0]/usr/tabsTABSTRIP_T1/tabpTAB40/ssub%_SUBSCREEN_T1:PP_ORDER_PROGRESS:0040/ctxtP_AUFNR").text = cell.Value
    session.findById("wnd[0]/tbar[1]/btn[8]").press
'   session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton "&LOAD"
'   session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem "&LOAD"
'   session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").currentCellRow = 14
'   session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").selectedRows = "14"
'   session.findById("wnd[1]/usr/cntlGRID/shellcont/shell").clickCurrentCell
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton "&PRINT_BACK"
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem "&PRINT_PREV_ALL"


End Sub



Sub ExtractHierarchy
    t = 1
    session.findById("wnd[0]/usr").VerticalScrollBar.Position = 0
    initWidth = session.findById("wnd[0]/usr/lbl[8,3]").CharWidth
    compWidth = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & ",5]").CharWidth
'   ws.Cells(2,1).Value = session.findById("wnd[0]/usr/lbl[8,3]").Text
'   ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[12,5]").Text
'   ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & ",5]").Text
'   ws.Cells(row,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 50 & ",5]").Text
'   ws.Cells(row,5).Value = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + 68 & ",5]").Text)
    tierProds.Item(CStr(t)) = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 28 & ",5]").Text)
'   On Error Resume Next
'   ws.Cells(row,6).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 88 & ",5]").Text
'   If Err.Number <> 0 Then
'       Err.Clear
'   End If
'   On Error Goto 0
    
    n = 0
    i = 7
    
    
    Do While failCase <> "Done"
    
        On Error Resume Next
        ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
        
        
        If Err.Number <> 0 Then
            Err.Clear
            Call FailCheck
            Select Case failCase
                Case "Shift Left"
                    i = i + 1
                    n = n - (4 * (currentTier - newTier))
                    'row = row + 1
                    t = newTier
                    'MsgBox("Shifting Left.")
                    ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
                    ws.Cells(row,1).Value = tierProds.Item(CStr(t))
                Case "Shift Right"
                    i = i + 1
                    n = n + 4
                    'row = row + 1
                    t = CInt(t) + 1
                    tierProds.Item(CStr(t)) = ws.Cells(row - 1,5).Value
                    'MsgBox("Shifting Right.")
                    ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text
                    ws.Cells(row,1).Value = tierProds.Item(CStr(t))
                Case "Done"
'                   MsgBox("Exiting via end of hierarchy")
                    'row = row + 1
                    Exit Do
            End Select            
        Else
            ws.Cells(row,1).Value = tierProds.Item(CStr(t))
        End If
        On Error Goto 0
        
        
        ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & "," & i & "]").Text
        
        ws.Cells(row,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 10 & "," & i & "]").Text
        
        ws.Cells(row,5).Value = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 28 & "," & i & "]").Text)
        
        On Error Resume Next
        ws.Cells(row,6).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 48 & "," & i & "]").Text
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error Goto 0
        
        ws.Cells(row,7).Value = CInt(t)
        
        ws.Cells(row,8).Value = tierProds.Item("1")
        
        
        If i = 48 Then
            session.findById("wnd[0]/usr").VerticalScrollBar.Position = session.findById("wnd[0]/usr").VerticalScrollBar.Position + 42
            Call ScrollCheck
        ElseIf i = 49 Then
            session.findById("wnd[0]/usr").VerticalScrollBar.Position = session.findById("wnd[0]/usr").VerticalScrollBar.Position + 43
            Call ScrollCheck        
        Else
            i = i + 1
        End If
        row = row + 1
    Loop
'   trow = trow + 1
    failCase = ""
    
End Sub



Sub MakeDictionary
    Set tierProds = CreateObject("Scripting.Dictionary")
    tierProds.RemoveAll
    tierProds.Add "1",""
    tierProds.Add "2",""
    tierProds.Add "3",""
    tierProds.Add "4",""
    tierProds.Add "5",""
    tierProds.Add "6",""
    tierProds.Add "7",""
    tierProds.Add "8",""
    tierProds.Add "9",""
    tierProds.Add "10",""
    tierProds.Add "11",""
    tierProds.Add "12",""
    tierProds.Add "13",""
    tierProds.Add "14",""
    tierProds.Add "15",""
    tierProds.Add "16",""
    tierProds.Add "17",""
    tierProds.Add "18",""
    tierProds.Add "19",""
    tierProds.Add "20",""
    tierProds.Add "21",""
    tierProds.Add "22",""
    tierProds.Add "23",""
    tierProds.Add "24",""
    tierProds.Add "25",""
    tierProds.Add "26",""
    tierProds.Add "27",""
    tierProds.Add "28",""
    tierProds.Add "29",""
    tierProds.Add "30",""
End Sub



Sub FailCheck
    On Error Resume Next
    session.findById("wnd[0]/usr/lbl[" & (x + n) + 4 & "," & i + 1 & "]").SetFocus
    
    If Err.Number <> 0 Then
        Err.Clear
        currentTier = (n / 4) + 1
        For num = 1 To currentTier
            session.findById("wnd[0]/usr/lbl[" & (x + n) - (4 * num) & "," & i + 1 & "]").SetFocus
            If Err.Number <> 0 Then
                Err.Clear
                failCase = "Done"
            Else
                failCase = "Shift Left"
                newTier = currentTier - num
                Exit For
            End If    
        Next
    Else
        failCase = "Shift Right"
    End If
    On Error Goto 0
End Sub



Sub ScrollCheck
    On Error Resume Next
    session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").SetFocus
    
    If Err.Number <> 0 Then
        Err.Clear
        i = 7
    Else
        If ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i & "]").Text And ws.Cells(row,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 10 & "," & i + 1 & "]").Text  Then
            session.findById("wnd[0]/usr/lbl[" & x + n & "," & i + 1 & "]").SetFocus
            If Err.Number <> 0 Then
                Err.Clear
                failCase = "Done"
            Else
                ws.Cells(row,2).Value = session.findById("wnd[0]/usr/lbl[" & x + n & "," & i + 1 & "]").Text
                ws.Cells(row,1).Value = tierProds.Item(CStr(t))
                ws.Cells(row,3).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + 9 & "," & i + 1 & "]").Text
                ws.Cells(row,4).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 10 & "," & i + 1 & "]").Text        
                ws.Cells(row,5).Value = Trim(session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 28 & "," & i + 1 & "]").Text)
                ws.Cells(row,6).Value = session.findById("wnd[0]/usr/lbl[" & initWidth + compWidth + 48 & "," & i + 1 & "]").Text
                ws.Cells(row,7).Value = CInt(t)
                ws.Cells(row,8).Value = tierProds.Item("1")
                failCase = "Done"
            End If
        Else
        i = 7
        End If
    End If
    On Error Goto 0
End Sub

Sub DefineTopLevel    

'part = cell.Value
val = 1

order = ws2.Cells(trow,2).Value & "/" & Left("00000",5 - Len(ws2.Cells(trow,3).Value)) & ws2.Cells(trow,3).Value

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmd04"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").text = ws2.Cells(trow,5).Value
session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").text = "500E"
session.findById("wnd[0]").sendVKey 0
vrc = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").visiblerowcount
Do
status = "none"
On Error Resume Next
compare = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5,"&val&"]").text    
If compare = order Then
'   MsgBox session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5,"&val&"]").text
'   Exit Do
    status = "success"
ElseIf val >= vrc Then
 ws2.Cells(trow,12).Value = "Shipped"
 ws2.Cells(trow,14).Value = "Shipped"
 Err.Clear
    status = "Shipped"
 Exit Sub
Else
val = val + 1
End If
Loop While status = "none"
status = "success"
'Select Cell & Get Order Report
session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5,"&val&"]").setFocus
session.findById("wnd[0]").sendVKey 38

'Open Print Preview
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").pressContextButton "&PRINT_BACK"
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell[0]").selectContextMenuItem "&PRINT_PREV_ALL"



End Sub

'=IF(F2<>" & Chr(34) & Chr(34) & ",VLOOKUP(F2&" & Chr(34) & "/" & Chr(34) & "&G2&" & Chr(34) & "/" & Chr(34) & "&D2,'Order Headers'!K:N,4,FALSE),IF(COUNTIF('Order Headers'!F:F,D2)=1,VLOOKUP(D2,'Order Headers'!F:N,9,FALSE),IF(COUNTIF('Order Headers'!L:L,D2&" & Chr(34) & "/" & Chr(34) & "&K2)=1,VLOOKUP(D2&" & Chr(34) & "/" & Chr(34) & "&K2,'Order Headers'!L:N,3,FALSE),VLOOKUP(D2,'Order Headers'!F:N,9,FALSE))))

'=IF(F2<>" & Chr(34) & Chr(34) & ",VLOOKUP(F2&" & Chr(34) & "/" & Chr(34) & "&G2&" & Chr(34) & "/" & Chr(34) & "&D2,'Order Headers'!K:N,4,FALSE),IF(COUNTIF('Order Headers'!F:F,D2)=1,VLOOKUP(D2,'Order Headers'!F:N,9,FALSE),IF(COUNTIF('Order Headers'!L:L,D2&" & Chr(34) & "/" & Chr(34) & "&K2)=1,VLOOKUP(D2&" & Chr(34) & "/" & Chr(34) & "&K2,'Order Headers'!L:N,3,FALSE),VLOOKUP(D2,'Order Headers'!F:N,9,FALSE))))


'Transfer Orders Rev Oct28 _ 3210766.xlsx


'/app/con[0]/ses[0]/wnd[0]/usr/lbl[24,48]


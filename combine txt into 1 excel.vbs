Sub ListAllFile3()

 Application.DisplayAlerts = False

Dim strFile As String
strFile = Dir("C:\Users\dma02\Desktop\changes\*.txt", vbNormal)
MsgBox (strFile)
Do While Len(strFile) > 0
Set ws = Worksheets.Add
ws.Name = strFile

With Worksheets(strFile).QueryTables.Add(Connection:="TEXT;C:\Users\dma02\Desktop\changes\" & strFile, Destination:=Range("$A$1"))
        .FieldNames = True
        .RowNumbers = False

        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFileStartRow = 5
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = True
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(xlSkipColumn, xlTextFormat, xlSkipColumn, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False

    End With

strFile = Dir
Loop

 Application.DisplayAlerts = True

End Sub
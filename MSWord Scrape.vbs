Dim tblMaterial
strDocPath="c:\temp\EBRTest.docx"
Set objWord = CreateObject("Word.Application")
objWord.Visible = True
Set objDoc = objWord.Documents.Open(strDocPath)

Set objRange = objDoc.Range()
'strFilePath = "c:\temp\my_future_table.txt"

'Set objFSO = CreateObject("scripting.filesystemobject")
'Set objTF = objFSO.opentextfile(strFilePath)
'strAll = objTF.readall
'arrVar = Split(strAll, vbNewLine)
'numcols = UBound(Split(arrVar(0), vbTab)) + 1

'objDoc.Tables.Add objRange, UBound(arrVar) - LBound(arrVar) + 1, numcols
For i =1 To 10
Set objTable = objDoc.Tables(i)
Dim strTest
	strTest=objtable.Cell(1,1).Range.Text
	MsgBox(strTest)
	If Left(strTest,9)="Line Item" Then
		tblMaterial=i
		MsgBox("Found Material Table")
		Exit For
	End If
	
'For lngrow = LBound(arrVar) To UBound(arrVar)
'    arrVar2 = Split(arrVar(lngrow), vbTab)
'    For lngcol = LBound(arrVar2) To UBound(arrVar2)
'     objTable.Cell(lngrow + 1, lngcol + 1).Range.Text = arrVar2(lngcol)
'	 objTable.Ce    
'    Next
'Next

'objTF.Close
'set objFSO = Nothing

'objTable.AutoFormat (9)
Next

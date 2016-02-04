'==== Use this sample file with Samples\BAPI\BAPI_Download Customer Master.iff ====
'==== This is a Pre-Run sample
'==== This script will work only with external excel file

Dim wBook		'==== Excel Work Book
Dim wSheet		'==== Excel Work Sheet

Dim i
Dim CustNos(10)		'==== Store sample data
Dim xlRow		'==== Write start row in worksheet

Set wBook = GetObject(#CURXLFILE#)		'==== Get Workbook from iBook/External excel file
Set wSheet = wBook.Worksheets(#CURXLSHEET#)	'==== Get selecetd Worksheet from Workbook

'====================================================
'==== Create Sample data to write in opened excel worksheet
'==== You can get data from any text file or database table or any source 
'==== Here in the example, we have used array

CustNos(0) = 515
CustNos(1) = 517
CustNos(2) = 520
CustNos(3) = 521
CustNos(4) = 523
CustNos(5) = 525
CustNos(6) = 538
CustNos(7) = 539
CustNos(8) = 540
CustNos(9) = 543

'====================================================

'==== Write data to excel ====

xlRow = 32				'=== Set starting excel row

For i = 0 To UBound(CustNos) - 1
	wSheet.Cells(xlRow, "A") = CustNos(i)
	xlRow = xlRow + 1
Next

'=============================

'==== Clean every object

wSheet = null
wBook = null

'====

WScript.Quit(0)   ' ==== ( No Error State - Success)

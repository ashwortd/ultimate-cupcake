Dim curWbook
Dim curWsheet

Dim ExtractToPath
Dim ExtractToFile

Dim objExcel
Dim wBook
Dim fpath

ExtractToPath ="C:\Users\TestUser\Desktop\DemoXML\"  'Path where XML is stored
ExtractToFile =  "DemoXMLFile.XML" 'Name of XML file to convert

fpath = ExtractToPath & ExtractToFile 

set curWbook= GetObject(#CURXLFILE#) 'The currrent external excel file
Set curWsheet = curWbook.Worksheets(#CURXLSHEET#) 'Curent worksheet of current external excel file

Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False

objExcel.Workbooks.OpenXML fpath, , 2 'Open the XML file as list using load option value = 2

Set wBook =  objExcel.WorkBooks(1)
wBook.Application.Visible = true

objExcel.ActiveSheet.Cells.Select 'Select all cells from XML workshet
objExcel.Selection.Copy 'Copy selection to clipboard

curWbook.Windows(curWbook.Name).Visible = True  'Make current excel visible
curWbook.Activate 
curWbook.Worksheets(#CURXLSHEET#).Activate 'Make selected worksheet current active worksheet

curWsheet.Cells("1","A").Select 'Goto first cell 
curWsheet.Paste 'Paste clipboard content

wBook.Close 'close temp workbook

objExcel.DisplayAlerts = true

objExcel.Quit 'Close the temp object

Set objExcel = Nothing
Set wBook = Nothing

curWbook.Application.Visible = true 'make current workbook visible
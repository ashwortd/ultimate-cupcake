Dim ex, wb, ws

	
	'******************
	'Steps 1-3*********
	'******************
	Set fso = WScript.CreateObject("Scripting.FileSystemObject") 
	Set ex = WScript.CreateObject("Excel.application")
	ex.Visible = True
	Set wb = ex.Workbooks.Open("D:\Documents and Settings\267041\Desktop\500E Open Jobs 6_21_2013.xls")
	Set ws = wb.Sheets(wb.ActiveSheet.Name)
	
	
	'******************
	'Step 4************
	'******************
	ws.Columns("A:A").Select
	ex.Selection.Delete -4131
	ws.Rows("7:7").Select
    ex.Selection.Delete -4162
    ws.Rows("1:5").Select
    ex.Selection.Delete -4162
	
	
	
	'******************
	'Step 5-6**********
	'******************
	ws.Cells(1,1).CurrentRegion.Select
	ex.Selection.AutoFilter
	
	ex.Selection.AutoFilter 5,""
	
	ws.Cells(1,1).CurrentRegion.Offset(1,0).Select
	ex.Selection.Delete -4162
	
	ws.Cells(1,1).CurrentRegion.Select
	ex.Selection.AutoFilter
	
	ws.Cells(1,1).Select
	
	
	
	'WScript.Sleep(100000)
		
	
	'******************
	'Step 7-9**********
	'******************
	ws.Cells(1,1).CurrentRegion.Select
	ex.Selection.AutoFilter
	
	ex.Selection.AutoFilter 19,"=*CNF*", 1, "<>*PCNF*"
	
	ws.Cells(1,1).CurrentRegion.Offset(1,0).Select
	ex.Selection.Delete -4162
	
	ws.Cells(1,1).CurrentRegion.Select
	ex.Selection.AutoFilter
	
	ws.Cells(1,1).Select
	





	ws.Columns.AutoFit
	
	
	WScript.Sleep(600000)
	'wb.SaveAs(newFilePath)
	'WScript.Sleep(100)
	'fso.DeleteFile(filePath)
	
	
    

	
	
	
	
	'ex.Save
	ex.Quit 
	
	Set fso = Nothing
	Set ex = Nothing
	Set wb = Nothing
	Set ws = Nothing
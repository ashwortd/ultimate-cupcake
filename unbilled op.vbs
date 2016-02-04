Dim objExcel,objWorkbook

Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open("D:\Documents and Settings\dma02\Desktop\Unbilled Shipment3.xlsm")
objExcel.Application.Visible = True
WScript.Sleep(60000)
objWorkbook.Close(True)
objExcel.SendKeys "+{TAB}"
objExcel.SendKeys "{ENTER}"
WScript.Quit
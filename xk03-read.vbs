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
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,POrg,Row
Set ExcelApp = CreateObject("Excel.Application")
POrg=InputBox("Which Purchasing Organization?(US44,US45,US49,US50)")

	If POrg="US44" Then
		Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\nsorrell\Desktop\Expeditor Load\RichmondVendor-ER-US44.xlsx")
	 ElseIf POrg="US45" Then
		Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\nsorrell\Desktop\Expeditor Load\RichmondVendor-ER-US45.xlsx")
	 ElseIf POrg="US49" Then
		Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\nsorrell\Desktop\Expeditor Load\RichmondVendor-ER-US49.xlsx")
	 Elseif POrg="US50" Then
		Set ExcelWorkbook = ExcelApp.Workbooks.Open("D:\Documents and Settings\nsorrell\Desktop\Expeditor Load\RichmondVendor-ER-US50.xlsx")
	 Else MsgBox("Purchase Organization Data File is not known")
		WScript.Quit
	End If
	
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

Row=InputBox("Row to start at")
session.findById("wnd[0]").maximize
Call ReadStart
Sub ReadStart
If ExcelSheet.Cells(Row,1).Value = "" Then
	MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.Quit
	End If
Session.findById("wnd[0]/tbar[0]/okcd").text = "/nxk03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/chkWRF02K-D0320").selected = true
session.findById("wnd[0]/usr/ctxtRF02K-LIFNR").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]/usr/ctxtRF02K-BUKRS").text = "5000"
session.findById("wnd[0]/usr/ctxtRF02K-EKORG").text = Porg
session.findById("wnd[0]/usr/chkWRF02K-D0320").setFocus
session.findById("wnd[0]").sendVKey 0
'session.findById("wnd[0]/tbar[0]/btn[3]").press
On Error Resume next
StrPARVW1=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,0]").text
	If Err.Number<>0 Then
		ExcelSheet.Cells(Row,9).Value = session.findById("wnd[0]/sbar").Text
		row=row+1
		Call ReadStart
		Err.Clear
	End if
On Error Goto 0	
StrPARVW2=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,1]").text
StrPARVW3=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,2]").text
StrPARVW4=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,3]").text
StrPARVW5=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,4]").text
StrPARVW6=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,5]").text
StrPARVW7=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,6]").text
StrPARVW8=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,7]").text
StrPARVW9=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,8]").text
StrPARVW10=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,9]").text
StrPARVW11=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWYT3-PARVW[0,10]").text
StrGPARN1=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,0]").text
StrGPARN2=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,1]").text
StrGPARN3=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,2]").text
StrGPARN4=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,3]").text
StrGPARN5=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,4]").text
StrGPARN6=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,5]").text
StrGPARN7=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,6]").text
StrGPARN8=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,7]").text
StrGPARN9=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,8]").text
StrGPARN10=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,9]").text
StrGPARN11=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/ctxtWRF02K-GPARN[2,10]").text
StrRNAME1=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,0]").text
StrRNAME2=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,1]").text
StrRNAME3=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,2]").text
StrRNAME4=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,3]").text
StrRNAME5=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,4]").text
StrRNAME6=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,5]").text
StrRNAME7=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,6]").text
StrRNAME8=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,7]").text
StrRNAME9=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,8]").text
StrRNAME10=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,9]").text
StrRNAME11=session.findById("wnd[0]/usr/tblSAPMF02KTCTRL_PARTNERROLLEN/txtWRF02K-REF_NAME1[3,10]").text
ExcelSheet.Cells(Row,9).Value=StrPARVW1
ExcelSheet.Cells(Row,10).Value=StrGPARN1
ExcelSheet.Cells(Row,11).Value=StrRNAME1
ExcelSheet.Cells(Row,12).Value=StrPARVW2
ExcelSheet.Cells(Row,13).Value=StrGPARN2
ExcelSheet.Cells(Row,14).Value=StrRNAME2
ExcelSheet.Cells(Row,15).Value=StrPARVW3
ExcelSheet.Cells(Row,16).Value=StrGPARN3
ExcelSheet.Cells(Row,17).Value=StrRNAME3
ExcelSheet.Cells(Row,18).Value=StrPARVW4
ExcelSheet.Cells(Row,19).Value=StrGPARN4
ExcelSheet.Cells(Row,20).Value=StrRNAME4
ExcelSheet.Cells(Row,21).Value=StrPARVW5
ExcelSheet.Cells(Row,22).Value=StrGPARN5
ExcelSheet.Cells(Row,23).Value=StrRNAME5
ExcelSheet.Cells(Row,24).Value=StrPARVW6
ExcelSheet.Cells(Row,25).Value=StrGPARN6
ExcelSheet.Cells(Row,26).Value=StrRNAME6
ExcelSheet.Cells(Row,27).Value=StrPARVW7
ExcelSheet.Cells(Row,28).Value=StrGPARN7
ExcelSheet.Cells(Row,29).Value=StrRNAME7
ExcelSheet.Cells(Row,30).Value=StrPARVW8
ExcelSheet.Cells(Row,31).Value=StrGPARN8
ExcelSheet.Cells(Row,32).Value=StrRNAME8
ExcelSheet.Cells(Row,33).Value=StrPARVW9
ExcelSheet.Cells(Row,34).Value=StrGPARN9
ExcelSheet.Cells(Row,35).Value=StrRNAME9
ExcelSheet.Cells(Row,36).Value=StrPARVW10
ExcelSheet.Cells(Row,37).Value=StrGPARN10
ExcelSheet.Cells(Row,38).Value=StrRNAME10
ExcelSheet.Cells(Row,39).Value=StrPARVW11
ExcelSheet.Cells(Row,40).Value=StrGPARN11
ExcelSheet.Cells(Row,41).Value=StrRNAME11
Row=row+1
Call ReadStart
End Sub

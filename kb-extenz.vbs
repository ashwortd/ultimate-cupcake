'*****************************************************
'	Thanks for choosing Ashworth's EZ-Script Service
' When it take to long to by hand Derek's your man
' Created for the sole use of KB 8/5/13
'*****************************************************


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
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row,i,z,x
Dim startingrow,wndstatus
Dim elementID, elementLeft, elementFinal
Dim Plant(41)
Plant(0)="5006"	
Plant(1)="5007"
Plant(2)="5009"
Plant(3)="500A"
Plant(4)="500B"
Plant(5)="500C"
Plant(6)="500D"
Plant(7)="500E"
Plant(8)="500F"
Plant(9)="500G"
Plant(10)="500H"
Plant(11)="500I"
Plant(12)="500J"
Plant(13)="500K"
Plant(14)="500L"
Plant(15)="50AA"
Plant(16)="5040"
Plant(17)="5050"
Plant(18)="5051"
Plant(19)="5060"
Plant(20)="5061"
Plant(21)="5062"
Plant(22)="5063"
Plant(23)="5064"
Plant(24)="5080"
Plant(25)="5090"
Plant(26)="5100"
Plant(27)="5110"
Plant(28)="5030"
Plant(29)="5070"
Plant(30)="5130"
Plant(31)="500M"
Plant(32)="500N"
Plant(33)="500P"
Plant(34)="500Q"
Plant(35)="500R"
Plant(36)="500T"
Plant(37)="500U"
Plant(38)="500V"
Plant(39)="5053"
Plant(40)="50AM"
Plant(41)="50BM"

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
Set ExcelWorkbook = ExcelApp.Workbooks.Open (objDialog.FileName)
Set ExcelSheet = ExcelWorkbook.Worksheets(1)

Session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").maximize
x=InputBox("Which plant to start (0-41)?","Reference Spreadsheet for Plant")
Row=InputBox("Row to start at")
startingrow=row
For i = x To 41
Row=startingrow
Do While ExcelSheet.Cells(Row,3).Value <> ""
Call partextend
Loop 
Next
If ExcelSheet.Cells(Row,1).Value=("0") Then
		'Call endscript
		MsgBox("The end has come")
		ExcelApp.Quit
		Set ExcelApp=Nothing
		Set ExcelWoorkbook=Nothing
		Set ExcelSheet=Nothing
		WScript.ConnectObject session,     "off"
   		WScript.ConnectObject application, "off"
		WScript.Quit
	End If

Sub partextend
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = " "
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = " "
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = ExcelSheet.Cells(Row,3).Value
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").setFocus
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").caretPosition = 14
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
If session.findById("wnd[0]/sbar").Text="Enter a material type" Then
	ExcelSheet.Cells(Row,(i+20)).Value ="Part not extended"
	Session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
	session.findById("wnd[0]").sendVKey 0
	Row=Row+1
	Exit Sub
End if	
Session.findbyid("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = false
Session.findById("wnd[1]/tbar[0]/btn[0]").press
On Error Resume next
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = Plant(i)
Session.findbyid("wnd[1]/usr/ctxtRMMG1-LGORT").text = " "
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = " "
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = " "
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1-PROPR").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = ExcelSheet.Cells(Row,1).Value
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-BWTAR").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGNUM").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGTYP").text=" "
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").caretPosition = 4
session.findById("wnd[1]").sendVKey 0
'On Error Resume Next
wndstatus=session.findById("wnd[2]/usr/txtMESSTXT1").text
	If wndstatus="Material already maintained for this" Then
		Session.findById("wnd[2]/tbar[0]/btn[0]").press
		Session.findById("wnd[1]/tbar[0]/btn[12]").press
		ExcelSheet.Cells(Row,(i+20)).Value ="Already Maintained"
		Row=Row+1
		wndstatus=0
	Exit sub
End If
On Error Goto 0
elementid = Session.ActiveWindow.GuiFocus.ID
elementLeft = Left(elementID, 50)
elementFinal = Right(elementLeft, 8)
'MsgBox(elementFinal)
If elementFinal="tabpSP24" Then
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	Session.findById("wnd[0]").sendVKey 0
	session.findById("wnd[0]").sendVKey 0
	ExcelSheet.Cells(Row,(i+20)).Value ="Completed"
	Row=Row+1
	Exit Sub
ElseIf elementFinal="tabpSP26" Then
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	Session.findById("wnd[0]").sendVKey 0
	ExcelSheet.Cells(Row,(i+20)).Value ="Completed"
	Row=Row+1
	Exit Sub
ElseIf elementFinal="tabpSP12" Then
	ExcelSheet.Cells(Row,(i+20)).Value ="Missing data to extend"
	Session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
	session.findById("wnd[0]").sendVKey 0
	Row=Row+1
	Exit Sub
End If

Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP04").select
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2161/ctxtMARC-MTVFP").text ="KP"
If ExcelSheet.Cells(Row,(i+62)).Value <>"" then
	Session.finbyid("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text=ExcelSheet.Cells(Row,(i+62)).Value
End If	
Session.findById("wnd[0]").sendVKey 0
If session.findById("wnd[0]/sbar").Text="Fill in all required entry fields" Then
	ExcelSheet.Cells(Row,(i+20)).Value ="Incomplete Part - Please Check"
	Session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
	session.findById("wnd[0]").sendVKey 0
	Row=Row+1
	Exit Sub
End if
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP07").select
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP09").select
session.findById("wnd[0]/tbar[0]/btn[11]").press
'session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
ExcelSheet.Cells(Row,(i+20)).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
End sub
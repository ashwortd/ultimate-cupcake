'**************************************************************************
'oOOOo.                      o                                            
 'O    `o                     O                 .oOOo.  oO   .oOOo. OooOOo 
 'o      O                    o                 O    o   O        O o      
 'O      o                    o                 o    O   o        o O      
 'o      O .oOo. `OoOo. .oOo. O  o        o   O `OooOo   O     .oO  ooOOo. 
 'O      o OooO'  o     OooO' OoO          OoO       O   o        o      O 
 'o    .O' O      O     O     o  O         o o       o   O        O      o 
' OooOO'   `OoO'  o     `OoO' O   o       O   O `OooO' OooOO `OooO' `OooO' 
'
'For all your answers and opinions give me a shout!
'**************************************************************************

'This script looks in column F of an excel spreadsheet for a quote number and it will bring back
'the Document Title text and the Mark Q/O text from that quote

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
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim ApplicationPath,CredentialsPath,FilePath
Dim ExcelApp,ExcelWorkbook,ExcelSheet,Row

Row=InputBox("Row to start at")
Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm02"
session.findById("wnd[0]").sendVKey 0
Do While ExcelSheet.Cells(Row,1).Value <>""
Session.findbyid("wnd[0]/usr/ctxtRMMG1-MATNR").text=ExcelSheet.Cells(Row,1).Value
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[19]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
'Session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,2]").select
Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
Session.findbyid("wnd[1]/usr/ctxtRMMG1-WERKS").text=" "
Session.findbyid("wnd[1]/usr/ctxtRMMG1-VKORG").text="5013"
Session.findbyid("wnd[1]/usr/ctxtRMMG1-VTWEG").text="01"
Session.findbyid("wnd[1]/tbar[0]/btn[0]").press
Session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP05/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2156/ctxtMVKE-MVGR1").text="1"
Session.findbyid("wnd[0]/tbar[0]/btn[11]").press
ExcelSheet.Cells(Row,2).Value = session.findById("wnd[0]/sbar").Text
Row=Row+1
Loop
MsgBox("The end has come")
ExcelApp.Quit
Set ExcelApp=Nothing
Set ExcelWoorkbook=Nothing
Set ExcelSheet=Nothing
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit


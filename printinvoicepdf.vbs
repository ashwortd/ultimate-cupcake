'*************Invoice Printing in PMx via CutePDF***********
'This script assumes that you have CutePDF installed on your computer
'if you have any questions please contact Derek Ashworth 860-285-9135
'copyright 2013 
'***********************************************************


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
Dim BDNumber,again,aw1,status
status = MsgBox("Please close all sessions of PMx but one",,"Information")
again = 6
Do While again = 6
BDNumber=InputBox("Input Billing Document Number","Invoice Printing")
If BDNumber="" Then
	Exit Do
End If
Set session    = connection.Children(0)
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nVF03"
session.findById("wnd[0]").sendVKey 0
session.findbyid("wnd[0]/usr/ctxtVBRK-VBELN").text = (BDNumber)
session.findById("wnd[0]/mbar/menu[0]/menu[11]").select
session.findById("wnd[1]/usr/tblSAPLZVMSGTABCONTROL").getAbsoluteRow(0).selected = true
session.findById("wnd[1]/tbar[0]/btn[6]").press
session.findById("wnd[2]/usr/chkNAST-DIMME").selected = false
session.findById("wnd[2]/usr/txtNAST-TDRECEIVER").text = ""
session.findById("wnd[2]/usr/txtNAST-TDRECEIVER").setFocus
session.findById("wnd[2]/usr/txtNAST-TDRECEIVER").caretPosition = 0
session.findById("wnd[2]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[86]").press
session.findById("wnd[0]/mbar/menu[4]/menu[8]").select

session.findById("wnd[0]/tbar[0]/okcd").text = "/nSP01"
session.findById("wnd[0]").sendVKey 0
session.findbyID("wnd[0]/tbar[1]/btn[8]").press
session.findbyId("wnd[0]/usr/chk[1,3]").selected = True
session.findbyid("wnd[0]/tbar[1]/btn[13]").press
set aw1 = session.activeWindow()
aw1.findbyid("wnd[0]/usr/ctxtTSP01_SP0R-RQDESTL").text="LOCL"
aw1.findbyid("wnd[0]/usr/cmbTSP01_SP0R-RQPOSNAME").key="CutePDF Writer"
aw1.findbyid("wnd[0]/tbar[1]/btn[13]").press
aw1.findbyid("wnd[0]").close
again=MsgBox("Do you have another?",4,"Invoice Print")
loop
Set session = connection.children(0)
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0

WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
   
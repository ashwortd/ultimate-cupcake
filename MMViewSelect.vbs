'****************************************
'Check for Logon status and connect to GUI
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   On Error Resume next
   Set connection = application.Children(0)
   If Err.Number<>0 Then
   	MsgBox("You are not connected to PMx, please connect and try again")
   	On Error Goto 0
   	WScript.Quit
   End If
   	
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
'******************************************
Dim strSelectViewFirstPos
Dim strTabName(14)
For i=0 To 14
	strTabName(i)= session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,"&i&"]").text
Next 

For b = 0 To 14
	MsgBox (strTabName(b))
Next 

	
Function selViewFirstPos()
selViewFirstPos = session.findbyid("wnd[1]/usr/tblSAPLMGMMTC_VIEW/txtMSICHTAUSW-DYTXT[0,0]").text
End Function
Function SetView(x)
 session.findById("wnd[1]/tbar[0]/btn[19]").press
	If x=1 Then
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 10
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(17).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(21).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(24).selected = True
		
	ElseIf x=2 Then
		
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(0).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(1).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(2).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(3).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(4).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(5).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(6).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(7).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(8).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(9).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(10).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(11).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(12).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").verticalScrollbar.position = 6
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(15).selected = True
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(19).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(20).selected = true
		session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(22).selected = true
	End If
End Function

MsgBox(selViewFirstPos())
If selViewFirstPos="Basic Data 1" Then
 SetView(1)
Else
 SetView(2)
 End if
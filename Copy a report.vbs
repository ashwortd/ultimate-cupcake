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
Dim x,vrc,y
Dim myNames(28)
i=1
y=1
p=0
vrc=session.findbyid("wnd[0]/usr/tblSAPMS38RTV3050").visiblerowcount
'Function GetName(x)
	For x = 0 To UBound(myNames)
	    If x=vrc Then
	    	session.findById("wnd[0]/usr/tblSAPMS38RTV3050").verticalScrollbar.position=vrc*i
	    	i=i+1
	    End if
		myNames(p)=y&"-"&session.findById("wnd[0]/usr/tblSAPMS38RTV3050/txtRS38R-QNAME1[0,"&(x)&"]").text
		y=y+1
	Next

	WScript.Echo Join( myNames, vbCrLf )
'End Function	
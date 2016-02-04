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
Dim elementID, elementLeft, elementFinal

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = "gp-17339"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[0]").press
WScript.Sleep(100)
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep(100)
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "500C"
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = "U03"
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = "001"
session.findById("wnd[1]/tbar[0]/btn[0]").press
WScript.Sleep(250)

elementID = session.ActiveWindow.GuiFocus.ID
elementLeft = Left(elementID, 50)
elementFinal = Right(elementLeft, 8)

MsgBox("The ID of the Element is:" & Chr(13) & elementID _
& Chr(13) & Chr(13) & "The left side of the ID is:" & Chr(13) _
& elementLeft & Chr(13) & Chr(13) & "The part of the ID we want is:" _
& Chr(13) & elementFinal)

session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION2").press

WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
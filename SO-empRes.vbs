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
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva03"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = "101637"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[2]/menu[2]/menu[10]").select
session.findById("wnd[0]/tbar[1]/btn[19]").press
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\09/ssubSUBSCREEN_BODY:SAPMV45A:4353/subSUBSCREEN_PARTNER_OVERVIEW:SAPLV09C:1000/tblSAPLV09CGV_TC_PARTNER_OVERVIEW/cmbGVS_TC_DATA-REC-PARVW[0,12]").setFocus

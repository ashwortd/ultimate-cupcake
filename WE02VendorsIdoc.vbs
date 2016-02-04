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

Dim V,W
V=InputBox("From Date?")
W=InputBox("To Date?")
	
session.findById("wnd[0]").resizeWorkingPane 122,28,false
session.findById("wnd[0]/tbar[0]/okcd").text = "WE02"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "dcormier"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 3
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-LOW").text = V
session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-HIGH").text = W
session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_IDOCTABBL/tabpSOS_TAB/ssub%_SUBSCREEN_IDOCTABBL:RSEIDOC2:1100/ctxtCREDAT-HIGH").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlIDOCLISTE/shellcont/shell").pressToolbarButton "&PRINT_BACK"
session.findById("wnd[1]").sendVKey 4
session.findById("wnd[2]/tbar[0]/btn[12]").press
session.findById("wnd[1]/tbar[0]/btn[13]").press
session.findById("wnd[0]").resizeWorkingPane 122,28,false
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press

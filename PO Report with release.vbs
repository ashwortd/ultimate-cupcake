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
Dim msgRet 
msgRet=MsgBox("This script will create a report to let you search"&vbCrLf&"by purchase order and show you the releases."&vbCrLf&"Push OK to continue",vbOKCancel)
If msgRet= vbCancel Then
   WScript.ConnectObject session,     "off"
   WScript.ConnectObject application, "off"
   WScript.Quit
End If	
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nsqvi"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "PO_Report"
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").caretPosition = 9
session.findById("wnd[0]/usr/btnP7").press
session.findById("wnd[1]/usr/radRB_PAINTER").select
session.findById("wnd[1]/usr/txtRS38R-HDTITLE").text = "PO_Report"
session.findById("wnd[1]/usr/txtRS38R-HDTEXT1").text = "This report will allow you to search for a customer PO and"
session.findById("wnd[1]/usr/txtRS38R-HDTEXT2").text = "it will show the release field"
session.findById("wnd[1]/usr/subSUBSOURCE:SAPMS38R:3110/ctxtRS38Q-DDNAME").text = "Vbak"
session.findById("wnd[1]/usr/radRB_PAINTER").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/lbl[1,0]").setFocus
session.findById("wnd[0]/usr/lbl[1,0]").caretPosition = 3
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[1]/shell").sapEvent "","","sapevent:RULER?BasicListWidth=254&B1=Apply"
session.findById("wnd[0]").resizeWorkingPane 156,47,false
session.findById("wnd[0]/shellcont[0]").dockerPixelSize = 395
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").expandNode "          1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "          2","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "          2","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").topNode = "          1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "          2","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "          2","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "          2","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "          2","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "          3","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "          3","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "          3","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "          3","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "          3","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "          3","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         17","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         17","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         17","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         19","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         19","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         19","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         20","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         20","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         20","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         21","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         21","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         21","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         22","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         22","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         22","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         23","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         23","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").topNode = "         11"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         23","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         39","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         39","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").topNode = "         22"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         39","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         39","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         39","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         39","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         43","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         43","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         43","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         43","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         43","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         43","COL2",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         48","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         48","COL1"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").topNode = "         32"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         48","COL1",true
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").selectItem "         48","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").ensureVisibleHorizontalItem "         48","COL2"
session.findById("wnd[0]/shellcont[0]/shellcont/shell/shellcont[0]/shell").changeCheckbox "         48","COL2",true
session.findById("wnd[0]/tbar[0]/btn[11]").press
mesRet2=MsgBox("Report complete, do not run this again.",vbOKOnly,"Complete")
session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
session.findById("wnd[0]").sendVKey 0
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
WScript.Quit
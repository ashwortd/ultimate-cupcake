Dim strPartNumber, strNewPart, strStatusBar, strCurrentTab,storlocstat

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   On Error Resume Next
	Set connection =application.Children(0)
 If Err.Number <> 0 Then
	MsgBox("You are not connected to PMx,please connect and try again")
	On Error Goto 0
	WScript.Quit
 End if
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
strNewPart="Y"
Do While strNewPart="Y"
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nmm01"
session.findById("wnd[0]").sendVKey 0
strPartNumber=InputBox("Material Number:","Caribe Copy")
If strPartNumber ="" Then
	WScript.ConnectObject session,     "off"
  	WScript.ConnectObject application, "off"
  	WScript.Quit
 End if
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = strPartNumber
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = strPartNumber
'session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").setFocus
'session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
strStatusBar=session.findById("wnd[0]/sbar").Text
If strStatusBar="Reference material does not exist" Then
	MsgBox("This part needs to be extended to Plant 500B first")
	WScript.ConnectObject session,     "off"
  	WScript.ConnectObject application, "off"
  	WScript.Quit
End If
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = false
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "5061"
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5063"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = "500b"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = "g001"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-BWTAR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text = "5013"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGNUM").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGTYP").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").caretPosition = 1
session.findById("wnd[1]/tbar[0]/btn[0]").press
strCurrentTab=Session.activewindow.guifocus.id
strCurrentTab=Left(strCurrentTab,50)
strCurrentTab=Right(strCurrentTab,8)
If strCurrentTab="tabpSP04" then
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP04/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2158/ctxtMVKE-DWERK").text = "5061"
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP05").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP06/ssubTABFRA1:SAPLMGMM:2000/subSUB5:SAPLMGD1:5802/ctxtMARC-PRCTR").text = "5060000002"
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP12/ssubTABFRA1:SAPLMGMM:2000/subSUB3:SAPLMGD1:2482/ctxtMARC-DISPO").text = "001"
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP13").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP14").select
	storlocstat=session.findbyid("wnd[0]/sbar").Text
	If storlocstat="The external procurement storage location G001 does not exist for plant 5061" Then
		session.findbyid("wnd[0]/usr/tabsTABSPR1/tabpSP13/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2484/ctxtMARC-LGFSB").text="0001"
		storlocstat="none"
	End If
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP15").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP19").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP21").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP22").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP24").select
	session.findById("wnd[0]/usr/tabsTABSPR1/tabpSP26").select
	session.findById("wnd[0]/tbar[0]/btn[11]").press
	
Else
	strNewPart=InputBox("This part is not eligible for automatic extension at this time. Do you have another part? (Y/N)")
End If
WScript.Sleep(500) 
session.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = strPartNumber
session.findById("wnd[0]/usr/cmbRMMG1-MBRSH").key = "A"
session.findById("wnd[0]/usr/cmbRMMG1-MTART").key = "ZENG"
session.findById("wnd[0]/usr/ctxtRMMG1_REF-MATNR").text = strPartNumber
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[1]/tbar[0]/btn[20]").press
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(13).selected = false
session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").getAbsoluteRow(14).selected = false
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtRMMG1-WERKS").text = "5061"
session.findById("wnd[1]/usr/ctxtRMMG1-LGORT").text = "0001"
session.findById("wnd[1]/usr/ctxtRMMG1-BWTAR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-VKORG").text = "5063"
session.findById("wnd[1]/usr/ctxtRMMG1-VTWEG").text = "99"
session.findById("wnd[1]/usr/ctxtRMMG1-LGNUM").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-LGTYP").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-WERKS").text = "5061"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGORT").text = "0001"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-BWTAR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VKORG").text = "5063"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-VTWEG").text = "01"
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGNUM").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1_REF-LGTYP").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").text = ""
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").setFocus
session.findById("wnd[1]/usr/ctxtRMMG1-DISPR").caretPosition = 1
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
strNewPart=InputBox("Another Part? (Y/N)")
Loop	
	WScript.ConnectObject session,     "off"
  	WScript.ConnectObject application, "off"
	WScript.Quit
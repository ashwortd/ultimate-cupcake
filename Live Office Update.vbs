' File:			Live Office Update.vbs
' Author:		Anthony Ciccarello
' Change Date:	08/08/2013
Option Explicit
Dim objExcel, objWorkbook, objSheet,filelocation,fileLocation2
Call main

Sub main()
	' User defined parameters
	fileLocation = "\\winvault01\briodata\BrioReps\Parts Offer Tracking Liveoffice.xlsx"
	Const showWindow = false
	Const autoSend = false			' Do you want to auto send (true) or see the email before sending (false)
	Const subjectText = "Open Parts Offers"
	Dim bodyText, recipientNames
	bodyText = "Please review the attached file of SAP quotes greater than $25k. " & _
			"If necessary, please update quote status, date of expected award and probability " & _
			"of award per the parts processing procedures and job aids." & vbCrLf
	recipientNames = Array("SPONZO Michael J.", "D'OSTUNI Pete", "LELASHER Kevin D.", "PLATNER Fran", _
		"ST-JOHN Walt", "BELFORTI Robert P.", "SMITH Barry", "BRINCKMAN Glenn A", "BEECHER Ann H.", "NORMAND Linda", "PEDRO Jennifer", "MALINOWSKI Jennifer A", "MAULUCCI Edward", "DOHERTY Michael")
		
	'Dim objExcel, objWorkbook, objSheet
    Dim addIns, index, liveOffice, addInFound
	
	Set objExcel = CreateObject("Excel.Application")			' Start Application
	objExcel.Application.Visible = showWindow
	Set objWorkbook = objExcel.Workbooks.Open(fileLocation)		' Open File
	
	' Look for Live Office Add-in
    addInFound = False
    Set addIns = objExcel.Application.COMAddIns
    For index = 1 To addIns.Count
        Set liveOffice = addIns(index)
        If InStr(liveOffice.Description, "Live Office") Then	' Live Office found
            addInFound = True
            Exit For
        End If
    Next
    If addInFound Then
        liveOffice.object.liveObjects.refresh
    Else
        MsgBox "Error: Could not find live office add-in", vbCritical, "Live Office Update Error"
    End If
	Call TxtNoteRtr
	
	' Save Changes and Exit
	objExcel.DisplayAlerts = False		' Ignore overwrite file alert
	'objWorkbook.SaveAs fileLocation		' Save
	fileLocation2 = "C:\Users\dma02\Desktop\Parts Offer Tracking Liveoffice "&DatePart("m",Now)&"-"&DatePart("d",Now)&"-"&DatePart("yyyy",Now)&".xlsx"
    objWorkbook.SaveAs fileLocation2
	objExcel.DisplayAlerts = True
	objExcel.Application.Quit			' Exit
	
			
	Call SendEmail(recipientNames, subjectText, bodyText, fileLocation, autoSend)
	
End Sub

Sub SendEmail(recipientNames, subjectText, bodyText, attachmentPath, autoSend)
	' Method parameter
	Const olMailItem = 0

    Dim objOutlook, objOutlookMsg, objOutlookRecip, objOutlookAttach
    ' Create the Outlook session.
    Set objOutlook = CreateObject("Outlook.Application")

    ' Create the message.
    Set objOutlookMsg = objOutlook.CreateItem(olMailItem)

    With objOutlookMsg
        ' Add the To recipient(s) to the message.
        Dim name
        For Each name In recipientNames
            Set objOutlookRecip = .Recipients.Add(name)
        Next

       ' Set the Subject, Body, and Importance of the message.
        .Subject = subjectText
        .Body = bodyText
       '.Importance = olImportanceHigh  'High importance

       ' Add attachments to the message.
        .Attachments.Add(fileLocation2)


       ' Resolve each Recipient's name.
       For Each objOutlookRecip In .Recipients
           objOutlookRecip.Resolve
       Next

       ' Should we display the message before sending?
       If Not autoSend Then
           .Display
       Else
       MsgBox("nana nana boo boo")
'           .Save
'           .Send
       End If
    End With
End Sub

'Subroutine added by Derek Ashworth to bring in text notes.
Sub TxtNoteRtr
'This script looks in column F of an excel spreadsheet for a quote number and it will bring back
'the Document Title text and the Mark Q/O text from that quote
Dim SapGuiApp,Connection,Session,FileObject,ofile,Counter
Dim row,application, SapGuiAuto 

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

Set objSheet = objWorkbook.ActiveSheet
objExcel.Visible=True

Row=2

Session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "/nva23"
session.findById("wnd[0]").sendVKey 0
Do While objSheet.Cells(Row,6).Value <>""
If objSheet.Cells(Row,6).Value ="Sum:" Then
	Row=Row+1
	Exit Do
End If

session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = objSheet.Cells(Row,6).Value
Session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\02/ssubSUBSCREEN_BODY:SAPMV45A:4411/subSUBSCREEN_TC:SAPMV45A:4912/tblSAPMV45ATCTRL_U_ERF_ANGEBOT/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]").sendVKey 2
objSheet.Cells(Row,33).Value=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/ctxtVBAP-AWAHR").text
Session.findbyid("wnd[0]/tbar[1]/btn[19]").press
objSheet.Cells(Row,34).Value=session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4451/ctxtVBAP-AWAHR").text
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press

Session.findById("wnd[0]/usr/subSUBSCREEN_HEADER:SAPMV45A:4021/btnBT_HEAD").press
Session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08").select

session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectitem "0020","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleclickitem "0020","Column1"
objSheet.Cells(Row,31).Value=(session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text)

session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").selectitem "0035","Column1"
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").doubleclickitem "0035","Column1"
objSheet.Cells(Row,32).Value=(session.findbyid("wnd[0]/usr/tabsTAXI_TABSTRIP_HEAD/tabpT\08/ssubSUBSCREEN_BODY:SAPMV45A:4152/subSUBSCREEN_TEXT:SAPLV70T:2100/cntlSPLITTER_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").text)
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Session.findbyid("wnd[0]/tbar[0]/btn[3]").press
Row=Row+1
Loop

'****Execl sheet cleanup
objSheet.Cells(1,31).Value="Document Description"
objSheet.cells(1,31).font.bold=True
objSheet.cells(1,31).font.colorindex = 2
objSheet.cells(1,31).interior.colorindex=11
objSheet.Columns(31).columnwidth=30
objSheet.Range("AE:AF").wraptext=true
objSheet.Cells(1,32).Value="Document Text Notes"
objSheet.cells(1,32).font.bold=True
objSheet.cells(1,32).font.colorindex = 2
objSheet.cells(1,32).interior.colorindex=11
objSheet.Columns(32).columnwidth=90
objSheet.columns.autofit


'****Disconnect SAP
WScript.ConnectObject session,     "off"
WScript.ConnectObject application, "off"
End Sub

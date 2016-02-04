Option Explicit
Const DefaultCalendarBorderColor = "#000000"
Const highlightcolor = "Yellow"
Dim objFSO, objTF
Dim weekstogo, borderSize, borderColor, monthName
Dim thisday, mydate, myday, iThisMonth, iThisYear, strFirstmdy
Dim iFirstDayofThisMonth, iDaysThisMonth, iDayToDisplay
Dim opLine, i, x, daysOffset, daysLeft, decWeekstogo, Msg
Dim iDayofYear, IE, wshShell, strOutput
Dim strWebPagename, tempfolder, iCount, oCount, MyItems, CurrAppt, dtThisMonday, dtTheSundayAfter
borderSize = 1
borderColor = DefaultCalendarBorderColor
thisday = Day(Date)
myDate = Date
myDay = DatePart("D", Date)
iThisMonth = Month(Date)
iThisYear = Year(Date)
strFirstmdy = (iThisMonth & "/1/" & iThisYear)
iFirstDayofThisMonth = DatePart("W", strFirstmdy)
' Store the month names into an array.
ReDim monthName(13)
monthName(0) = "Space" 'Set element 0 to garbage so I don't have to do math later
monthName(1) = "January"
monthName(2) = "February"
monthName(3) = "March"
monthName(4) = "April"
monthName(5) = "May"
monthName(6) = "June"
monthName(7) = "July"
monthName(8) = "August"
monthName(9) = "September"
monthName(10) = "October"
monthName(11) = "November"
monthName(12) = "December"
'Calculate number of days this month
If iThisMonth = 12 Then
iDaysThisMonth = DateDiff("d",strFirstmdy,("1/1/" & (iThisYear+1)))
Else
iDaysThisMonth = DateDiff("d",strFirstmdy,((iThisMonth+1) & "/1/" & iThisYear))
End If
'*------------*&
'* Begin creation of the output file here
'*------------*&
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
' We don't use "temporary" file names, because they really aren't, and they hang around forever,
' or until someone specifically deletes them.
strWebPagename = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(".") & _
"\VBSCalendar.html"
Set objTF = objFSO.CreateTextFile(strWebPageName)
objTF.writeline("<%@ LANGUAGE=""VBSCRIPT"" %>")
objTF.writeline("<HTML>")
objTF.writeline("<HEAD>")
objTF.writeline("<TITLE>My Calendar</TITLE>")
objTF.writeline("</HEAD>")
objTF.writeline("<body bgcolor=""#FFFFD8"">")
Msg = FormatDateTime(Now(),vbLongDate)
objTF.writeline("<center><h1>" & msg & "</h1>")
OpLine = "<table cellpadding=""4"" cellspacing=""1"" border=""0"" bgcolor=""#ffffff"">"
OpLine = OpLine & "<tr><td colspan=""7""align=""center"" bgcolor=Yellow><b>"
OpLine = OpLine & (monthName(iThisMonth) & " " & iThisYear) & "</b></td></tr>"
objTF.writeLine(OpLine)
' Write the row of weekday initials
objTF.writeLine("<tr>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">S</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">M</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">T</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">W</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">R</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">F</td>")
objTF.writeLine("<td align=""center"" bgcolor=""#A1C6D1"">S</td>")
objTF.writeLine("</tr>")
objTF.writeLine("<tr>")
'Now write the first row
For i = 1 to 7
if i = iFirstDayofThisMonth Then
iDayToDisplay = 1
elseif i > iFirstDayofThisMonth Then
iDayToDisplay = iDayToDisplay + 1
else
iDayToDisplay="&nbsp;"
end if
if iDayToDisplay = thisDay Then
Msg = "<td align=center bgcolor=Yellow><b>" & iDayToDisplay & "</b></td>"
else
Msg = "<td align=center>" & iDayToDisplay & "</td>"
end If
objTF.writeLine(Msg)
Next
' Now, display the rest of the month.
' First figure out how many weeks are left to write
daysOffSet = 8 - iFirstDayofThisMonth
daysLeft = iDaysThisMonth - daysOffset
decWeekstogo = Round((daysLeft/7),2)
' I think this logic is screwy.
If decweekstogo > 4 Then
weekstogo = 5
ElseIf decweekstogo > 3 and decweekstogo <= 4 Then
weekstogo = 4
Else
weekstogo = 3
End if
' Now write the rows and populate the data
For x = 1 To weekstogo
objTF.writeLine("<tr>")
For i = 1 To 7
If iDayToDisplay < iDaysThisMonth then
iDayToDisplay = iDayToDisplay + 1
else
iDayToDisplay = "&nbsp;"
End If
if iDayToDisplay = thisDay then
Msg = "<td align=center bgcolor=Yellow><b>" & iDayToDisplay & "</b></td>"
else
Msg = "<td align=center>" & iDayToDisplay & "</td>"
end If
objTF.writeLine(Msg)
Next
objTF.writeLine("</tr>")
Next
objTF.writeLine("</table>")
objTF.writeLine("</center>")
' The calendar is finished displaying. Now display some information that *might* be useful.
iDayofYear = DateDiff("d","01/01/" & year(now),month(now) & "/" & day(now) & "/" & year(now)) + 1
objTF.writeLine("The Day of the Year is " & iDayofYear)
ProcessOutlook '* Display any meetings from Outlook
objTF.writeLine("</BODY>")
objTF.writeLine("</HTML>")
objTF.close
'*------------*&
'* Now display the web page
'*------------*&
Set wshShell = WScript.CreateObject ("WSCript.shell")
Set IE = CreateObject("InternetExplorer.Application")
IE.visible = 1
IE.navigate("O:\CustSvc\Parts\Pmx Scripting\Scripts\Scripts\VBSCalendar.html")
'* We can't delete the file because we get here right after the web page is displayed, and it is still in use.
'objFSO.DeleteFile(strWebPagename)
WScript.quit
'*-------------*
'* Go to Outlook for appointments
'*-------------*
Sub ProcessOutlook
Dim objOutlook, objNameSpace, objFolder
'Const olMailItem = 0
'Const olTaskItem = 3
Const olFolderTasks = 13
Const olFolderCalender = 9
'Create Outlook, Namespace, Folder Objects and Task Item
Set objOutlook = CreateObject("Outlook.application")
Set objNameSpace = objOutlook.GetNameSpace("MAPI")
Set objFolder = objNameSpace.GetDefaultFolder(olFolderCalender)
Set MyItems = objFolder.Items
MyItems.IncludeRecurrences = True
myItems.Sort "[Start]"
' Check for meetings THIS week
dtThisMonday = dateadd("d", 1 - weekday(date), date) + 1
strOutput = strOutput & "<Center><h3> Meetings This Week </h3></center>"
DispOneWeek
' Check for meetings NEXT week
dtThisMonday = dateadd("d", +7, dtThisMonday)
strOutput = strOutput & "<Center><h3> Meetings NEXT Week </h3></center>"
DispOneWeek
' Check for meetings NEXT week
dtThisMonday = dateadd("d", +7, dtThisMonday)
strOutput = strOutput & "<Center><h3> And, the week AFTER that </h3></center>"
DispOneWeek
'Display results to user, if any.
If strOutput > "" Then
objTF.writeLine(strOutput)
Else
objTF.writeLine("Meetings for this week: NONE<br>")
End If
'Clean up
Set objFolder = Nothing
Set objNameSpace = Nothing
set objOutlook = Nothing
End sub
'*-------------*
'* Sub DispOneWeek
'*-------------*
' This routine will display one week's worth of Outlook reminders.
Sub DispOneWeek
oCount = 0 : iCount = 0
dtTheSundayAfter = DateAdd("d", +6, dtThisMonday)
For Each CurrAppt in MyItems
'If CurrAppt.BusyStatus = 2 and CurrAppt.Sensitivity = 0 then ‘????
iCount = iCount + 1
If iCount > 365 then ' Limit the number of recurring entries to look at
exit for ' Bail on THIS item only
end if
If CurrAppt.Start >= dtThisMonday And _
CurrAppt.Start <= dttheSundayAfter Then
CrOpLine
End If
Next
End sub
'*-------------*
'* Sub CrOpLine - Create an Output line for DispOneWeek
'*-------------*
Sub CrOpLine
oCount = oCount + 1
strOutput = strOutput & oCount & ". " & _
"<b>Subject:</b> " & CurrAppt.Subject & _
" <b>Date/Time:</b> " & CurrAppt.Start & _
" <b>Duration</b> " & CurrAppt.Duration & "<br>"
' "; recurrence pattern=" & CurrAppt.GetRecurrencePattern & "<br><br>"
End sub
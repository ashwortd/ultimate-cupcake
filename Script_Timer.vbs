Dim start, finish

start = Now()

WScript.Sleep(10000)

finish = Now()

If DateDiff("s",start,finish) < 60 Then
	MsgBox("Script completed in: " & DateDiff("s",start,finish) & " seconds.")

ElseIf DateDiff("s",start,finish) > 120 And DateDiff("s",start,finish) Mod 60 <> 1 Then
	MsgBox("Script completed in: " & Int(DateDiff("s",start,finish)/60) & " minutes and "& DateDiff("s",start,finish) Mod 60 & " seconds.")

ElseIf DateDiff("s",start,finish) < 120 And DateDiff("s",start,finish) Mod 60 <> 1 Then
	MsgBox("Script completed in: " & Int(DateDiff("s",start,finish)/60) & " minute and "& DateDiff("s",start,finish) Mod 60 & " seconds.")

ElseIf DateDiff("s",start,finish) = 61 Then
	MsgBox("Script completed in: " & Int(DateDiff("s",start,finish)/60) & " minute and "& DateDiff("s",start,finish) Mod 60 & " second.")
	
Else
	MsgBox("Script completed in: " & Int(DateDiff("s",start,finish)/60) & " minutes and "& DateDiff("s",start,finish) Mod 60 & " second.")

End If

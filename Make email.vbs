Dim MyAppID,x
Set x=CreateObject("wscript.shell")
'y.ActiveWindow
x.Run "outlook.exe"
WScript.Sleep 3000
x.SendKeys"^1"
x.SendKeys"^n"
x.sendkeys"{tab}{tab}{tab}Part extension complete{tab}"
x.SendKeys"Thanks for using the automated part extension system{enter}Regards,{enter}Derek"
		
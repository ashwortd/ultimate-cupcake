Dim d   ' Create some variables.
Set d = CreateObject("Scripting.Dictionary")

d.Add "a", "Athens"   ' Add some keys and items.
d.Add "b", "Belgrade"
d.Add "c", "Cairo"

d.Item("c") = "NewValue"

itemDemo = d.Key("Cairo")   ' Get the key.
MsgBox itemDemo
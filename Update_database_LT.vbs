Call updateAccess
Public Sub updateAccess()

Pathname="C:\Users\dma02\Desktop\Derek\Fruzi Test\"
Filename = Pathname & "Material Master Exception database.accdb"

Set con=CreateObject("ADODB.Connection")
connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & Filename

Dim sql
Dim newTable 
con.Open connectionString

' Save current table ("BEFORE") to another table ("BEFORE_yyyymmdd_hh_mmss")
newTable = "1-FIXEDVENDOR_old"
sql = "INSERT * INTO " & newTable & " FROM 1-FIXEDVENDOR"
con.Execute sql

' Delete rows of current table ("BEFORE")
sql = "DELETE FROM 1-FIXEDVENDOR"
con.Execute sql

' Insert new rows into current table ("BEFORE") from my Excel Sheet
sql = "INSERT INTO 1-FIXEDVENDOR ([Material], [Vendor]) " & _
      "SELECT * FROM [Excel 8.0;HDR=YES;DATABASE=" & ThisWorkbook.FullName & "].[" & ThisWorkbook.Sheets("CODE_BY_STORE").Name & "$]"
con.Execute sql

con.Close
Set con = Nothing

End Sub
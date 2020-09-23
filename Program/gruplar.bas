Attribute VB_Name = "gruplar_liste"
Function grup_listesi()
grup_ekle.ListView1.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * From gruplar ", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With grup_ekle.ListView1
Set itm = .ListItems.Add(, , rs!grupadi, 1, 1)
itm.SubItems(1) = rs!id
End With
rs.MoveNext
Loop
rs.Close
End Function

Attribute VB_Name = "Sorgular"
Function gruplar()
kullanici_ekle.Combo1.Clear
Dim conn As New ADODB.Connection 'Baðlantýmýzý tanýmladýk.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * from gruplar", conn, adOpenKeyset, adLockPessimistic
Do Until rs.EOF
kullanici_ekle.Combo1.AddItem rs!grupadi
rs.MoveNext
Loop
rs.Close
End Function
Function liste_kullanicilar()
'Form1 Baþlangýçtaki Kullanýcýlarý EKLE
Form1.ListView2.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'Sadece Aktif Olanlar Gözükmeli
rs.Open "Select * From kullanicilar Where aktif ='1'", conn, adOpenKeyset, adLockOptimistic
'rs.Open "Select * from kullanicilar ", conn, adOpenKeyset, adLockPessimistic
Do Until rs.EOF
Dim itm As ListItem
With Form1.ListView2
Set itm = .ListItems.Add(, , rs!adi_soyadi, 1, 1)
itm.SubItems(1) = rs!kullanici_adi
End With
rs.MoveNext
Loop
rs.Close
End Function
Function liste_kullanicilar1()
'Kullanýcý Ekle Bölümüne Kullanýcýlar Eklenecek.
kisi_sec.ListView1.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'rs.Open "Select * from kullanicilar ", conn, adOpenKeyset, adLockPessimistic
rs.Open "Select * From kullanicilar Where aktif ='1'", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With kisi_sec.ListView1
Set itm = .ListItems.Add(, , rs!adi_soyadi, 1, 1)
itm.SubItems(1) = rs!kullanici_adi
End With
rs.MoveNext
Loop
rs.Close
End Function
Function okunmamis_mesajlar()
'Okunmamýþ Mesajlarý Gösterteceðiz.
Form1.okunmamis.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'Sadece Aktif Olanlar Gözükmeli
rs.Open "select * from mesajlar where kime='" & Baglan.Text1.Text & "' and okundu='0'", conn, adOpenKeyset, adLockOptimistic
'rs.Open "Select * from kullanicilar ", conn, adOpenKeyset, adLockPessimistic
Do Until rs.EOF
Dim itm As ListItem
With Form1.okunmamis
Set itm = .ListItems.Add(, , rs!okundu)
'itm.SubItems(1) = rs!kullanici_adi
End With
rs.MoveNext
Loop
rs.Close
End Function
Function arsiv_listesi()
'Arþiv Listesi Mesajlarý Gösterteceðiz.
Form1.arsiv.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'Sadece Aktif Olanlar Gözükmeli
rs.Open "select * from mesajlar where kime='" & Baglan.Text1.Text & "' and gonderilmedi='1'", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With Form1.arsiv
Set itm = .ListItems.Add(, , rs!gonderilmedi)
End With
rs.MoveNext
Loop
rs.Close
End Function
Function secilen_gruplar()
kisi_sec.Combo1.Clear
kisi_sec.Combo1.AddItem "Tümü"
Dim conn As New ADODB.Connection 'Baðlantýmýzý tanýmladýk.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * from gruplar", conn, adOpenKeyset, adLockPessimistic
Do Until rs.EOF
kisi_sec.Combo1.AddItem rs!grupadi
rs.MoveNext
Loop
rs.Close
End Function
Function liste_kullanicilar2()
'Kullanýcý Ekle Bölümüne Kullanýcýlar Eklenecek.
kisi_sec.ListView1.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * From kullanicilar Where aktif ='1' and grup= '" & kisi_sec.Combo1.Text & " '", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With kisi_sec.ListView1
Set itm = .ListItems.Add(, , rs!adi_soyadi, 1, 1)
itm.SubItems(1) = rs!kullanici_adi
End With
rs.MoveNext
Loop
rs.Close
End Function
Function kayitli_kullanicilar_listesi()
kullanici_listesi.ListView1.ListItems.Clear
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'Tümü Gözükmeli
rs.Open "Select * From kullanicilar ", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With kullanici_listesi.ListView1
Set itm = .ListItems.Add(, , rs!adi_soyadi)
itm.SubItems(1) = rs!kullanici_adi
itm.SubItems(2) = rs!sifre
itm.SubItems(3) = rs!kayit_tarihi
itm.SubItems(4) = rs!grup
itm.SubItems(5) = rs!aciklama

End With
rs.MoveNext
Loop
rs.Close
End Function

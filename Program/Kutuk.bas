Attribute VB_Name = "Mesaj_Oku"
Function gelen_mesajlar()
Form1.ListView1.ListItems.Clear
Form1.ListView1.Refresh
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
Dim giris
giris = Baglan.Text1.Text
'Sorgu Ekran�
'Okunmam�� Olcak
'Silinmemi� Olcak
'Gonderilen Olmayacak
'Gonderilmemi� Olmayacak
rs.Open _
"Select * From mesajlar Where kime ='" & giris & "' and silindi='0' and gonderilmedi ='0'", conn, adOpenKeyset, adLockOptimistic
Do Until rs.EOF
Dim itm As ListItem
With Form1.ListView1
On Error Resume Next
Set ListView1.SmallIcons = ImageList5
'T�m Kay�tlar� Listele
Set itm = .ListItems.Add(, , , , 7)
itm.SubItems(4) = rs!mesajid
itm.SubItems(5) = rs!kimden & " " & rs!txtbilgi
itm.SubItems(6) = rs!konu
itm.SubItems(7) = rs!gonderim_tarihi
itm.SubItems(8) = rs!okundu_tarih
'Mesaj Durumu
itm.SubItems(9) = rs!okundu
itm.SubItems(10) = rs!silindi
'Ek Bilgileri
itm.SubItems(3) = rs!ek
' Durum Bilgileri
itm.SubItems(1) = rs!acil
itm.SubItems(2) = rs!inceleyin
itm.SubItems(11) = rs!id
'itm.SubItems(3) = rs!bilgilendirme
'itm.SubItems() = rs!yanitlayin
'itm.SubItems(6) = rs!mesaj
Dim i
For i = 1 To Form1.ListView1.ListItems.Count
If Form1.ListView1.ListItems.Item(i).SubItems(9) = "0" Then
'E�er Okunmad� De�eri 0 ise Mesaj� Kal�n Harflerle G�ster.
itm.ListSubItems.Item(1).Bold = True
itm.ListSubItems.Item(2).Bold = True
itm.ListSubItems.Item(3).Bold = True
itm.ListSubItems.Item(4).Bold = True
itm.ListSubItems.Item(5).Bold = True
itm.ListSubItems.Item(6).Bold = True
itm.ListSubItems.Item(7).Bold = True
itm.ListSubItems.Item(8).Bold = True
itm.ListSubItems.Item(9).Bold = True
itm.ListSubItems.Item(10).Bold = True
Form1.ListView1.ListItems.Item(i).SmallIcon = 10 '7
Else
itm.ListSubItems.Item(1).Bold = False
itm.ListSubItems.Item(2).Bold = False
itm.ListSubItems.Item(3).Bold = False
itm.ListSubItems.Item(4).Bold = False
itm.ListSubItems.Item(5).Bold = False
itm.ListSubItems.Item(6).Bold = False
itm.ListSubItems.Item(7).Bold = False
itm.ListSubItems.Item(8).Bold = False
itm.ListSubItems.Item(9).Bold = False
itm.ListSubItems.Item(10).Bold = False
Form1.ListView1.ListItems.Item(i).SmallIcon = 11 '2
End If
Next i
End With
rs.MoveNext
Loop
rs.Close
Form1.Label19.Caption = (0)
okunmamis_mesajlar
Form1.Label19.Caption = "(" & Form1.okunmamis.ListItems.Count & ")" 'Ka� Tane Yeni Mesaj Oldu�unu Sorgula
End Function




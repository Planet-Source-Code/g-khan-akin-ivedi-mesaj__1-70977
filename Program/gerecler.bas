Attribute VB_Name = "gerecler"
Function kontor()
'Gönderilen Mesajlarýn ID Numaralarý +1 olarak Gösteriliyor.
'Hata Tespiti: Mesaj ID Kayýt Numarasýna Göre Kontor Ekliyo Buda kayýt _
silindiði zaman sýra karýþýyo doðal olarak ve ayný numaradan iki _
kayýt oluyo. buda silme _ okuma gibi iþlemlerde  sorun cýkartýyo. _
Yapman Gereken Enson kayda gidip oradaki son mesajýn mesaj ýd sini alýp sonuna 1 eklemen _
BaSkada olur yolu yok.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs1 As New ADODB.Recordset
rs1.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
'yeni_mesaj_gonder.Label5.Caption = rs1.RecordCount
'yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount + 1
'*********************************** Kayýt Numarasý(Sabit)
yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount
yeni_mesaj_gonder.Label8.Caption = yeni_mesaj_gonder.Label8.Caption + 1
'************************************ Mesaj Numarasý
yeni_mesaj_gonder.Label10.Caption = rs1.RecordCount
yeni_mesaj_gonder.Label10.Caption = yeni_mesaj_gonder.Label10.Caption + 1
'************************************ Kayýt Sayýsýný Göster
yeni_mesaj_gonder.StatusBar1.Panels.Item(9).Text = rs1.RecordCount + 1
Close
End Function

Function kontor2()
'Gönderilen Mesajlarýn ID Numaralarý +1 olarak Gösteriliyor.
'Hata Tespiti: Mesaj ID Kayýt Numarasýna Göre Kontor Ekliyo Buda kayýt _
silindiði zaman sýra karýþýyo doðal olarak ve ayný numaradan iki _
kayýt oluyo. buda silme _ okuma gibi iþlemlerde  sorun cýkartýyo. _
Yapman Gereken Enson kayda gidip oradaki son mesajýn mesaj ýd sini alýp sonuna 1 eklemen _
BaSkada olur yolu yok.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs1 As New ADODB.Recordset
rs1.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
'yeni_mesaj_gonder.Label5.Caption = rs1.RecordCount
'yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount + 1
'*********************************** Kayýt Numarasý(Sabit)
mesaji_ilet.Label5.Caption = rs1.RecordCount
mesaji_ilet.Label5.Caption = mesaji_ilet.Label5.Caption + 1
'************************************ Mesaj Numarasý
mesaji_ilet.Label6.Caption = rs1.RecordCount
mesaji_ilet.Label6.Caption = mesaji_ilet.Label6.Caption + 1
'************************************ Kayýt Sayýsýný Göster
'mesaji_ilet.StatusBar1.Panels.Item(9).Text = rs1.RecordCount + 1
Close
End Function


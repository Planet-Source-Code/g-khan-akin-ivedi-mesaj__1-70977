Attribute VB_Name = "gerecler"
Function kontor()
'G�nderilen Mesajlar�n ID Numaralar� +1 olarak G�steriliyor.
'Hata Tespiti: Mesaj ID Kay�t Numaras�na G�re Kontor Ekliyo Buda kay�t _
silindi�i zaman s�ra kar���yo do�al olarak ve ayn� numaradan iki _
kay�t oluyo. buda silme _ okuma gibi i�lemlerde  sorun c�kart�yo. _
Yapman Gereken Enson kayda gidip oradaki son mesaj�n mesaj �d sini al�p sonuna 1 eklemen _
BaSkada olur yolu yok.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs1 As New ADODB.Recordset
rs1.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
'yeni_mesaj_gonder.Label5.Caption = rs1.RecordCount
'yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount + 1
'*********************************** Kay�t Numaras�(Sabit)
yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount
yeni_mesaj_gonder.Label8.Caption = yeni_mesaj_gonder.Label8.Caption + 1
'************************************ Mesaj Numaras�
yeni_mesaj_gonder.Label10.Caption = rs1.RecordCount
yeni_mesaj_gonder.Label10.Caption = yeni_mesaj_gonder.Label10.Caption + 1
'************************************ Kay�t Say�s�n� G�ster
yeni_mesaj_gonder.StatusBar1.Panels.Item(9).Text = rs1.RecordCount + 1
Close
End Function

Function kontor2()
'G�nderilen Mesajlar�n ID Numaralar� +1 olarak G�steriliyor.
'Hata Tespiti: Mesaj ID Kay�t Numaras�na G�re Kontor Ekliyo Buda kay�t _
silindi�i zaman s�ra kar���yo do�al olarak ve ayn� numaradan iki _
kay�t oluyo. buda silme _ okuma gibi i�lemlerde  sorun c�kart�yo. _
Yapman Gereken Enson kayda gidip oradaki son mesaj�n mesaj �d sini al�p sonuna 1 eklemen _
BaSkada olur yolu yok.
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs1 As New ADODB.Recordset
rs1.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
'yeni_mesaj_gonder.Label5.Caption = rs1.RecordCount
'yeni_mesaj_gonder.Label8.Caption = rs1.RecordCount + 1
'*********************************** Kay�t Numaras�(Sabit)
mesaji_ilet.Label5.Caption = rs1.RecordCount
mesaji_ilet.Label5.Caption = mesaji_ilet.Label5.Caption + 1
'************************************ Mesaj Numaras�
mesaji_ilet.Label6.Caption = rs1.RecordCount
mesaji_ilet.Label6.Caption = mesaji_ilet.Label6.Caption + 1
'************************************ Kay�t Say�s�n� G�ster
'mesaji_ilet.StatusBar1.Panels.Item(9).Text = rs1.RecordCount + 1
Close
End Function


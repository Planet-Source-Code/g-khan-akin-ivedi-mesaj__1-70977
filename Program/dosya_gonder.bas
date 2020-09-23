Attribute VB_Name = "DosyaGonder"
Function dosya_gonder()
'Dosya Transferi
If yeni_mesaj_gonder.Text4.Text = "" Then 'Seçili Dosya Yok ise
Else
' Dosyalar Oncelikle Server Üzerinde Dosya  Klasörünün Ýçerisine Gönderilecek.
' Kullanýcýlar Dosyalarý Server Üzerinden Kopyalama Yapacaklar.
' Önce Virgül ile ayrýlan dosyalarý teker teker ayýr.
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = yeni_mesaj_gonder.Text4.Text ' Okunacak Metin.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
Dim kaynak, hedef, sonuc, hata, dosya, cevap
'------------------------------------------
kaynak = se_next_to ' kaynak dosyalar
'------------------------------------------
hedef = programayarlari.Text2.Text & se_next_to  'Server Üzerindeki Adres...
'------------------------------------------
dosya = Dir(programayarlari.Text2.Text & se_next_to)
If dosya = "" Then 'Eðer Dosya Yok Ýse
FileCopy kaynak, hedef 'Kopyalama iþlemini Baþlat.Dosyalar Server Üzerine Kopyalanýyor
Else
MsgBox "Dosya Server Üzerinde Mevcut Lütfen dosyanýzý Farklý Bir Ýsim ile Tekrar Gödnerin." & vbCrLf & vbCrLf & "Gönderilemeyen Dosya : " & se_next_to, vbExclamation, "Bilgi Mesajý"
yeni_mesaj_gonder.Text4.Text = "" ' Mesaj Giderse Bile Dosya Eki Gitmesin.
yeni_mesaj_gonder.Text4.SetFocus
End If
End If
Loop ' Kopyalama Ýþlemini Dosya Sayýsýna Göre Tekrarlýyoruz.
Exit Function
hata:
MsgBox "Dosya Kopyalama Hatasý.Sistem Yöneticinize Baþvurunuz.", vbCritical, "Dosya Transfer Hatasý"
Exit Function
'Yukarýda Belirtilen Yol Serverde Bir Yol Olacak Unutma !!
End If
End Function

Function dosya_kontrol()
' Eðer Bu Mesajda Ek Geldiyse Ekleri Toobar'a Yükle
If Form1.Label15.Caption = "" Then 'Eðer Ek Yoksa Tollbari Gizle
Form1.Toolbar2.Visible = False
Else
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Clear
Form1.Toolbar2.Refresh
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = Form1.Label15.Caption  ' Okunacak Metin.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , se_next_to
End If
Loop
'Dögü Bittikten Sonra Son Olarak Tümünü Kaydet Butonu Ekle
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "-"
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "Tümünü kaydet"
End If
End Function
Function dosya_kontrol2()
' Eðer Bu Mesajda Ek Geldiyse Ekleri Toobar'a Yükle
'If Form1.Label15.Caption = "" Then 'Eðer Ek Yoksa Tollbari Gizle
'Form1.Toolbar2.Visible = False
'Else
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Clear
yeni_mesaj_gonder.Toolbar2.Refresh
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = yeni_mesaj_gonder.Label25.Caption  ' Okunacak Metin.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , se_next_to
End If
Loop
'Dögü Bittikten Sonra Son Olarak Tümünü Kaydet Butonu Ekle
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "-"
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "Tümünü kaydet"
'End If
End Function


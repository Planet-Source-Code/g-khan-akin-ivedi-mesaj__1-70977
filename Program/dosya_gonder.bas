Attribute VB_Name = "DosyaGonder"
Function dosya_gonder()
'Dosya Transferi
If yeni_mesaj_gonder.Text4.Text = "" Then 'Se�ili Dosya Yok ise
Else
' Dosyalar Oncelikle Server �zerinde Dosya  Klas�r�n�n ��erisine G�nderilecek.
' Kullan�c�lar Dosyalar� Server �zerinden Kopyalama Yapacaklar.
' �nce Virg�l ile ayr�lan dosyalar� teker teker ay�r.
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
hedef = programayarlari.Text2.Text & se_next_to  'Server �zerindeki Adres...
'------------------------------------------
dosya = Dir(programayarlari.Text2.Text & se_next_to)
If dosya = "" Then 'E�er Dosya Yok �se
FileCopy kaynak, hedef 'Kopyalama i�lemini Ba�lat.Dosyalar Server �zerine Kopyalan�yor
Else
MsgBox "Dosya Server �zerinde Mevcut L�tfen dosyan�z� Farkl� Bir �sim ile Tekrar G�dnerin." & vbCrLf & vbCrLf & "G�nderilemeyen Dosya : " & se_next_to, vbExclamation, "Bilgi Mesaj�"
yeni_mesaj_gonder.Text4.Text = "" ' Mesaj Giderse Bile Dosya Eki Gitmesin.
yeni_mesaj_gonder.Text4.SetFocus
End If
End If
Loop ' Kopyalama ��lemini Dosya Say�s�na G�re Tekrarl�yoruz.
Exit Function
hata:
MsgBox "Dosya Kopyalama Hatas�.Sistem Y�neticinize Ba�vurunuz.", vbCritical, "Dosya Transfer Hatas�"
Exit Function
'Yukar�da Belirtilen Yol Serverde Bir Yol Olacak Unutma !!
End If
End Function

Function dosya_kontrol()
' E�er Bu Mesajda Ek Geldiyse Ekleri Toobar'a Y�kle
If Form1.Label15.Caption = "" Then 'E�er Ek Yoksa Tollbari Gizle
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
'D�g� Bittikten Sonra Son Olarak T�m�n� Kaydet Butonu Ekle
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "-"
Form1.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "T�m�n� kaydet"
End If
End Function
Function dosya_kontrol2()
' E�er Bu Mesajda Ek Geldiyse Ekleri Toobar'a Y�kle
'If Form1.Label15.Caption = "" Then 'E�er Ek Yoksa Tollbari Gizle
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
'D�g� Bittikten Sonra Son Olarak T�m�n� Kaydet Butonu Ekle
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "-"
yeni_mesaj_gonder.Toolbar2.Buttons.Item(1).ButtonMenus.Add , , "T�m�n� kaydet"
'End If
End Function


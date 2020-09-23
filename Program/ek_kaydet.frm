VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ek_kaydet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ekleri Kaydet"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "ek_kaydet.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   1440
      TabIndex        =   11
      Top             =   3510
      Width           =   5955
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Tümünü Kaydet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      TabIndex        =   8
      Top             =   2970
      Width           =   1410
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   135
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ek_kaydet.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   3510
      Width           =   5865
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Ka&pat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5265
      TabIndex        =   6
      Top             =   720
      Width           =   1410
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Kaydet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5265
      TabIndex        =   5
      Top             =   315
      Width           =   1410
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3075
      Left            =   45
      TabIndex        =   4
      Top             =   315
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8996
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Gözat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4545
      TabIndex        =   3
      Top             =   3780
      Width           =   1500
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   90
      TabIndex        =   1
      Text            =   "C:\ALINAN DOSYALARIM\"
      Top             =   3780
      Width           =   4380
   End
   Begin VB.Label Label2 
      Caption         =   "Kaydedilecek Ekler:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   2490
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   9
      Top             =   5265
      Width           =   4380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kaydedilecek Yer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   3420
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Listview'e Eklenen Dosyalar Buradan Ekleniyor."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   6075
      Width           =   4380
   End
End
Attribute VB_Name = "ek_kaydet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Dim strResFolder As String
strResFolder = BrowseForFolder(hWnd, "Lütfen Herhangi bir Dizin Belirleyiniz.")
If strResFolder = "" Then
Else
Text1.Text = strResFolder & "\"
End If
End Sub

Private Sub Command4_Click()

 End Sub

Private Sub Command8_Click()


End Sub


Private Sub Command5_Click()
On Error GoTo hata
If Label1.Caption = "" Then
MsgBox "Kopyalama Yapýlack Dosyayý Seçiniz.", vbCritical, "Uyarý"
Else
'Kopyalama Ýþlemine Baþla
Dim kaynak, hedef
kaynak = programayarlari.Text2.Text
hedef = Text1.Text
'Eðer Dosya Hedefidne Yok ise Ýþlemi Baþlat
'Burada Ayný Ýsimli dosyadan Varmý Yokmu Kontrol Et.
dosya = Dir(hedef & Label1.Caption)
If dosya = "" Then
FileCopy kaynak & Label1.Caption, hedef & Label1.Caption
MsgBox "Kayýt Ýþlemi Tamamlandý.", vbInformation, "Tamamlandý."
Exit Sub
' Kopyalama Baþarý Ýle Tamamlandý.
Else
Dim soru, cevap
 soru = MsgBox("Bu Dosya Zaten Hedefinizde Mevcut,Bu Dosyayý Yeni Bir Ýsimle Kaydetmek Ýstemisiniz.?", vbQuestion + vbYesNo, "Uyarý")
If soru = vbYes Then
cevap = InputBox(" Yeni Dosya ismini Giriniz.! Mutlaka Uzantý belirtiniz (*.exe) gibi", "Yeni Dosya Adý", "Kopyasý" & Label1.Caption)
'Yeni Ýsim verildikten Sonra Tekrar Kontrol Et.Eðer Yine Warsa Tekrar yeni Ýsim Sor.
tekrar = Dir(hedef & cevap) 'Hedefte Bu Dosyadan Warmý.?
If tekrar = "" Then 'Eðer Dosya Yoksa...
FileCopy kaynak & Label1.Caption, hedef & cevap
MsgBox "Kayýt Ýþlemi Tamamlandý.", vbInformation, "Tamamlandý."
Else
MsgBox "Bu Dosya Ýsminde Baþka Bir Dosya Daha var.Lüten Yeni Bir Ýsim Belirleyin."
End If
MsgBox "Kopyalama Ýþlemi Yapýlamadý [ Ýptal Ettiniz. ]", vbInformation, "Bilgi."
End If
End If
Exit Sub
hata:
MsgBox "Yanlýþ Hedef Yolu Belirttiniz Yada,Kaynak Dosyanýzýn Yeri Doðru Deðil.Lütfen Kontrol Ediniz.[ Dosya Kopyalanamadý ]", vbCritical, "Hata"
Exit Sub
End If
End Sub
Private Sub Command6_Click()
Form1.Label15.Caption = ""
Form1.Label13.Caption = ""
Unload Me

End Sub
Private Sub Command7_Click()
'Tümünü Kopyalama
On Error GoTo hata
If Text1.Text = "" Then 'Hedef Boþ Ýse
MsgBox "Hedef Belirtmediniz.!", vbCritical, "Hata" 'Ýþlemi Baþlatma
Else
Dim kaynak, hedef, sonuc            'Deðiþken Tanýmlar
kaynak = programayarlari.Text2.Text   ' Kaynak
hedef = Text1.Text ' Hedef
'Tanýmlamalar Bitti.
'Kopyalama Ýþlemince Bi Sakýnca Yok.
'****************************Bu Bölüm ";"Arasýdnaki Dosyalarý Ayýrt Ediyor.
Dim se_next_to As String
Dim se_mail_to As String
Dim dosya
se_email_to = ek_kaydet.Label3.Caption  ' Okunacak Metin.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
dosya = Dir(hedef & se_next_to)
If dosya = "" Then 'Dosya Hedefte Yok iSe....
FileCopy kaynak & se_next_to, hedef & se_next_to
 'Dosya Adetine Göre Ýþlemi Tekrarla.

'Unload Me 'Ýþlem Bittikten Sonra Uygulamayý Sonlandýr.
Else
'Ýþlem Baþladý Ancak Ayný Dosyadan Hedef te war ise ne olacak.
'1- Dosya Ýsmi Deðiþecek
'2- Mevcut Dosyanýn Üzerine Yazýlacak.
'--------------------------------------------------------------
' Dosya Yolu Zaten war.War oland eðiþsin mi.?
Dim uyar
uyar = MsgBox(hedef & se_next_to & "zaten var." & vbCrLf & "var olan dosya deðiþsin mi.?", vbCritical + vbYesNoCancel, "Uyarý")
If uyar = vbYes Then
FileCopy kaynak & se_next_to, hedef & se_next_to
'
End If
End If
End If
Loop
Unload Me
End If
Exit Sub
hata:
MsgBox "Bilinmeyen Hata Oluþtu [ Tümünü Kaydet ]" & vbCrLf & " Muhtemelen Belirttiðiniz Hedefte Böyle Bir Dizin Ýsmi Yok.", vbCritical, "Hata"
Exit Sub

End Sub

Private Sub Label3_Change()
On Error GoTo hata
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = ek_kaydet.Label3.Caption  ' Okunacak Metin.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
ListView1.ListItems.Add , , se_next_to, 1, 1
End If
Loop
Exit Sub
hata:
MsgBox "Sistem Hatasý [ Dosya Listeye Eklenemedi. #1250 ] ", vbCritical, "Önemli Hata"
Exit Sub
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Label1.Caption = ListView1.ListItems.Item(ListView1.SelectedItem.Index).Text
End Sub



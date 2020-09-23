VERSION 5.00
Begin VB.Form kullanici_ekle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yeni Kullanýcý"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   Icon            =   "kullanici_ekle.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   5055
      Left            =   45
      TabIndex        =   37
      Top             =   7650
      Visible         =   0   'False
      Width           =   7485
      Begin VB.CommandButton Command7 
         Caption         =   "Tamam"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2880
         TabIndex        =   39
         Top             =   4545
         Width           =   1815
      End
      Begin VB.Label Label24 
         Caption         =   "[ TIKLA ]"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4005
         TabIndex        =   41
         Top             =   2295
         Width           =   1050
      End
      Begin VB.Label Label23 
         Caption         =   "Yeni yetki Paketi Satýn Almak istiyorum..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   40
         Top             =   2295
         Width           =   2940
      End
      Begin VB.Label Label22 
         Caption         =   "Mevcut Programýnýzýn Yetki Sýnýrlamasý Nedeni ile Yeni Kayýt Yapamazsýnýz.!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   990
         TabIndex        =   38
         Top             =   1755
         Width           =   5460
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "c"
      Height          =   465
      Left            =   3060
      TabIndex        =   36
      Top             =   6795
      Width           =   735
   End
   Begin VB.Frame Frame3 
      Caption         =   "Uyarý"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   45
      TabIndex        =   30
      Top             =   3600
      Width           =   3390
      Begin VB.Label Label15 
         Caption         =   $"kullanici_ekle.frx":06EA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   135
         TabIndex        =   31
         Top             =   225
         Width           =   3120
      End
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Aktif"
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
      Left            =   5490
      TabIndex        =   27
      Top             =   3015
      Value           =   1  'Checked
      Width           =   1905
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Sil"
      Enabled         =   0   'False
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
      Left            =   3015
      TabIndex        =   26
      Top             =   5535
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Deðiþikliði Kaydet"
      Enabled         =   0   'False
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
      Left            =   45
      TabIndex        =   25
      Top             =   5535
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Bul.."
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
      Left            =   1485
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5535
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Standart Yetki Seçenekleri"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   19
      Top             =   4680
      Value           =   1  'Checked
      Width           =   3660
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Normal Yetki Seçenekleri (Admin) Yetkili"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   18
      Top             =   4230
      Width           =   3660
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Programýn Tüm Özelliklerinden Yararlanabilir."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   17
      Top             =   3780
      Width           =   3660
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   1755
      TabIndex        =   15
      Top             =   3420
      Width           =   6000
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   2430
      Width           =   2940
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      TabIndex        =   13
      Top             =   2835
      Width           =   2940
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      MaxLength       =   12
      TabIndex        =   12
      Top             =   2025
      Width           =   2940
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      MaxLength       =   12
      TabIndex        =   11
      Top             =   1620
      Width           =   2940
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1215
      Width           =   2940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "K&apat"
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
      Left            =   5940
      TabIndex        =   2
      Top             =   5535
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kaydet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4410
      TabIndex        =   1
      Top             =   5535
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -180
      TabIndex        =   0
      Top             =   5400
      Width           =   7980
   End
   Begin VB.Shape Shape1 
      Height          =   1230
      Left            =   1665
      Top             =   6165
      Width           =   3975
   End
   Begin VB.Label Label19 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   4095
      TabIndex        =   35
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label Label18 
      BackColor       =   &H0080C0FF&
      Height          =   285
      Left            =   2295
      TabIndex        =   34
      Top             =   6795
      Width           =   330
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Kayýtlý Kullanýcý Adeti"
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
      Left            =   3780
      TabIndex        =   33
      Top             =   6435
      Width           =   1545
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Kullaným Limiti"
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
      Left            =   1755
      TabIndex        =   32
      Top             =   6435
      Width           =   1680
   End
   Begin VB.Label Label14 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   45
      TabIndex        =   29
      Top             =   900
      Width           =   4605
   End
   Begin VB.Label Label13 
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
      Left            =   45
      TabIndex        =   28
      Top             =   5130
      Width           =   3300
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1250"
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
      Left            =   6660
      TabIndex        =   23
      Top             =   855
      Width           =   780
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý Kayýt Numarasý:"
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
      Left            =   4770
      TabIndex        =   22
      Top             =   855
      Width           =   1770
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "(Max 12)"
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
      Left            =   4590
      TabIndex        =   21
      Top             =   2070
      Width           =   690
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "(Max 12)"
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
      Left            =   4590
      TabIndex        =   20
      Top             =   1665
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4635
      Picture         =   "kullanici_ekle.frx":07DD
      ToolTipText     =   "Yeni Bir Grup Eklemek için Týklayýn..."
      Top             =   2430
      Width           =   240
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Yetki Seçenekleri"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   16
      Top             =   3330
      Width           =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Açýklama /  Görevi"
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
      Left            =   180
      TabIndex        =   10
      Top             =   2880
      Width           =   1410
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup"
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
      Left            =   180
      TabIndex        =   9
      Top             =   2475
      Width           =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Þifre"
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
      Left            =   180
      TabIndex        =   8
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý Adý"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1665
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Adý Soyadý"
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
      Left            =   180
      TabIndex        =   6
      Top             =   1260
      Width           =   1050
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Yeni Kullanýcý"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   90
      TabIndex        =   4
      Top             =   450
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   6705
      Picture         =   "kullanici_ekle.frx":0D67
      Stretch         =   -1  'True
      Top             =   90
      Width           =   690
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   825
      Left            =   -45
      TabIndex        =   3
      Top             =   0
      Width           =   7620
   End
End
Attribute VB_Name = "kullanici_ekle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub kullanici_varmi()
'Kullanýcý Adýný Kontrol Et.
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from kullanicilar WHERE kullanici_adi ='" & Text2.Text & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
MsgBox "[" & Text2.Text & "] Daha Önce Kayýt Edilmiþ Lütfen Farklý bir Kullanýcý Adý Belirleyin.", vbCritical, "Güvenlik Uyarýsý"
Text2.Text = ""
Text2.SetFocus
suz.Close
Else
Text3.SetFocus
End If
End Sub
Private Sub Check1_Click()
If Check1.Value = 1 Then
Check2.Value = 0
Check3.Value = 0
Else
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Check1.Value = 0
Check3.Value = 0

End If
End Sub


Private Sub Check3_Click()
If Check3.Value = 1 Then
Check1.Value = 0
Check2.Value = 0
Else
End If
End Sub


Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.SetFocus
End If
End Sub


Private Sub Command1_Click()
If Text1.Text = "" Or Text1.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Combo1.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * from kullanicilar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew
        rs!id = Label12.Caption             ' KAYIT NUMARASI
        rs!adi_soyadi = Text1.Text          ' ADI SOYADI
        rs!kullanici_adi = Text2.Text       ' KULLANICI ADI
        rs!sifre = Text3.Text               ' SIFRE
        rs!grup = Combo1.Text               ' GRUP
        rs!aciklama = Text4.Text            ' ACIKLAMA / BOLUMU
        rs!yonetici = Check1.Value          ' YETKI - YONETICI
        rs!kullanici = Check2.Value         ' YETKI-KULLANICI
        rs!sinirla = Check3.Value           ' YETKI-SINIRLA
        rs!kayit_tarihi = Label13.Caption   ' KAYIT TARIHI
        rs!aktif = Check4.Value             ' AKTIFLIK DURUMU
        
rs.Update
rs.Close
MsgBox "Kayýt Baþarý ile Eklendi. - " & Text1.Text, vbInformation, "Eklendi"
liste_kullanicilar ' Ana Ekrandaki Kullanýcýlar Bölümü Hemen Yenilensin.
Dim soru
soru = MsgBox("Baþka Kullanýcý Ekleyecekmisiniz.?", vbYesNo + vbQuestion, "Bilgi")
If soru = vbYes Then
'Evet Dediyse
Text1.Text = "": Text2.Text = "": Text3.Text = "": Text4.Text = "": Text1.SetFocus
Else
Unload Me
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
'Kullanýcý Bilgilerini Deðiþtirmek için Kayýt Bul..
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
Dim kullanici 'Kullanýcý Adýna Göre Arama Yapýlacak
kullanici = InputBox("Bulmak istediðiniz Kullanýcýnýn Kullanýcý Adýný Giriniz.?", "Bul...")
suz.Open "Select * from kullanicilar WHERE kullanici_adi ='" & kullanici & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
Text1.Text = suz![adi_soyadi]
Text2.Text = suz![kullanici_adi]
Text3.Text = suz![sifre]
Combo1.Text = suz![grup]
Text4.Text = suz![aciklama]
Check1.Value = suz![yonetici]
Check2.Value = suz![kullanici]
Check3.Value = suz![sinirla]
Check4.Value = suz![aktif]
Label13.Caption = suz![kayit_tarihi]
suz.Close
Else
MsgBox "Böyle Bir Kullanýcý Bulunamadý.", vbCritical, "Hata Kodu (C100)"
Text1.SetFocus
End If
End Sub

Private Sub Command4_Click()
'On Error Resume Next
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from kullanicilar where kullanici_adi = '" & kullanici_ekle.Text2.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                            rs!adi_soyadi = Text1.Text
                            rs!kullanici_adi = Text2.Text
                            rs!sifre = Text3.Text
                            rs!grup = Combo1.Text
                            rs!aciklama = Text4.Text
                            rs!yonetici = Check1.Value
                            rs!kullanici = Check2.Value
                            rs!sinirla = Check3.Value
                            rs!kayit_tarihi = Label13.Caption
                            rs!aktif = Check4.Value
                                rs.Update
                                rs.Close
                                MsgBox "Kayýt Düzenlendi.", , "Düzenlendi."
                                
                                liste_kullanicilar ' Ana Ekrandaki Kullanýcýlar Bölümü Hemen Yenilensin.
                                
Else
      MsgBox "Kayýt Seçmediniz.Yada Yanlýþ Ýþlem Gerçekleþtirdiniz.Lütfen Daha Sonra Tekrar Deneyiniz.!", vbCritical, "Hata"
 End If
End Sub


Private Sub Command5_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Kayýt Seçmediniz! Yada Üzerinde Çalýþtýðýnýz Kayýt Bozuk.", vbCritical, "Hata"
Else
'Kullanýcý Kayýtlarý Silinemez Sadece Sistemde Görünmez Olur
'Update Ýþlemini Yap.
Dim soru
soru = MsgBox(Text1.Text & " Kayýtlý Kullanýcý Silinecek Eminmisiniz.?", vbInformation + vbYesNo, "Sil")
If soru = vbYes Then
Check4.Value = 0
Command4_Click
MsgBox "Kullancý Silindi.", vbInformation, "Tamam"
Form1.ListView2.Refresh
Unload Me
Else
MsgBox "Kullanýcý Silinemedi.", vbCritical, "Hata"
End If

End If
End Sub


Private Sub Command6_Click()
Dim a, b As String
a = Label18
b = Label19
If b > a Then
'MsgBox "Kullanýcý Limitiniz Dolmuþ.Lütfen Yetki Paketi Satýn alýn.", vbCritical, "Uyarý"
'End
Else
End If
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Load()
On Error GoTo hata
gruplar 'Kayýtlý Gruplarý Combo ya yukle
Label13.Caption = Date & " " & Time
'Kullanýcý Bilgilerini Görüntüle
'kullanici_ekle.Label18.Caption = program_hakkinda.Text1.Text
'kullanici_ekle.Label20.Caption = "1000" 'program_hakkinda.Text1.Text
kullanici_ekle.Label19.Caption = Form1.ListView2.ListItems.Count
Command6_Click
If kullanici_ekle.Label18.Caption = kullanici_ekle.Label19.Caption Or kullanici_ekle.Label18.Caption < kullanici_ekle.Label19.Caption Then
'Eðer Deðerler Eþit ise
kullanici_ekle.Frame4.Visible = True
Else
kullanici_ekle.Frame4.Visible = False
End If
Exit Sub
hata:
MsgBox "Sistemde Bilinmedik Hata Lütfen Üreticiye Bildirin." & vbCrLf & "Hata Meydana Gelen Birim Kullanýcý Ekle."
Exit Sub
End Sub
Private Sub Image2_Click()
grup_ekle.Show
End Sub

Private Sub Label24_Click()
MsgBox "www.axpirine.edu.tr.tc\satinal ", vbInformation, "Satýn Al"
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub
Private Sub Text2_Change()
Command4.Enabled = True
Command5.Enabled = True
'ilk harfi büyük yap.
If Len(Text2.Text) = 1 Then
Text2.Text = Format(Text2.Text, ">")
SendKeys "{End}"
Else
Text2.Text = Mid$(Text2.Text, 1, 1) + Mid$(Text2.Text, 2, 50)
End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
kullanici_varmi
End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Combo1.SetFocus
End If
End Sub



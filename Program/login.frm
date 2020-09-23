VERSION 5.00
Begin VB.Form Baglan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   5385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4515
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form2"
   Picture         =   "login.frx":0000
   ScaleHeight     =   5385
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   -180
      ScaleHeight     =   345
      ScaleWidth      =   6960
      TabIndex        =   4
      Top             =   5040
      Width           =   6990
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Ivedi Mesaj 1.0 Build 101 2005 Copyright (c)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   765
         TabIndex        =   29
         Top             =   45
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Kullanýcý Limit Konrtolü"
      Height          =   1050
      Left            =   6795
      TabIndex        =   23
      Top             =   3330
      Width           =   3210
      Begin VB.CommandButton Command1 
         Caption         =   "C"
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   27
         Top             =   675
         Width           =   555
      End
      Begin VB.Label Label22 
         Caption         =   "Kayýtlý Kullanýcý Sayýsý:"
         Height          =   240
         Left            =   45
         TabIndex        =   26
         Top             =   675
         Width           =   1545
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   25
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label20 
         Caption         =   "Programýn Kullaným Limiti:"
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
         Left            =   45
         TabIndex        =   24
         Top             =   315
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Kullanýcý Yetki Kontrol Paneli"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   6840
      TabIndex        =   16
      Top             =   4410
      Visible         =   0   'False
      Width           =   3165
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
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
         Left            =   1935
         TabIndex        =   22
         Top             =   1170
         Width           =   465
      End
      Begin VB.Label Label18 
         Caption         =   "Genel ( Yönetici)"
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
         Left            =   315
         TabIndex        =   21
         Top             =   1170
         Width           =   1230
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1935
         TabIndex        =   20
         Top             =   855
         Width           =   465
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1935
         TabIndex        =   19
         Top             =   540
         Width           =   465
      End
      Begin VB.Label Label15 
         Caption         =   "Yetki Seçeneði 2:"
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
         Left            =   315
         TabIndex        =   18
         Top             =   855
         Width           =   1365
      End
      Begin VB.Label Label14 
         Caption         =   "Yetki Seçeneði 1:"
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
         Left            =   315
         TabIndex        =   17
         Top             =   540
         Width           =   1365
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   630
      Top             =   2655
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   90
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4545
      Width           =   1545
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   90
      TabIndex        =   0
      Top             =   3960
      Width           =   1545
   End
   Begin VB.Label Label13 
      Caption         =   "Grubu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6795
      TabIndex        =   15
      Top             =   3015
      Width           =   2715
   End
   Begin VB.Label Label12 
      Caption         =   "Id"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6795
      TabIndex        =   14
      Top             =   2700
      Width           =   2715
   End
   Begin VB.Label Label11 
      Caption         =   "Aktif"
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
      Left            =   6795
      TabIndex        =   13
      Top             =   2430
      Width           =   2715
   End
   Begin VB.Label Label10 
      Caption         =   "kayit_tarihi"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6795
      TabIndex        =   12
      Top             =   2115
      Width           =   2715
   End
   Begin VB.Label Label9 
      Caption         =   "aciklama"
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
      Left            =   6795
      TabIndex        =   11
      Top             =   1845
      Width           =   2715
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
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
      Left            =   135
      TabIndex        =   10
      Top             =   45
      Width           =   4245
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   1665
      Picture         =   "login.frx":ABEF
      Top             =   4545
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   1980
      Picture         =   "login.frx":AF79
      Top             =   4545
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1980
      Picture         =   "login.frx":B303
      Top             =   4590
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1665
      Picture         =   "login.frx":B68D
      Top             =   4590
      Width           =   240
   End
   Begin VB.Label Label7 
      Caption         =   "sinirla"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6795
      TabIndex        =   9
      Top             =   1530
      Width           =   2715
   End
   Begin VB.Label Label6 
      Caption         =   "kullanici"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6795
      TabIndex        =   8
      Top             =   1215
      Width           =   2715
   End
   Begin VB.Label Label5 
      Caption         =   "yonetici"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6795
      TabIndex        =   7
      Top             =   900
      Width           =   2715
   End
   Begin VB.Label Label3 
      Caption         =   "þifre"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6795
      TabIndex        =   5
      Top             =   585
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   2475
      Left            =   135
      Picture         =   "login.frx":BA17
      Stretch         =   -1  'True
      Top             =   -135
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Þifre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   3
      Top             =   4275
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý Adý:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   3735
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      Height          =   5415
      Left            =   0
      Top             =   0
      Width           =   4515
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý  Seçilmedi"
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
      Height          =   240
      Left            =   810
      TabIndex        =   6
      Top             =   2475
      Width           =   2040
   End
End
Attribute VB_Name = "Baglan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub kullanici_sifre()
'On Error GoTo hata
'Kullanýcý Adýný Kontrol Et.
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from kullanicilar WHERE kullanici_adi ='" & Text1.Text & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
Label3.Caption = suz![sifre]
Label4.Caption = suz![adi_soyadi]
'Label5.Caption = suz![yonetici]
'Label6.Caption = suz![kullanici]
'Label7.Caption = suz![sinirla]
'Kullanýcý Kontrolü
Label19.Caption = suz![yonetici]
Label16.Caption = suz![kullanici]
Label17.Caption = suz![sinirla]
'Kontrol Bitti.
Label9.Caption = suz![aciklama]
Label10.Caption = suz![kayit_tarihi]
Label11.Caption = suz![aktif]
Label12.Caption = suz![id]
Label13.Caption = suz![grup]
suz.Close
Text2.SetFocus
Else
MsgBox "Yanlýþ Kullanýcý Adý Giridiniz.'", vbCritical, "Hata Kodu (C100)"
Text1.Text = ""
Text1.SetFocus
End If
'Exit Sub
'hata:
'MsgBox "Data Dosyasý Bulunamadý." & vbCrLf & vbCrLf & programayarlari.Text1.Text & vbCrLf & vbCrLf & " Adresi Hatalý." & " Ayarlar Bölümünden Doðru Ayarlarý Belirtin.", vbCritical, "Veri Ýletiþim Hatasý [ IVEDI MESSAGE 1.0beta]"
'End
'Exit Sub
End Sub

Private Sub Command1_Click()
Dim a, b As String
a = Label21
b = Label23
If b > a Then
MsgBox "Kullanýcý Limitiniz Dolmuþ.Dilerseniz Kullanýcý Limitinizi Arttýrabilirsiniz..!", vbCritical, "Uyarý"
End
Else
'MsgBox "VBCritical", vbCritical, "Uyarý"
End If
End Sub

Private Sub Form_DblClick()
If Label8.Caption = "" Then 'Eðer Yazý Yok ise Kapatma
Else
Unload Me
End If
End Sub

Private Sub Form_Load()
'Kayýt Kontrol iþlemi Baþlatýldý.
'Label21.Caption = program_hakkinda.Text1.Text
Label23.Caption = Form1.ListView2.ListItems.Count
'Command1_Click ' Kullanýcý Sýnýrlamasýný Kontrol Et.
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = False
Image5.Visible = False
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Visible = True
End Sub


Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Visible = True
End Sub


Private Sub Image4_Click()
End

End Sub


Private Sub Image5_Click()
If Text2.Text = Label3.Caption Then
Form1.Label8.Caption = Baglan.Label4.Caption
Form1.Label9.Caption = Label9.Caption 'görevi
Form1.kullanici_adi.Caption = Text1.Text 'Kullanýcý ADINI Gonder..
Hide 'Unload Me
Form1.Show
Else
MsgBox "Yanlýþ Þifre Giriþi Yaptýnýz.Lütfen Tekrar Deneyin.", vbCritical, "Hata Kodu (C200)"
Text2 = ""
Text2.SetFocus
End If
End Sub

Private Sub Label17_Change()
'Bu Bölümde Kullanýcý Sadece Kullanýcý Ekleyebilir Kullanýcý Þifrelerini Göremez Deðiþtiremez
If Baglan.Label17.Caption = "1" Then
'Kullanýcýnýn Bu Özelliði Yok ise
Form1.Toolbar1.Buttons(12).Visible = False 'Kullanýcý Ekle Bölümünü Görmesin.
Form1.Toolbar1.Buttons(13).Visible = False 'Kullanýcý Ekle Bölümünü Cizgisi Görmesin.
Else
'Eðer Kullanýcý Yetki Seceneklerini Kullanacak ise
End If
End Sub

Private Sub Label19_Change()
'Bu Bölümde Kullanýcý Tam Yetkili
If Baglan.Label19.Caption = "1" Then
'Kullanýcýnýn Bu Özelliði Yok ise
Form1.Toolbar1.Buttons(12).Visible = True 'Kullanýcý Ekle Bölümünü Görmesin.
Form1.Toolbar1.Buttons(13).Visible = True 'Kullanýcý Ekle Bölümünü Cizgisi Görmesin.
Form1.yonetim.Visible = True
Else
'Eðer Kullanýcý Yetki Seceneklerini Kullanacak ise
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'Text2.SetFocus
kullanici_sifre
End If
End Sub


Private Sub Text2_GotFocus()
kullanici_sifre
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
If Text2.Text = Label3.Caption Then
Form1.Label8.Caption = Baglan.Label4.Caption ' Form1 e kullanýcýnýn ismini göster
Form1.Label9.Caption = Label9.Caption 'görevi
Form1.kullanici_adi.Caption = Text1.Text 'Kullanýcý ADINI Gonder..
Hide 'Unload Me
Form1.Show
Else
MsgBox "Yanlýþ Þifre Giriþi Yaptýnýz.Lütfen Tekrar Deneyin.", vbCritical, "Hata Kodu (C200)"
Text2 = ""
Text2.SetFocus
End If
End If
End Sub



VERSION 5.00
Begin VB.Form programayarlari 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Ayarlarý"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9240
   Icon            =   "program_ayarlari.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   195
      Left            =   3150
      TabIndex        =   29
      Top             =   5265
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Frame Frame2 
      Caption         =   "Yardým"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1800
      TabIndex        =   22
      Top             =   135
      Visible         =   0   'False
      Width           =   7395
      Begin VB.Label Label9 
         Caption         =   "Program Hakkýnda Kullaným ve Kurulum bilgileri için yardým.pdf dosyasýna Bakabilirsiniz."
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
         Left            =   405
         TabIndex        =   23
         Top             =   675
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seçenekler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1800
      TabIndex        =   16
      Top             =   135
      Visible         =   0   'False
      Width           =   7395
      Begin VB.CommandButton Command7 
         Caption         =   "Kaydet"
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
         Left            =   6120
         TabIndex        =   26
         Top             =   4230
         Width           =   1140
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Gönderilen Mesajlarda Tüm Bilgilerim Gözüksün."
         Height          =   240
         Left            =   225
         TabIndex        =   19
         Top             =   450
         Value           =   1  'Checked
         Width           =   3795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mesajlarýmda Ýmza Ekim Gözüksün"
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
         TabIndex        =   18
         Top             =   3105
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox Text3 
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
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   3375
         Width           =   2940
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Program Ayarlarý"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   1800
      TabIndex        =   6
      Top             =   135
      Width           =   7395
      Begin VB.CommandButton Command9 
         Caption         =   "Otomatik Ayarla"
         Height          =   375
         Left            =   5715
         TabIndex        =   30
         Top             =   3780
         Width           =   1590
      End
      Begin VB.TextBox Text4 
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
         Left            =   1665
         TabIndex        =   27
         Top             =   1170
         Width           =   4875
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Gözat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6615
         TabIndex        =   15
         Top             =   810
         Width           =   555
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Gözat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6615
         TabIndex        =   14
         Top             =   405
         Width           =   555
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Kaydet"
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
         Left            =   5715
         TabIndex        =   11
         Top             =   4185
         Width           =   1590
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   240
         Left            =   7020
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   240
         Left            =   7020
         TabIndex        =   9
         Top             =   1845
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1665
         TabIndex        =   8
         Top             =   810
         Width           =   4875
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
         Height          =   285
         Left            =   1665
         TabIndex        =   7
         Top             =   405
         Width           =   4875
      End
      Begin VB.Label Label12 
         Caption         =   "Ýmza Dosyasý:"
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
         TabIndex        =   28
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label Label11 
         Caption         =   $"program_ayarlari.frx":06EA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   225
         TabIndex        =   25
         Top             =   3285
         Width           =   4875
      End
      Begin VB.Label Label10 
         Caption         =   $"program_ayarlari.frx":0817
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   180
         TabIndex        =   24
         Top             =   1980
         Width           =   4560
      End
      Begin VB.Label Label6 
         Caption         =   "Dosya Adresi:"
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
         TabIndex        =   13
         Top             =   855
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "DataBase Yolu:"
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
         TabIndex        =   12
         Top             =   450
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Kapat"
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
      Left            =   7695
      TabIndex        =   1
      Top             =   5130
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   60
      Left            =   1755
      TabIndex        =   0
      Top             =   4995
      Width           =   9015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      TabIndex        =   21
      Top             =   2790
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Yardým"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      TabIndex        =   20
      Top             =   2070
      Width           =   1500
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Seçenekler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Ayarlarý"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   90
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   810
      Width           =   1545
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ayarlar.ver.1.004"
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
      Left            =   135
      TabIndex        =   2
      Top             =   5310
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Height          =   8880
      Left            =   0
      TabIndex        =   3
      Top             =   -270
      Width           =   1725
   End
End
Attribute VB_Name = "programayarlari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MultiPage1_Change()

End Sub

Private Sub Command1_Click()
Command2_Click
Command8_Click
Command4_Click
End Sub

Private Sub Command2_Click()
On Error Resume Next
Open App.Path & "\dosya.yol" For Output As #1
Print #1, Text2.Text
Close #1
'MsgBox " Kayýtlar baþarý ile aktarýlmýþtýr..", vbInformation, "Tamamlandý"
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
Open App.Path & "\yol.yol" For Output As #1
Print #1, Text1.Text
Close #1
MsgBox " Ayarlarýnýz Kayýt Edildi..", vbInformation, "Tamamlandý"
End Sub

Private Sub Command5_Click()
Dim strResFolder As String
strResFolder = BrowseForFolder(hWnd, "Lütfen Herhangi bir Dizin Belirleyiniz.")
If strResFolder = "" Then
Else
Text1.Text = strResFolder & "\"
End If
End Sub

Private Sub Command6_Click()
Dim strResFolder As String
strResFolder = BrowseForFolder(hWnd, "Lütfen Herhangi bir Dizin Belirleyiniz.")
If strResFolder = "" Then
Else
Text2.Text = strResFolder & "\"
End If
End Sub


Private Sub Command7_Click()
On Error Resume Next
Open App.Path & "\ayarlar.spr" For Output As #1
Print #1, Text3.Text
Print #2, Check1.Value
Print #3, Check2.Value
Close #1
MsgBox " Ayarlarýnýz Kayýt Edildi..", vbInformation, "Tamamlandý"
End Sub

Private Sub Command8_Click()
On Error Resume Next
Open App.Path & "\imza.spr" For Output As #1
Print #1, Text4.Text
Close #1
'MsgBox " Kayýtlar baþarý ile aktarýlmýþtýr..", vbInformation, "Tamamlandý"
End Sub

Private Sub Command9_Click()
MsgBox " Bu Seçenek Sadece Tek Kullanýcý için Ayarlanabilir Özelliktir.Eðer Bu Programý Að Üzerinde Kullanacak Ýseniz Bu Özelliði Kullanmayýn.", vbInformation, "Dikkat"
Text1.Text = App.Path & "\Data\"
Text2.Text = App.Path & "\Dosya\"
Text4.Text = App.Path & "\imza\"
Command1_Click
Command1.Enabled = False
End Sub

Private Sub Form_Load()
On Error GoTo hata
' Program Ayarlarýný Yaptýktan Sonra Ayarlarý Servere Kaydetsin.
'DataBase Yol Kayýt Dosyasý
'**********************************************Program Ayarlarý 1
Open App.Path & "\yol.yol" For Input As #1
Line Input #1, yol
Text1 = yol
Close #1
'FileSend Yol Kayýt Dosyasý
'**********************************************Program Ayarlarý 2
Open App.Path & "\dosya.yol" For Input As #1
Line Input #1, yol
Text2 = yol
Close #1
'**********************************************Ýmza Ayarlarý 1
Open App.Path & "\imza.spr" For Input As #1
Line Input #1, yol
Text4 = yol
Close #1
'**********************************************Ýmza Ayarlarý 2
Open App.Path & "\ayarlar.spr" For Input As #1
Line Input #1, yol
Text3 = yol
Close #1
Exit Sub
hata:
MsgBox "DataBase veya Dosya Kayýt Dosyalarýndan biri Bulunamadý.Lütfen Porgramýn Kurulu Olduðu Dizinde yol.yol ve dosya.yol dosyalarýnýn olup olmadýðýný kontrol edin.", vbCritical, "Uyarý"
End 'Programý Açmadan Sonlandýr.
Exit Sub
End Sub

Private Sub Label1_Click()
Frame4.Visible = False
Frame1.Visible = True
Frame2.Visible = True
End Sub


Private Sub Label2_Click()
MsgBox "Ýnternet Baðlantýnýz Yok.", vbInformation, "Uyarý"
End Sub

Private Sub Label7_Click()
Frame4.Visible = True
Frame1.Visible = False
Frame2.Visible = False
End Sub


Private Sub Label8_Click()
Frame4.Visible = False
Frame1.Visible = True
Frame2.Visible = False
End Sub



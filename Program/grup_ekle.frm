VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form grup_ekle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grup Ekle"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "grup_ekle.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   945
      TabIndex        =   15
      Top             =   945
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   4185
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
            Picture         =   "grup_ekle.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Deðiþtir"
      Enabled         =   0   'False
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
      Left            =   3510
      TabIndex        =   12
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sil"
      Enabled         =   0   'False
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
      Left            =   2295
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Yeni Grup"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   4320
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2985
      Left            =   4320
      TabIndex        =   9
      Top             =   1035
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   5265
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Grup Adý"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kodu"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.TextBox Text2 
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
      Height          =   330
      Left            =   945
      TabIndex        =   7
      ToolTipText     =   "Bir Grup Adý Belirleyin Örn.( Muhasebe )"
      Top             =   1935
      Width           =   3300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "K&apat"
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
      Left            =   5940
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Kaydet"
      Enabled         =   0   'False
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
      Left            =   4725
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   -1035
      TabIndex        =   4
      Top             =   4050
      Width           =   8520
   End
   Begin VB.TextBox Text1 
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
      Height          =   330
      Left            =   945
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "Bir Grup Kodu Belirleyin. Örn: Muh ( Muhasebe )"
      Top             =   1440
      Width           =   1590
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup ID:"
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
      TabIndex        =   14
      Top             =   990
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mevcut Grup Listesi"
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
      Left            =   4365
      TabIndex        =   13
      Top             =   810
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup Adý"
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
      TabIndex        =   8
      Top             =   1980
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup Kodu"
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
      TabIndex        =   3
      Top             =   1485
      Width           =   825
   End
   Begin VB.Image Image1 
      Height          =   555
      Left            =   6345
      Picture         =   "grup_ekle.frx":0C84
      Stretch         =   -1  'True
      Top             =   90
      Width           =   600
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grup Ýþlemleri"
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
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   3165
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   780
      Left            =   -90
      TabIndex        =   0
      Top             =   0
      Width           =   7350
   End
End
Attribute VB_Name = "grup_ekle"
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
suz.Open "Select * from gruplar WHERE id ='" & Text3.Text & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
MsgBox "[" & Text3.Text & "] Daha Önce Kayýt Edilmiþ Lütfen Farklý bir ID Numarasý Belirleyin.", vbCritical, "Güvenlik Uyarýsý"
Text3.Text = ""
Text3.SetFocus
suz.Close
Else
Text1.SetFocus
End If
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open "Select * from gruplar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew
        rs!id = Text3.Text                 ' ID
        rs!grupkodu = Text1.Text           ' GRUP KODU
        rs!grupadi = Text2.Text            ' GRUP ADI
rs.Update
rs.Close
MsgBox "Kayýt Baþarý ile Eklendi. - " & " [ " & Text1.Text & "-" & Text2.Text & "]", vbInformation, "Eklendi"
grup_listesi
Dim soru
soru = MsgBox("Baþka Kayýt Ekleyecekmisiniz.?", vbYesNo + vbQuestion, "Bilgi")
If soru = vbYes Then
'Evet Dediyse
Text1.Text = "": Text2.Text = "": Text3.Text = "": Text1.SetFocus
Else
Unload Me
End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Command1.Enabled = True
'Command4.Enabled = True
'Command5.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text3.SetFocus
Command3.Enabled = False 'Yeni Grup Butonunu Gizle

End Sub

Private Sub Command4_Click()
MsgBox "Kayýt Edilen Gruplar Silinemez." & vbCrLf & "Bu Grup Ýþlem Görmüþ."
End Sub

Private Sub Command5_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text3.SetFocus
If Command5.Caption = "Deðiþtir" Then
Command5.Caption = "Onayla" ' Onayla Olarak Deðiþtir.
Else
'Kayýt Kodlarý
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from gruplar where id = '" & Text3.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
rs!grupkodu = Text1.Text
rs!grupadi = Text2.Text
rs.Update
rs.Close
MsgBox "Kayýt Düzenlendi.", , "Düzenlendi."
grup_listesi
Else
MsgBox "Kayýt Seçmediniz.Yada Yanlýþ Ýþlem Gerçekleþtirdiniz.Lütfen Daha Sonra Tekrar Deneyiniz.!", vbCritical, "Hata"
End If
'Onayla Yazmýyorsa Eski Haline Geç
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Command5.Enabled = False
Command5.Caption = "Deðiþtir"
End If
End Sub

Private Sub Form_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Command5.Enabled = False
End Sub

Private Sub Form_Load()
grup_listesi
End Sub

Private Sub ListView1_Click()
Text3.Text = ListView1.SelectedItem.SubItems(1)
grup_kontrol_sorgusu
Command5.Enabled = True
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If

End Sub


Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
kullanici_varmi 'Ayný ID Giriþini Yapmamak Ýçin Kontrol Et.
End If
End Sub

Private Sub grup_kontrol_sorgusu()
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from gruplar WHERE id ='" & Text3.Text & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
Text1.Text = suz![grupkodu]
Text2.Text = suz![grupadi]
suz.Close
Else
End If
End Sub

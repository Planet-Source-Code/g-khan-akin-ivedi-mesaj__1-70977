VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form mesaji_ilet 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13065
   Icon            =   "ilet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8038.381
   ScaleMode       =   0  'User
   ScaleWidth      =   17849.37
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   540
      Top             =   7290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   135
      Top             =   7335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
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
      Height          =   330
      Left            =   540
      TabIndex        =   10
      Top             =   1935
      Visible         =   0   'False
      Width           =   12435
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4875
      Left            =   45
      TabIndex        =   6
      Top             =   2970
      Width           =   12930
      _ExtentX        =   22807
      _ExtentY        =   8599
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"ilet.frx":06EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   45
      TabIndex        =   5
      Top             =   2295
      Width           =   12930
      Begin VB.CommandButton Command3 
         Caption         =   "eklenti"
         Height          =   285
         Left            =   9810
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "bilgi"
         Height          =   285
         Left            =   8910
         TabIndex        =   17
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton Command1 
         Caption         =   "gönder"
         Height          =   285
         Left            =   7920
         TabIndex        =   16
         Top             =   225
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Yanýtlayýn"
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
         Left            =   5130
         TabIndex        =   15
         Top             =   225
         Width           =   1230
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Bilgilendirme"
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
         Left            =   3150
         TabIndex        =   14
         Top             =   225
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ýncele"
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
         Left            =   1710
         TabIndex        =   13
         Top             =   225
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Acil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   12
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "DN"
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
         Left            =   12060
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "MN"
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
         Left            =   11250
         TabIndex        =   19
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
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
      Left            =   540
      TabIndex        =   4
      Top             =   1530
      Width           =   12480
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
      Left            =   540
      TabIndex        =   3
      ToolTipText     =   "Bilgi Gönderm Yapacaðýnýz Kiþilerin adýný yazdýktan sonra ; koymayý unutmayýn."
      Top             =   1125
      Width           =   12480
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
      Left            =   540
      TabIndex        =   2
      Top             =   720
      Width           =   12480
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -180
      TabIndex        =   1
      Top             =   585
      Width           =   13740
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   1005
      ButtonWidth     =   1905
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Gönder     "
            Key             =   "a"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Ek     "
            Key             =   "b"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Kapat     "
            Key             =   "c"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ilet.frx":0769
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ilet.frx":0BBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ilet.frx":0D15
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      Caption         =   "tarih hatasý"
      Height          =   240
      Left            =   90
      TabIndex        =   21
      Top             =   8010
      Width           =   3120
   End
   Begin VB.Label Label4 
      Caption         =   "Ek:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   11
      Top             =   1980
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
      Caption         =   "Konu:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   1575
      Width           =   465
   End
   Begin VB.Label Label2 
      Caption         =   "Bilgi:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   8
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Kime:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   7
      Top             =   765
      Width           =   465
   End
End
Attribute VB_Name = "mesaji_ilet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo hata
If Text1.Text = "" Or Text3.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!" & vbCrLf & "Kime ; Konu ; Mesaj ;", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
'dosya_gonder ' Bu Bölüm iptal edildi.
'Önce Dosya Transfer Ýþlemini Baþlatmamýz Gerekiyor.
'Eðer Dosya Kopyalamada Hata Meydana Gelirse Böylelikle Mesaj Defalarca
'Kayýt Edilmez........
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = Text1.Text ' Deðiþken Satýr.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then

Label5.Caption = Label5.Caption + 1 ' Mesaj Numarasý

rs.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew ' 1 DEN 16 YA KADAR GERI SAYIM ISLEMINCE BILINMEYEN BIR DENKLEM
                    rs!id = Label5.Caption
                    rs!mesajid = Label6.Caption
                    rs!kimden = Baglan.Text1.Text ' ilk acýlýstaki kullanýcý adý
                    rs!txtbilgi = "(" & Baglan.Label4.Caption & ")"
                    rs!kime = se_next_to
                    rs!konu = Text3.Text
                    rs!gonderim_tarihi = Label7.Caption
                    rs!okundu = "0"
                    rs!silindi = "0"
                    rs!gonderilen = "1" 'Eðer Sürekli Sen Gönderiyosan Gönderildi Olur.
                    rs!gonderilmedi = "0"
                    rs!acil = Check1.Value
                    rs!inceleyin = Check2.Value
                    rs!bilgilendirme = Check3.Value
                    rs!yanitlayin = Check4.Value
                    rs!okundu_tarih = "okunmadý"
                    rs!mesaj = RichTextBox1.Text
                    If Text4.Text = "" Then
                    rs!Ek = "0" 'Eðer Ek Yoksa 0 Deðeri
                    rs!atac = "0" 'Dosyada yok gibi sýfýr deðeri ver.
                    Else
                    rs!Ek = "1" 'Eðer Ek Warsa 1 Deðeri
                    rs!atac = Text4.Text
                    End If
rs.Update
rs.Close
End If
Loop
MsgBox " '' " & Text3.Text & " ''" & vbCrLf & vbCrLf & "    Mesajýnýz Baþarý ile Gönderildi...", vbInformation, "Tamamlandý."
Command2_Click ' Bilgi Gödnermek Ýstediðiniz Kiþilere Gidiyor.
gelen_mesajlar
okunmamis_mesajlar
sol_menu
Unload Me
End If
Exit Sub
hata:
MsgBox "Bilinmeyen Bir Sistem Hatasý Meydana Geldi.Lütfen Sistem Yöneticinize Baþvurun." & vbCrLf & "Çok Fazla Eklenti Dosyasý Eklenmiþ Olabilir.", vbCritical, "Hata"
Exit Sub
End Sub

Private Sub Command2_Click()
'On Error GoTo hata
If Text1.Text = "" Or Text3.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!" & vbCrLf & "Kime ; Konu ; Mesaj ;", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = Text2.Text ' Deðiþken Satýr.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
Label5.Caption = Label5.Caption + 1 ' Mesaj Numarasý
rs.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew ' 1 DEN 16 YA KADAR GERI SAYIM ISLEMINCE BILINMEYEN BIR DENKLEM
                              rs!id = Label5.Caption
                    rs!mesajid = Label6.Caption
                    rs!kimden = Baglan.Text1.Text ' ilk acýlýstaki kullanýcý adý
                    rs!txtbilgi = "(" & Baglan.Label4.Caption & ")"
                    rs!kime = se_next_to
                    rs!konu = "Ýletildi Raporu" 'Text3.Text
                    rs!gonderim_tarihi = Label7.Caption
                    rs!okundu = "0"
                    rs!silindi = "0"
                    rs!gonderilen = "1" 'Eðer Sürekli Sen Gönderiyosan Gönderildi Olur.
                    rs!gonderilmedi = "0"
                    'rs!acil = Check1.Value
                    'rs!inceleyin = Check2.Value
                    'rs!bilgilendirme = Check3.Value
                    'rs!yanitlayin = Check4.Value
                    rs!okundu_tarih = "okunmadý"
                  ' rs!mesaj = RichTextBox1.Text
                  ' rs!atac = Text4.Text
                  ' If Text4.Text = "" Then
                  ' rs!ek = "0" 'Eðer Ek Yoksa 0 Deðeri
                  ' Else
                  ' rs!ek = "1" 'Eðer Ek Warsa 1 Deðeri
                  ' End If
                    rs!mesaj = Text3.Text & " Konulu Mesaj Gitmesi Gereken Birimlere Ulaþtý." & vbCrLf & "Mesajý Gönderen: " & Baglan.Text1.Text & " ( " & Baglan.Label4.Caption & " ) " & vbCrLf & "Tarih : " & Date & " " & Time
rs.Update
rs.Close
End If
Loop
'MsgBox " '' " & Text3.Text & " ''" & vbCrLf & vbCrLf & "    Mesajýnýz Baþarý ile Gönderildi...", vbInformation, "Tamamlandý."
gelen_mesajlar
okunmamis_mesajlar
Unload Me
End If
'Exit Sub
'hata:
'MsgBox "Bilinmeyen Bir Sistem Hatasý Meydana Geldi.Lütfen Sistem Yöneticinize Baþvurun.", vbCritical, "Hata"
'Exit Sub
End Sub

Private Sub Command3_Click()
cmd1.CancelError = True
On Error GoTo hata
cmd1.Action = 1
Text4.Text = Text4.Text + cmd1.FileTitle & ";"
mesaji_ilet.Text4.Visible = True
mesaji_ilet.Label4.Visible = True
Exit Sub
hata:
Exit Sub
End Sub


Private Sub Form_Load()
kontor2
'mesaji_ilet.Caption = yeni_mesaj_gonder.Label16.Caption & " Nolu Mesaj Cevaplanýyor..."
Label7.Caption = Date & " " & Time
End Sub

Private Sub Label4_Click()
Command3_Click
End Sub

Private Sub Text4_DblClick()
mesaji_ilet.Text4.Visible = False
mesaji_ilet.Label4.Visible = False
End Sub


Private Sub Timer1_Timer()
mesaji_ilet.RichTextBox1.SetFocus
Timer1.Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "a"
Command1_Click
'MsgBox "gÖNDER"
Case "b"
Command3_Click
'MsgBox "Eklenti."
Case "c"
Unload Me
End Select
End Sub

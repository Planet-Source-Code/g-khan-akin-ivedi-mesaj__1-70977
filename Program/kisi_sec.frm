VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form kisi_sec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kiþiler"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "kisi_sec.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "kisi_sec.frx":06EA
      Left            =   45
      List            =   "kisi_sec.frx":06EC
      TabIndex        =   12
      Text            =   "Tümü"
      Top             =   810
      Width           =   2940
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -675
      TabIndex        =   11
      Top             =   6750
      Width           =   8700
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2925
      Top             =   3375
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
            Picture         =   "kisi_sec.frx":06EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Tümünü Seç.."
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
      Left            =   45
      TabIndex        =   10
      Top             =   1170
      Width           =   3030
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Ýptal"
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
      Left            =   3060
      TabIndex        =   9
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Tamam"
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
      Left            =   4275
      TabIndex        =   8
      Top             =   6840
      Width           =   1185
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3060
      TabIndex        =   5
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3060
      TabIndex        =   4
      Top             =   2340
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Left            =   45
      TabIndex        =   3
      Top             =   6435
      Width           =   3030
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   3465
      TabIndex        =   2
      ToolTipText     =   "Bilgi Gönderielcek Kiþiler"
      Top             =   4140
      Width           =   1950
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   3465
      TabIndex        =   1
      ToolTipText     =   "Mesaj Gönderilecek Kiþiler"
      Top             =   1125
      Width           =   1950
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4920
      Left            =   45
      TabIndex        =   0
      Top             =   1440
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   8678
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16761024
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kullanýcý"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kullanýcý"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   600
      Left            =   4725
      Picture         =   "kisi_sec.frx":0C88
      Stretch         =   -1  'True
      Top             =   135
      Width           =   510
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Kiþi Listesi"
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
      Left            =   90
      TabIndex        =   14
      Top             =   360
      Width           =   2580
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Height          =   825
      Left            =   -135
      TabIndex        =   13
      Top             =   -45
      Width           =   6360
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   45
      Picture         =   "kisi_sec.frx":1212
      Stretch         =   -1  'True
      ToolTipText     =   "Seçtiðiniz Kiþileri Listenizden Silmek Ýçin ; Silinecek Kiþinin Üzerinde Çift Týklayýn..."
      Top             =   6840
      Width           =   360
   End
   Begin VB.Label Label2 
      Caption         =   "Bilgi Mesajý Gidecekler:"
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
      Left            =   3465
      TabIndex        =   7
      Top             =   3870
      Width           =   1905
   End
   Begin VB.Label Label1 
      Caption         =   "Mesaj Gönderilecekler:"
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
      Left            =   3510
      TabIndex        =   6
      Top             =   855
      Width           =   1725
   End
End
Attribute VB_Name = "kisi_sec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
'Seçili Olursa Tümünü Seç
Dim i
For i = 1 To ListView1.ListItems.Count
ListView1.ListItems.Item(i).Checked = True
Next i
Else
Dim ii
For ii = 1 To ListView1.ListItems.Count
ListView1.ListItems.Item(ii).Checked = False
Next ii
'Seçimleri Ýptal Et.
End If
End Sub

Private Sub Combo1_Click()
If Combo1.Text = "Tümü" Then 'Tümü Seçeneði Seçildi Ýse
liste_kullanicilar1
Else
liste_kullanicilar2
End If
End Sub


Private Sub Command1_Click()
 'Mesaj Gönderilecek Kiþilerin Listesi
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
               selFlg = True
            Exit For
        End If
    Next
    
    If selFlg = False Then
        MsgBox "Lütfen Listenize Eklenecek Kiþiyi Seçin.", vbInformation, "Hata"
      ListView1.SetFocus
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
           List1.AddItem (ListView1.ListItems(i).ListSubItems(1))
            ListView1.ListItems(i).Checked = False
              
        End If
    Next
    selFlg = False
    Check1.Value = 0
End Sub

Private Sub Command2_Click()
'Bilgi Mesajý Gödnerilecek Kiþilerin Listesi
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
           selFlg = True
            Exit For
        End If
    Next
    
    If selFlg = False Then
        MsgBox "Lütfen Listenize Eklenecek Kiþiyi Seçin.", vbInformation, "Hata"
      ListView1.SetFocus
        Exit Sub
    End If
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Checked = True Then
           List2.AddItem (ListView1.ListItems(i).ListSubItems(1))
            ListView1.ListItems(i).Checked = False
          
        End If
    Next
    selFlg = False
      Check1.Value = 0
End Sub


Private Sub Command3_Click()
    Dim ToList As String
    Dim CcList As String
    Dim BccList As String
    If Len(Trim(yeni_mesaj_gonder.Text1.Text)) > 4 Then
        ToList = yeni_mesaj_gonder.Text1.Text & ";"
    End If
    If Len(Trim(yeni_mesaj_gonder.Text2.Text)) > 4 Then
        CcList = yeni_mesaj_gonder.Text2.Text & ";"
    End If
     If List1.ListCount = 0 And List2.ListCount = 0 Then
        MsgBox "Kiþi Listeniz Boþ", vbInformation, "Hata"
       ListView1.SetFocus
        Exit Sub
    End If
    For i = 0 To List1.ListCount - 1
       If InStr(1, yeni_mesaj_gonder.Text1.Text, List1.List(i)) = 0 Then
         ToList = ToList & List1.List(i) & ";"
       End If
    Next
    For i = 0 To List2.ListCount - 1
        If InStr(1, yeni_mesaj_gonder.Text2.Text, List2.List(i)) = 0 Then
            CcList = CcList & List2.List(i) & ";"
        End If
    Next
      
    If Len(ToList) > 2 Then
       yeni_mesaj_gonder.Text1.Text = Mid(ToList, 1, Len(ToList) - 1) & ";"
    End If
    If Len(CcList) > 2 Then
       yeni_mesaj_gonder.Text2.Text = Mid(CcList, 1, Len(CcList) - 1) & ";"
    End If
    Unload Me
End Sub

Private Sub Command4_Click()
Unload Me

End Sub
Private Sub Form_Load()
liste_kullanicilar1
secilen_gruplar
End Sub
Private Sub List1_DblClick()
'Çift Týklanan Öðeyi Sil.
List1.RemoveItem (List1.ListIndex)
End Sub
Private Sub List2_DblClick()
'Çift Týklanan Öðeyi Sil.
List2.RemoveItem (List2.ListIndex)
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Dim baslik As Integer
Dim bul As String
bul = Text1.Text 'InputBox("Aranak Kiþi: " & Adi, "Arama")
'baslik = lvwTex 'Ýl Kayýttaki Bilgileri Arama için Kullanabilirsin.lvwPartial
baslik = lvwSubItem 'Alt menülerde ara
Dim altmenu As ListItem
Set altmenu = ListView1.FindItem(bul, baslik, , lvwTex)
If altmenu Is Nothing Then
MsgBox "Böyle Bir Kayýt Yok.'" & vbCrLf, vbInformation + vbOKOnly, "Bulamadýk Be Gülüm"
Exit Sub
Else
altmenu.EnsureVisible
altmenu.Selected = True
ListView1.SetFocus
End If
End If
End Sub



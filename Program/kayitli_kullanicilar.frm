VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form kullanici_listesi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kayýtlý Kullanýcýlar"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   Icon            =   "kayitli_kullanicilar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Kapat"
      Height          =   375
      Left            =   7650
      TabIndex        =   1
      Top             =   7560
      Width           =   1590
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6585
      Left            =   45
      TabIndex        =   0
      Top             =   945
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   11615
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Adý Soyadý"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kullanýcý Adý"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Þifresi"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Kayýt Tarihi"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Grubu"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Bölümü"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   7650
      Width           =   7350
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   270
      Top             =   7650
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "kayitli_kullanicilar.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "KAYITLI KULLANICILAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5625
      TabIndex        =   3
      Top             =   405
      Width           =   3570
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   315
      Picture         =   "kayitli_kullanicilar.frx":08C4
      Stretch         =   -1  'True
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   915
      Left            =   -720
      TabIndex        =   2
      Top             =   -135
      Width           =   10230
   End
End
Attribute VB_Name = "kullanici_listesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Form_Load()
kayitli_kullanicilar_listesi
Label3.Caption = "Sistemde " & kullanici_listesi.ListView1.ListItems.Count & " Kullanýcý Mevcut."
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 Dim Orden As Integer
    ListView1.SortKey = ColumnHeader.Index - 1
    Orden = ListView1.SortKey
    ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
    ListView1.Sorted = True
    
End Sub



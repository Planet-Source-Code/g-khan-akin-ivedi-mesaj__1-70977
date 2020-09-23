VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "IVEDI Mesaj 1.00"
   ClientHeight    =   7395
   ClientLeft      =   2040
   ClientTop       =   2640
   ClientWidth     =   11385
   Icon            =   "FrmClient.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   11385
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      BorderStyle     =   0  'None
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
      Left            =   3195
      Locked          =   -1  'True
      TabIndex        =   57
      Text            =   "Bu Bölümde Görüntülenecek Mesaj Yok."
      Top             =   1530
      Visible         =   0   'False
      Width           =   4425
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Left            =   1350
      TabIndex        =   55
      Top             =   2250
      Width           =   555
   End
   Begin MSComctlLib.ListView arsiv 
      Height          =   285
      Left            =   9225
      TabIndex        =   54
      Top             =   1395
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8438015
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.TextBox Text6 
      Height          =   690
      Left            =   270
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   50
      Text            =   "FrmClient.frx":06EA
      Top             =   5670
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.Frame Frame4 
      Caption         =   "Bul"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3825
      TabIndex        =   45
      Top             =   4815
      Visible         =   0   'False
      Width           =   5055
      Begin VB.Frame Frame5 
         Height          =   60
         Left            =   45
         TabIndex        =   49
         Top             =   1530
         Width           =   4965
      End
      Begin VB.CommandButton Command5 
         Height          =   510
         Left            =   4275
         Picture         =   "FrmClient.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Bul..."
         Top             =   1620
         Width           =   645
      End
      Begin VB.TextBox Text5 
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
         Height          =   285
         Left            =   135
         TabIndex        =   46
         Top             =   1710
         Width           =   4065
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   360
         Picture         =   "FrmClient.frx":1550
         Top             =   360
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   105
         Left            =   4860
         Picture         =   "FrmClient.frx":1ADA
         Top             =   180
         Width           =   120
      End
      Begin VB.Label Label24 
         Caption         =   "Bu Bölüme Mesajýn Kimden Geldiðini Yazmalaýsýnýz? Tam Olarak Metni Belirmeye Çalýþýn."
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
         Left            =   945
         TabIndex        =   48
         Top             =   225
         Width           =   3840
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Dosya Kaydet"
      Height          =   330
      Left            =   9450
      TabIndex        =   43
      Top             =   6660
      Visible         =   0   'False
      Width           =   1725
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   9450
      Top             =   5715
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer3 
      Interval        =   20000
      Left            =   10530
      Top             =   3600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10125
      Top             =   3600
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2115
      ScaleHeight     =   240
      ScaleWidth      =   510
      TabIndex        =   40
      Top             =   1530
      Width           =   510
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Kontrol"
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
      Left            =   360
      TabIndex        =   24
      Top             =   3690
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton Command6 
         Caption         =   "arþiv"
         Height          =   240
         Left            =   45
         TabIndex        =   53
         Top             =   2835
         Width           =   2085
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1125
         TabIndex        =   51
         Top             =   1665
         Width           =   960
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Okunmad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   42
         Top             =   2520
         Width           =   825
      End
      Begin VB.TextBox Text4 
         Height          =   330
         Left            =   1125
         TabIndex        =   34
         Top             =   1305
         Width           =   960
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Okudu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         TabIndex        =   32
         Top             =   2520
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Sil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         TabIndex        =   31
         Top             =   2520
         Width           =   465
      End
      Begin VB.TextBox Text3 
         Height          =   330
         Left            =   1125
         TabIndex        =   30
         Top             =   945
         Width           =   960
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
         Height          =   330
         Left            =   1620
         TabIndex        =   27
         Top             =   585
         Width           =   465
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
         Left            =   1125
         TabIndex        =   25
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label21 
         Caption         =   "0"
         Height          =   240
         Left            =   1170
         TabIndex        =   58
         Top             =   2205
         Width           =   870
      End
      Begin VB.Label Label20 
         Caption         =   "Hatalý Mesaj"
         Height          =   285
         Left            =   1170
         TabIndex        =   56
         Top             =   1980
         Width           =   870
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Arþiv"
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
         TabIndex        =   52
         Top             =   1710
         Width           =   825
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Height          =   285
         Left            =   1125
         TabIndex        =   38
         Top             =   585
         Width           =   465
      End
      Begin VB.Label Label14 
         Caption         =   "Okundu.Tarih:"
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
         TabIndex        =   33
         Top             =   1350
         Width           =   1050
      End
      Begin VB.Label Label13 
         Caption         =   "Silindi.!"
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
         TabIndex        =   29
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label12 
         Caption         =   "Okundu Mu?"
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
         TabIndex        =   28
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "Message ID:"
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
         TabIndex        =   26
         Top             =   270
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList ImageList6 
      Left            =   3600
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":1F78
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   2970
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":2292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":26E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":2B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":2F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":33DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":36F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":3A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":3B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":3CC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":4114
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":46AE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   2295
      Top             =   6615
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":4C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":5192
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":56DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":5C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6170
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList kullanicilar 
      Left            =   0
      Top             =   6660
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
            Picture         =   "FrmClient.frx":65C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame dikey 
      BorderStyle     =   0  'None
      Height          =   5790
      Left            =   2880
      MousePointer    =   9  'Size W E
      TabIndex        =   4
      Top             =   1305
      Width           =   45
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1905
      Left            =   45
      TabIndex        =   18
      Top             =   4725
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   3360
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      Icons           =   "kullanicilar"
      SmallIcons      =   "kullanicilar"
      ColHdrIcons     =   "kullanicilar"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
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
         Object.Width           =   5028
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Kodu"
         Object.Width           =   2
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   1755
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6B5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":6CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":7250
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":76A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":77FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -45
      TabIndex        =   16
      Top             =   540
      Width           =   15435
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1170
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":7956
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":7DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":81FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":864C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":8A9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":8EF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   585
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":904A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":91A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":95F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":9A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":9BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":A9F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":AD8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":B328
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":B77A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmClient.frx":D5FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame yatay 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000B&
      Height          =   75
      Left            =   2835
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   4950
      Width           =   6885
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   1005
      ButtonWidth     =   1905
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Posta Oluþtur"
            Key             =   "a"
            Object.ToolTipText     =   "Yeni Posta"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Yanýtla"
            Key             =   "b"
            Object.ToolTipText     =   "Yanýtla"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ýlet"
            Key             =   "c"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sil"
            Key             =   "d"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   4
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gönder Al"
            Key             =   "e"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Bul"
            Key             =   "f"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kullanýcýlar"
            Key             =   "g"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "a"
                  Text            =   "Kullanýcý Bilgileri"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "b"
                  Text            =   "Kullanýcý Deðiþtir"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "c"
                  Text            =   "Kullanýcý Ekle"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "d"
                  Text            =   "Kullanýcý Düzenle"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Yardým"
            Key             =   "h"
            ImageIndex      =   10
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "e"
                  Text            =   "Yardým Konularý"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "f"
                  Text            =   "Program Hakkýnda"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "g"
                  Text            =   "Versiyon"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "h"
                  Text            =   "Update"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kapat"
            Key             =   "i"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture2 
         Height          =   330
         Left            =   11430
         ScaleHeight     =   270
         ScaleWidth      =   3015
         TabIndex        =   36
         Top             =   135
         Visible         =   0   'False
         Width           =   3075
         Begin VB.Label Label15 
            Height          =   240
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   2985
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   14535
         Picture         =   "FrmClient.frx":E44E
         ScaleHeight     =   375
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   45
         Width           =   465
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   7050
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17711
            MinWidth        =   17711
            Text            =   "Hazýr..."
            TextSave        =   "Hazýr..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Çevirimiçi Çalýþýyor"
            TextSave        =   "Çevirimiçi Çalýþýyor"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   5001
            MinWidth        =   5009
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial TUR"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2985
      Left            =   2970
      TabIndex        =   0
      Top             =   1035
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5265
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      OLEDragMode     =   1
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList5"
      SmallIcons      =   "ImageList5"
      ColHdrIcons     =   "ImageList3"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial TUR"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   776
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2
         ImageIndex      =   3
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Width           =   2
         ImageIndex      =   4
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "MesajNo"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Kimden"
         Object.Width           =   4304
         ImageIndex      =   5
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Konu"
         Object.Width           =   7303
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Gönderi Tarihi"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Okundu Tarihi"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "okundu"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "silindi"
         Object.Width           =   2
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "id"
         Object.Width           =   2
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1965
      Left            =   3195
      TabIndex        =   2
      Top             =   4860
      Width           =   6150
      _ExtentX        =   10848
      _ExtentY        =   3466
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"FrmClient.frx":EB38
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView Treeview1 
      Height          =   3660
      Left            =   45
      TabIndex        =   7
      Top             =   1035
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   6456
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList4"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   3195
      TabIndex        =   9
      Top             =   4185
      Width           =   12030
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   570
         Left            =   11520
         TabIndex        =   23
         Top             =   90
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList6"
         DisabledImageList=   "ImageList6"
         HotImageList    =   "ImageList6"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   5
            EndProperty
         EndProperty
      End
      Begin VB.Label Label23 
         Caption         =   "Label23"
         Height          =   240
         Left            =   2430
         TabIndex        =   44
         Top             =   225
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.Label Label10 
         Height          =   240
         Left            =   3825
         TabIndex        =   21
         ToolTipText     =   "'Mesaj ID Numarasýna Göre Sorgu"
         Top             =   225
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   9135
         TabIndex        =   15
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarih:"
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
         Left            =   8595
         TabIndex        =   14
         Top             =   135
         Width           =   555
      End
      Begin VB.Label Label6 
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
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   855
         TabIndex        =   13
         Top             =   405
         Width           =   600
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   855
         TabIndex        =   12
         Top             =   135
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Konu:"
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
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Kimden:"
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
         TabIndex        =   10
         Top             =   135
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView okunmamis 
      Height          =   285
      Left            =   9225
      TabIndex        =   22
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   503
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16744576
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2970
      TabIndex        =   39
      Top             =   720
      Width           =   3390
   End
   Begin VB.Label Label16 
      Caption         =   "Toolbar2 de Seçtiðin Seçenek..."
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
      Left            =   9270
      TabIndex        =   35
      Top             =   1980
      Width           =   2445
   End
   Begin VB.Label kullanici_adi 
      Caption         =   "kullanici_adi"
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
      Left            =   9270
      TabIndex        =   19
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   450
      TabIndex        =   17
      Top             =   720
      Width           =   4020
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Baðlý Kullanýcý Bulunamadý."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   12330
      TabIndex        =   8
      Top             =   675
      Width           =   2805
   End
   Begin VB.Image Image1 
      Height          =   270
      Left            =   90
      Picture         =   "FrmClient.frx":EBB8
      Top             =   675
      Width           =   270
   End
   Begin VB.Label Label7 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   630
      Width           =   15420
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3000
      Left            =   2970
      Top             =   1035
      Width           =   6255
   End
   Begin VB.Menu dosya 
      Caption         =   "&Dosya"
      Begin VB.Menu menuyeni 
         Caption         =   "Yeni"
         Begin VB.Menu menumesaj 
            Caption         =   "Mesaj"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu menuaca 
         Caption         =   "Aç"
         Begin VB.Menu menuac 
            Caption         =   "Ivedi Mesaj Dosyasý ( *.Txt )"
            Shortcut        =   ^O
         End
         Begin VB.Menu emlac 
            Caption         =   "Microsoft Outlook Express ( *.Eml )"
         End
      End
      Begin VB.Menu mnucizgi 
         Caption         =   "-"
      End
      Begin VB.Menu kaydetsave 
         Caption         =   "Kaydet"
         Shortcut        =   ^S
      End
      Begin VB.Menu cizgi52 
         Caption         =   "-"
      End
      Begin VB.Menu mnuyazdir 
         Caption         =   "Yazdýr"
         Shortcut        =   ^P
      End
      Begin VB.Menu menucizgi 
         Caption         =   "-"
      End
      Begin VB.Menu mnukimlikde 
         Caption         =   "Kimlik Deðiþtir"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnukimlikbilgileri 
         Caption         =   "Kimlik Bilgileri"
      End
      Begin VB.Menu cizgi50 
         Caption         =   "-"
      End
      Begin VB.Menu mesajoz 
         Caption         =   "Mesaj Özellikleri"
         Enabled         =   0   'False
      End
      Begin VB.Menu cizgi51 
         Caption         =   "-"
      End
      Begin VB.Menu mnuayarlar 
         Caption         =   "Ayarlar"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnukapat 
         Caption         =   "Kapat"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu gizlimenu 
      Caption         =   "gizlimenu"
      Visible         =   0   'False
      Begin VB.Menu mnugonder 
         Caption         =   "Gönder"
         Begin VB.Menu mnuarsiv 
            Caption         =   "Arþiv"
         End
         Begin VB.Menu cizgiikibin 
            Caption         =   "-"
         End
         Begin VB.Menu email 
            Caption         =   "E Mail"
         End
      End
      Begin VB.Menu cigibin 
         Caption         =   "-"
      End
      Begin VB.Menu gizliac 
         Caption         =   "Aç"
      End
      Begin VB.Menu gizlicizgi 
         Caption         =   "-"
      End
      Begin VB.Menu gizlisil 
         Caption         =   "Sil"
      End
      Begin VB.Menu gizliokunmadi 
         Caption         =   "Okunmadý Olarak Ýþaretle"
         Enabled         =   0   'False
      End
      Begin VB.Menu gizliisaretokundu 
         Caption         =   "Okundu Olarak Ýþaretle"
         Enabled         =   0   'False
      End
      Begin VB.Menu gizlicizgi2 
         Caption         =   "-"
      End
      Begin VB.Menu gizlimesajbilgi 
         Caption         =   "Mesaj Bilgi Ekraný"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu yonetim 
      Caption         =   "Yönetim"
      Begin VB.Menu mesajlar 
         Caption         =   "Mesajlar"
         Begin VB.Menu tummesajlar 
            Caption         =   "Tüm Mesajlarý Göster"
         End
         Begin VB.Menu gonderilendosyalar 
            Caption         =   "Gönderilen Tüm Dosyalarýn Listesi"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu cizgi7 
         Caption         =   "-"
      End
      Begin VB.Menu kullanicilistesi 
         Caption         =   "Kullanýcý Listesi"
      End
      Begin VB.Menu cizgi15 
         Caption         =   "-"
      End
      Begin VB.Menu yonetimiki 
         Caption         =   "Kullanýcý Yönetimi"
         Begin VB.Menu yeni_kullanici 
            Caption         =   "Yeni Kullanýcý Ekle"
         End
         Begin VB.Menu kullanici_duzenle 
            Caption         =   "Kullanýcý Düzenle"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public HiddenPreview As Boolean

'SystRay Bitiþ
Private Sub ekranda_goster()
'Seçilen Mesaj Ekranda Gösteriliyor....
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from mesajlar WHERE id ='" & Label10.Caption & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
Label3.Caption = suz![kimden] '&
Label23.Caption = suz![txtbilgi]
Label6.Caption = suz![konu]
Label4.Caption = suz![gonderim_tarihi]
RichTextBox1 = suz![mesaj]
Text1.Text = suz![id]
Text2.Text = suz![okundu]
Text3.Text = suz![silindi]
Text4.Text = suz![okundu_tarih]
Text7.Text = suz![gonderilmedi]
Label20.Caption = suz![mesajid]
Label21.Caption = suz![id]
suz.Close
Else
MsgBox "Hatalý Mesaj Biçimi", vbCritical, "Hata Kodu (C100)"
End If
End Sub


Private Sub dosya_eki_sorgula()
'On Error Resume Next
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
'Toolbarý Yenile
Form1.Toolbar2.Refresh
Dim suz2 As New Recordset
suz2.Open "Select * from mesajlar WHERE id ='" & Label10.Caption & "'", conn, adOpenKeyset, adLockOptimistic
If suz2.RecordCount <> 0 Then
'uyarý Eðer Ek Bölümnde ( 0 ) yazýyosa bu satýr hata verecektir bilgin olsun.
'Atac Bolumunde Ek Yok Ama Sýfýr Deðeride Yoksa Hata Verir Unutma.
Label15.Caption = suz2![atac]
suz2.Close
ek_olup_olmadigini_sorgula ' Toolbar Gizle Göster ver.1.0
Else
MsgBox "Hatalý Mesaj Biçimi [ Ekteki Dosyalar Açýlamýyor.]", vbCritical, "Hata Kodu (C100)"
End If
End Sub
Private Sub yeni_mesaj()
Dim resim
yeni_mesaj_gonder.Toolbar1.Buttons.Item(3).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(4).Visible = False
On Error GoTo hata
resim = programayarlari.Text4.Text & kullanici_adi & ".jpg"
yeni_mesaj_gonder.Picture3.Picture = LoadPicture(resim)
yeni_mesaj_gonder.imza.Text = Label8.Caption & vbCrLf & Label9.Caption
Exit Sub
hata:
yeni_mesaj_gonder.Picture3.Print "Ýmza Yok"
yeni_mesaj_gonder.imza.Text = Label8.Caption & vbCrLf & Label9.Caption
yeni_mesaj_gonder.Show
Exit Sub
End Sub

Private Sub ek_olup_olmadigini_sorgula()
'Eðer Dosyada Ek Yok Ýse Toolbarý Gizle ( Gösterme Varsa Göster)
If Form1.Label15.Caption = "0" Then ' Ek Yok
Form1.Toolbar2.Visible = False ' Toolbar Gizlendi.
Else
Form1.Toolbar2.Visible = True ' Toolbar Gösteriliyor.
End If
End Sub

Private Sub mesaj_goruntule()
'On Error Resume Next
If Form1.ListView1.ListItems.Count = "0" Then 'Eðer Hiç Mesaj Yok ise
Else
' Çift Týklama Yapýldýgýnda Bilgileri Aktar
yeni_mesaj_gonder.Label16.Caption = Form1.Label21.Caption
'yeni_mesaj_gonder.Label16.Caption = Form1.Text1.Text
yeni_mesaj_gonder.Frame3.Visible = True
'Bazý Menüler Gizlenecek veya Kilitlenecek
yeni_mesaj_gonder.Toolbar1.Buttons.Item(2).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(6).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(1).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(5).Visible = False
yeni_mesaj_gonder.Check1.Enabled = False
yeni_mesaj_gonder.Check2.Enabled = False
yeni_mesaj_gonder.Check3.Enabled = False
yeni_mesaj_gonder.Check4.Enabled = False
' Ek Olup Olmadýðýný Kontrol Et...
If Form1.Label15.Caption = "0" Or Form1.Label15.Caption = "" Then ' Ek Yok
yeni_mesaj_gonder.Toolbar2.Visible = False ' Toolbar Gizlendi.
Else
yeni_mesaj_gonder.Toolbar2.Visible = True ' Toolbar Gösteriliyor.
'Eðer Dosya warsa toolbarda göster ( Dosyayý )
yeni_mesaj_gonder.Label25.Caption = Form1.Label15.Caption
dosya_kontrol2 ' Ekte Gözüken Dosyalarý Toolbara Yükle

End If
yeni_mesaj_gonder.Show
End If
End Sub

Private Sub Command1_Click()
Dim soru
soru = MsgBox("Bu Mesaj Silinecek Eminmisiniz.?", vbQuestion + vbYesNo, "Dikkat")
If soru = vbYes Then
If Text1.Text = "" Then
MsgBox "Silinecek Mesajý Seçmediniz.", vbCritical, "Uyarý"
Else
Text3.Text = "1"
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from mesajlar where id = '" & Text1.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                                rs![silindi] = Text3.Text
                                rs![gonderilmedi] = "0"
                                rs.Update
                                rs.Close
End If
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Label17.Caption = ""
Label3.Caption = ""
Label6.Caption = ""
Label4.Caption = ""
RichTextBox1.Text = ""
gelen_mesajlar 'Ýþlem Birtince Gelen Mesajlarý Yenile
sol_menu 'Sol taraftaki Menüyüde...
MsgBox "Mesajýnýz Silindi.", vbInformation, "Silindi"
gelen_mesajlar 'Ýþlem Birtince Gelen Mesajlarý Yenile
okunmamis_mesajlar
Form1.Label18.Refresh
End If
End If

End Sub

Private Sub Command2_Click()
'kaydediyoruz.
'On Error Resume Next
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from mesajlar where id = '" & Text1.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                                rs![okundu_tarih] = Text4.Text
                                rs![okundu] = Text2.Text
                                rs.Update
                                rs.Close
End If
'gelen_mesajlar
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Label17.Caption = ""

End Sub

Private Sub Command3_Click()
'kaydediyoruz.
'On Error Resume Next
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from mesajlar where id = '" & Text1.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                                rs![okundu_tarih] = "okunmadý"
                                rs![okundu] = "0"
                                rs.Update
                                rs.Close
End If
gelen_mesajlar
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Label17.Caption = ""
'Temizleme iþlemini bitir.
' Kayýtlarý Guncelle.....
' Sürekli Kontrol Edilecek Olan Menüleri En Sona Yaz _
Gelen Mesajlar Fonsksiyonunu iki Defa Caðýr Yenilemede Sorun Cýkartýyo.
sol_menu
gelen_mesajlar '1
okunmamis_mesajlar
gelen_mesajlar '2
End Sub


Private Sub Command4_Click()
On Error Resume Next
Dim sFile As String
With cmd1
.DialogTitle = "Farklý Kaydet"
.CancelError = False
.Filter = "Text Dosya Türü (*.txt)|*.txt" ' Süper Tekstil
.ShowSave
If Len(.FileName) = 0 Then
Exit Sub
End If
sFile = .FileName
Open sFile For Output As #1
Print #1, "Süper Tekstil San.Tic.Aþ."
Print #1, "Gönderen: " & Label3.Caption & " " & Label23.Caption
Print #1, "Gönderim Tarihi :" & Label4.Caption
Print #1, "Konu :" & Label6.Caption
Print #1, "****************************Mesajý***************************"
'Print #1, RichTextBox1.Text
Print #1, Text6.Text
Close #1
MsgBox " Kayýtlar baþarý ile aktarýlmýþtýr..", vbInformation, "Tamamlandý."
End With
End Sub

Private Sub Command5_Click()
Dim baslik As Integer
Dim bul As String
bul = Text5.Text 'InputBox("Aranak Kiþi: " & Adi, "Arama")
'baslik = lvwTex 'Ýl Kayýttaki Bilgileri Arama için Kullanabilirsin.
baslik = lvwSubItem 'Alt menülerde ara
Dim altmenu As ListItem
Set altmenu = ListView1.FindItem(bul, baslik, , lvwPartial)
If altmenu Is Nothing Then
MsgBox bul & " Böyle Bir Kayýt Yok.'" & vbCrLf, vbInformation + vbOKOnly, "Arama"
Exit Sub
Else
altmenu.EnsureVisible
altmenu.Selected = True
ListView1.SetFocus
End If
End Sub

Private Sub Command6_Click()
'kaydediyoruz.
'On Error Resume Next
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from mesajlar where id = '" & Text1.Text & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                                rs![okundu_tarih] = Text4.Text
                                rs![gonderilmedi] = "1"
                                rs.Update
                                rs.Close
End If
'gelen_mesajlar
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Label17.Caption = ""
'Kayýtlarý Güncelle
sol_menu
gelen_mesajlar '1
okunmamis_mesajlar
gelen_mesajlar '2
MsgBox "Mesaj Arþiv Klasörünüze Kayýt Edildi.", vbInformation, "Arþiv"
gelen_mesajlar '3
sol_menu
End Sub

Private Sub dikey_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Res As Long
dikey.BackColor = vbBlack
DoEvents
ReleaseCapture
On Error Resume Next
Res = SendMessage(dikey.hWnd, WM_SYSCOMMAND, 61458, 0)
dikey.BackColor = vbButtonFace
If dikey.Left > 9180 Then dikey.Left = 9000
If dikey.Left < 2675 Then dikey.Left = 2700
If yatay.Top < 1500 Then yatay.Top = 2200
ListView1.Left = dikey.Left + 60
Shape1.Left = dikey.Left + 60
Frame1.Left = dikey.Left + 60 ' 60
Label6.Width = Form1.Width - dikey.Left - 5300 ' 5300
Label3.Width = Form1.Width - dikey.Left - 1100
Label4.Width = Form1.Width - dikey.Left - 1200 ' 1300
Label1.Width = Form1.Width - dikey.Left - 450
Label2.Width = Form1.Width - dikey.Left - 450
Label5.Width = Form1.Width - dikey.Left - 4300
Treeview1.Width = dikey.Left - 50
'List1.Width = dikey.Left - 50
ListView2.Width = dikey.Left - 50
yatay.Left = dikey.Left + 60
dikey.Top = Treeview1.Top - 50
RichTextBox1.Left = dikey.Left + 60
ListView1.Width = Form1.Width - dikey.Left - 260
Shape1.Width = Form1.Width - dikey.Left - 260
RichTextBox1.Width = Form1.Width - dikey.Left - 285
Frame1.Width = Form1.Width - dikey.Left - 275 ' 275
'ListView1.ColumnHeaders.Item(3).Width = Form1.Width - dikey.Left - 4340
End Sub

Private Sub email_Click()
MsgBox "Sayýn Kullanýcý Bu Modül Henüz Eklenmedi.", vbInformation, "Uyarý"

End Sub

Private Sub emlac_Click()
MsgBox "invalid.program.error."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'******************************************************
'F2 Tuþuna Basýlýrsa Yeni Mesaj
If KeyCode = vbKeyF2 Then
yeni_mesaj_gonder.Show
End If
'******************************************************
'F3 Tuþuna Basýlýrsa Mesajlarý ara
If KeyCode = vbKeyF3 Then
Frame4.Visible = True
End If
'******************************************************
'F5 Tuþuna Basýlýrsa Mesajlarý Tazele
If KeyCode = vbKeyF5 Then
gonder_al.Show
End If
'******************************************************
'Delete Tuþuna Basýlýrsa Mesajlarý Sil
If KeyCode = vbKeyDelete Then
If Label3.Caption = "" Then
Else
Command1_Click ' Mesajý Sil
End If
End If
'******************************************************
'Enter Tuþuna Basýlýrsa Mesajlarý Görüntüle
If KeyCode = vbKeyReturn Then
If Label3.Caption = "" Then
Else
yeni_mesaj_gonder.Caption = Form1.Label21.Caption & " Nolu Mesaj Okunuyor."
Call mesaj_goruntule
End If
End If
'******************************************************
End Sub

Private Sub Form_Load()

okunmamis_mesajlar           ' Okunmamýþ Mesajlarýn Sayýsýný Göster
sol_menu                     ' Treeview içindeki menuler yuklenýyor.
liste_kullanicilar           ' Kullanýcýlarý Listede Göster
gelen_mesajlar               ' Gelen Mesajlarý Göster
'
'
'
StatusBar1.Panels(3) = Date & "  " & Time


End Sub

Private Sub Form_Resize()
On Error Resume Next
If Form1.ScaleHeight = 11835 Or Form1.Width = 11505 Then
'MsgBox "Hata Yaptýnýz."
Else
ListView1.Width = Form1.Width - 3150
Shape1.Width = Form1.Width - 3150
Frame1.Width = Form1.Width - 3100 ' 3100
DoEvents
If Not HiddenPreview = True Then RichTextBox1.Height = Form1.Height - 600
If HiddenPreview = True Then ListView1.Height = Form1.Height - 3340 ' 2340 ' 3340
If HiddenPreview = True Then Shape1.Height = Form1.Height - 2340
DoEvents
Label6.Width = Form1.Width - 8000 ' 8100
DoEvents
DoEvents
'List1.Height = Form1.Height - 5600 ' 4430
ListView2.Height = Form1.Height - 5800 '5600 ' 4430
DoEvents
yatay.Width = Frame1.Width
'yatay.Width = Toolbar2.Width
DoEvents
yatay.Left = Frame1.Left
'yatay.Left = Toolbar2.Left
DoEvents
dikey.Height = Form1.Height - 1900
DoEvents
If Not yatay.Top <= 2000 Then yatay.Top = Form1.Height - 4500
If Me.Height < 6800 Then yatay.Top = 2100 ' 2100
If Not HiddenPreview = True Then ListView1.Height = yatay.Top - 1020 ' 1430
If Not HiddenPreview = True Then Shape1.Height = yatay.Top - 1430 ' 1430
Frame1.Top = yatay.Top - 20 ' 20
Toolbar2.Top = yatay.ToolTipText - 20
DoEvents
RichTextBox1.Height = Form1.Height - ListView1.Height - 2800 ' aþaðýya
RichTextBox1.Top = yatay.Top + 720
dikey.Top = Treeview1.Top - 50
Label3.Width = Form1.Width - dikey.Left - 1100
Label4.Width = Form1.Width - dikey.Left - 1300
Label1.Width = Form1.Width - dikey.Left - 450
Label2.Width = Form1.Width - dikey.Left - 450
Label5.Width = Form1.Width - dikey.Left - 4300
Label6.Width = Form1.Width - dikey.Left - 5300 ' 5300
ListView1.Left = dikey.Left + 60 ' 60
Shape1.Left = dikey.Left + 60
yatay.Left = dikey.Left + 60
Frame1.Left = dikey.Left + 60
RichTextBox1.Left = dikey.Left + 60
yatay.Left = Frame1.Left
Treeview1.Width = dikey.Left - 50
'List1.Width = dikey.Left - 50 '<--saga
ListView1.Width = dikey.Left - 50 '<--saga
DoEvents
Label8.Left = Me.Width - 2400
DoEvents
RichTextBox1.Left = dikey.Left + 60
DoEvents
ListView1.Width = Form1.Width - dikey.Left - 260
DoEvents
Shape1.Width = Form1.Width - dikey.Left - 260
DoEvents
RichTextBox1.Width = Form1.Width - dikey.Left - 275
DoEvents
Frame1.Width = Form1.Width - dikey.Left - 260
Toolbar1.Width = Form1.Width - dikey.Left - 260
DoEvents
If Label8.Left < 6550 Then Label8.Visible = False Else Label8.Visible = True
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cancel = True
'Programý Kapat
End Sub

Private Sub gizliac_Click()
' Dosya Açmak Ýçin
' Çift Týklama Yapýldýgýnda Bilgileri Aktar
yeni_mesaj_gonder.Label16.Caption = Form1.ListView1.SelectedItem.SubItems(11)
yeni_mesaj_gonder.Frame3.Visible = True
'Bazý Menüler Gizlenecek veya Kilitlenecek
yeni_mesaj_gonder.Toolbar1.Buttons.Item(2).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(6).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(1).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(5).Visible = False
yeni_mesaj_gonder.Check1.Enabled = False
yeni_mesaj_gonder.Check2.Enabled = False
yeni_mesaj_gonder.Check3.Enabled = False
yeni_mesaj_gonder.Check4.Enabled = False
' Ek Olup Olmadýðýný Kontrol Et...
If Form1.Label15.Caption = "0" Or Form1.Label15.Caption = "" Then ' Ek Yok
yeni_mesaj_gonder.Toolbar2.Visible = False ' Toolbar Gizlendi.
Else
yeni_mesaj_gonder.Toolbar2.Visible = True ' Toolbar Gösteriliyor.
'Eðer Dosya warsa toolbarda göster ( Dosyayý )
yeni_mesaj_gonder.Label25.Caption = Form1.Label15.Caption
dosya_kontrol2 ' Ekte Gözüken Dosyalarý Toolbara Yükle
End If
yeni_mesaj_gonder.Show
End Sub

Private Sub gizliisaretokundu_Click()
Command2_Click
End Sub

Private Sub gizlimesajbilgi_Click()
MsgBox "Bu Mesajý Okuduðunuz Tespit Edildi.", , "Mesaj Bilgi Ekraný"

End Sub

Private Sub gizliokunmadi_Click()
Command3_Click
' Sürekli Kontrol Edilecek Olan Menüleri En Sona Yaz _
Gelen Mesajlar Fonsksiyonunu iki Defa Caðýr Yenilemede Sorun Cýkartýyo.
sol_menu
gelen_mesajlar '1
okunmamis_mesajlar
gelen_mesajlar '2
End Sub

Private Sub gizlisil_Click()
If Label3.Caption = "" Then
'MsgBox "Mesaj Yok."
Else
'Mesajý Silmeden önce Mutlaka seçmelisin.?
Command1_Click
End If
End Sub




Private Sub Image2_Click()
Frame4.Visible = False
End Sub

Private Sub kaydetsave_Click()
If Label3.Caption = "" Then ' eðer mesaj seçilmedi ise
MsgBox "Kayýt Edilecek Mesajý Seçmediniz.!", vbCritical, "Hata"
Else
Command4_Click
'Dim sFile As String
'With cmd1
'.DialogTitle = "Farklý Kaydet"
'.CancelError = False
'.Filter = "SPR Dosya Türü (*.spr)|*.spr" ' Süper Tekstil
'.ShowSave
'If Len(.FileName) = 0 Then
'Exit Sub
'End If
'sFile = .FileName
'End With
'Form1.RichTextBox1.SaveFile sFile
'MsgBox "Mesaj Kayýt Edildi.", vbInformation, "Dosya Kaydet"
End If
End Sub

Private Sub kullanici_duzenle_Click()
'Kullanýcý Düzenle
kullanici_ekle.Command3.Visible = True
kullanici_ekle.Command4.Visible = True
kullanici_ekle.Command5.Visible = True
kullanici_ekle.Command1.Enabled = False
kullanici_ekle.Check2.Value = 0
kullanici_ekle.Label14.Caption = "Dikkat : Kullanýcý Adýný Deðiþtiremezsiniz.."
kullanici_ekle.Label12.Caption = "" 'Ýd Numarasýný Gizle
kullanici_ekle.Text2.Enabled = False 'Kullanýcý Adý Deðiþtirilemesin.
kullanici_ekle.Show
End Sub

Private Sub kullanicilistesi_Click()
kullanici_listesi.Show
End Sub

Private Sub Label16_Change()
If Label16.Caption = "Tümünü kaydet" Then
ek_kaydet.Label3.Caption = Form1.Label15.Caption
ek_kaydet.Show
Else
uygulama_calistir.Show
End If
'
End Sub

Private Sub Label17_Change()
'noyu_sorgula
'Eðer Mesaja Týklanýrsa Mesaj Nosunu Text1 De Göster Sonra _
Okundu Deðerini 1 yap ve Okundu Tarihini Ekle.Kaydet
If Label17.Caption = "0" Then 'Mesaj Okunmamýþ ise iþlemi Baþlat
Text4.Text = Date & "-" & Time
Text2.Text = "1"
Timer2.Enabled = True 'Okuma Durumunu Görüntüle Okundu Bilgisi Gözüksün.
Command2_Click
Else
End If
End Sub

Private Sub Label18_Change()
'Bu Alana Deðiþiminde Sürekli Aþaðýdaki Temizleme Olaylarý Meydana Gelsin
Form1.Label6.Caption = ""
Form1.Label3.Caption = ""
Form1.Label4.Caption = ""
Form1.RichTextBox1.Text = ""
Toolbar2.Visible = False
'Dim i
'For i = 0 To 100

'Seçenek 1
If Label18.Caption = "Gelen Mesajlar" Then
Form1.ListView1.ColumnHeaders.Item(6).Text = "Kimden"
Label19.Caption = "(" & Form1.okunmamis.ListItems.Count & ")" 'Kaç Tane Yeni Mesaj Olduðunu Sorgula
gelen_mesajlar               ' Gelen Mesajlarý Göster
'sol_menu                     ' Treeview içindeki menuler yuklenýyor.
Form1.mnugonder.Enabled = True
End If


'Seçenek 2
If Label18.Caption = "Giden Mesajlar" Then
Form1.ListView1.ColumnHeaders.Item(6).Text = "Kime"

giden_mesajlar
Form1.mnugonder.Enabled = False
End If


'Seçenek 3
If Label18.Caption = "Gitmeyen Mesajlar" Then
Form1.ListView1.ColumnHeaders.Item(6).Text = "Kime"
gitmeyen_mesajlar
End If


'Seçenek 4
If Label18.Caption = "Silinmiþ Mesajlar" Then
Form1.ListView1.ColumnHeaders.Item(6).Text = "Kimden"
silinmis_mesajlar
End If

'Seçenek 5
If Label18.Caption = "Arþiv" Then
Form1.ListView1.ColumnHeaders.Item(6).Text = "Kimden"
gitmeyen_mesajlar
End If
'Next i
End Sub

Private Sub Label19_Change()
'Eðer Hiç Mesaj Yok ise Ekranda Mavi Yazý Cýkartma
If Label19.Caption = (0) Then
Form1.Label19.Visible = False
Else
Form1.Label19.Visible = True
mesaj_geldii.Show
End If
End Sub

Private Sub ListView1_Click()
'On Error Resume Next
' Form1.Label17.Caption = Form1.ListView1.SelectedItem.SubItems(9)
'Listviewdeki secilen mesajýn indexi alýnacak
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
 Dim Orden As Integer
    ListView1.SortKey = ColumnHeader.Index - 1
    Orden = ListView1.SortKey
    ListView1.SortOrder = Abs(Not ListView1.SortOrder = 1)
    ListView1.Sorted = True
    
End Sub


Private Sub ListView1_DblClick()
On Error Resume Next
If Form1.ListView1.ListItems.Count = "0" Then 'Eðer Hiç Mesaj Yok ise
Else
If Form1.Label3.Caption = "" Then 'Eðer Seçili Mesaj Yok Ýse Çift Týklandýðýnda Boþ Formu Açma
Else
' Çift Týklama Yapýldýgýnda Bilgileri Aktar
'yeni_mesaj_gonder.Label16.Caption = Form1.ListView1.SelectedItem.SubItems(11)
yeni_mesaj_gonder.Label16.Caption = Form1.Label21.Caption
yeni_mesaj_gonder.Frame3.Visible = True
'Bazý Menüler Gizlenecek veya Kilitlenecek
yeni_mesaj_gonder.Toolbar1.Buttons.Item(2).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(6).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(1).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(5).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(12).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(13).Visible = False
yeni_mesaj_gonder.mesaji_gonder.Enabled = False
yeni_mesaj_gonder.daha_sonra_gonder.Enabled = False
yeni_mesaj_gonder.Check1.Enabled = False
yeni_mesaj_gonder.Check2.Enabled = False
yeni_mesaj_gonder.Check3.Enabled = False
yeni_mesaj_gonder.Check4.Enabled = False
' Ek Olup Olmadýðýný Kontrol Et...
If Form1.Label15.Caption = "0" Or Form1.Label15.Caption = "" Then ' Ek Yok
yeni_mesaj_gonder.Toolbar2.Visible = False ' Toolbar Gizlendi.
Else
yeni_mesaj_gonder.Toolbar2.Visible = True ' Toolbar Gösteriliyor.
'Eðer Dosya warsa toolbarda göster ( Dosyayý )
yeni_mesaj_gonder.Label25.Caption = Form1.Label15.Caption
dosya_kontrol2 ' Ekte Gözüken Dosyalarý Toolbara Yükle
End If

yeni_mesaj_gonder.Caption = Form1.Label21.Caption & " Nolu Mesaj Okunuyor."
yeni_mesaj_gonder.Show
End If
End If
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Form1.Label18.Caption = "Giden Mesajlar" Then
Form1.Label10.Caption = ListView1.SelectedItem.SubItems(11)
If Label10.Caption = "" Then
'Ekraný Temizlele,
Label3.Caption = ""
Label4.Caption = ""
Label6.Caption = ""
RichTextBox1.Text = ""
Else
ekranda_goster
dosya_eki_sorgula 'Once Kaydý kontrol Et.
dosya_kontrol 'Toolbar da Gizleme Gösterme Ýþlemi yapýlacak.
End If
'End If
Else
Form1.Label17.Caption = Form1.ListView1.SelectedItem.SubItems(9)
Form1.Label10.Caption = ListView1.SelectedItem.SubItems(11)
If Label10.Caption = "" Then
'Ekraný Temizlele,
Label3.Caption = ""
Label4.Caption = ""
Label6.Caption = ""
RichTextBox1.Text = ""
Else
ekranda_goster
'Dosya Tarnsferi ve Kontrol Komutlarý
dosya_eki_sorgula 'Once Kaydý kontrol Et.
dosya_kontrol 'Toolbar da Gizleme Gösterme Ýþlemi yapýlacak.
End If
End If
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Sað Tuþa Basýldýgýnda
If Form1.Label18.Caption = "Arþiv" Then
Form1.mnugonder.Enabled = False
If Button = 2 Then
PopupMenu gizlimenu
End If
'Eðer Arþiv Seçili Ýse Sað Tuþu Gizle
Else
If Button = 2 Then
PopupMenu gizlimenu
End If
End If
End Sub

Private Sub ListView2_DblClick()
'Eðer Sol Menüde Kullanýcýya Çift Týklandýysa.
yeni_mesaj_gonder.Text1.Text = ListView2.SelectedItem.SubItems(1) & ";"
yeni_mesaj_gonder.Show

End Sub

Private Sub menuac_Click()
spr_mesaj_oku.Show
End Sub




Private Sub menumesaj_Click()
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(3).Visible = False
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(4).Visible = False
'yeni_mesaj_gonder.Show
Call yeni_mesaj
End Sub

Private Sub mnuarsiv_Click()
If Form1.Label3.Caption = "" Then 'Eðer Mesaj Seçili Deðilse Arþive Gönderme
MsgBox "Arþive Gidecek Mesajý Seçmediniz.!", vbExclamation, "Uyarý"
Else
Command6_Click
End If
End Sub

Private Sub mnuayarlar_Click()
programayarlari.Show
End Sub

Private Sub mnukapat_Click()
End
End Sub

Private Sub mnukimlikbilgileri_Click()
kullanici_bilgileri.Show
End Sub

Private Sub mnukimlikde_Click()
'Kimlik Deðiþtir.
Baglan.Image3.Visible = False
Baglan.Image4.Visible = False
Baglan.Label8.Caption = "Bu Uygulamayý Kapatmak için Üzerine Çift Týklayýn..."
Baglan.Show
End Sub

Private Sub mnuyazdir_Click()
On Error Resume Next

    With cmd1
        .DialogTitle = "Yazdýr"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If spr_mesaj_oku.RichTextBox1.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter

        
          Form1.RichTextBox1.SelPrint .hDC
        'End If
    End With
End Sub

Private Sub RichTextBox1_Change()
Text6.Text = RichTextBox1.Text
End Sub

Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Dim baslik As Integer
Dim bul As String
bul = Text5.Text 'InputBox("Aranak Kiþi: " & Adi, "Arama")
'baslik = lvwTex 'Ýl Kayýttaki Bilgileri Arama için Kullanabilirsin.
baslik = lvwSubItem 'Alt menülerde ara
Dim altmenu As ListItem
Set altmenu = ListView1.FindItem(bul, baslik, , lvwPartial)
If altmenu Is Nothing Then
MsgBox "Böyle Bir Kayýt Yok.'" & vbCrLf, vbInformation + vbOKOnly, "Arama"
Exit Sub
Else
altmenu.EnsureVisible
altmenu.Selected = False
ListView1.SetFocus
End If
End If

End Sub


Private Sub Timer2_Timer()
' 4 sn sonra mesajlarýn okundugu bildirilsin.
'Temizleme iþlemini bitir.
' Kayýtlarý Guncelle.....
okunmamis_mesajlar
sol_menu
gelen_mesajlar
okunmamis_mesajlar
'Ýþlemi Ýkinciye Tekrarla
sol_menu
gelen_mesajlar
okunmamis_mesajlar
Timer2.Enabled = False ' Sayacý Durdur.
End Sub

Private Sub Timer3_Timer()
' 5 SN.DE BÝR MESAJLARIN KONTROLÜ YAPILACAK.
'gelen_mesajlar
'sol_menu
'Form1.Label18.Caption = "Gelen Mesajlar"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
'**********************************************************
Case "a"
Call yeni_mesaj
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(3).Visible = False
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(4).Visible = False
'yeni_mesaj_gonder.Show
'**********************************************************
Case "b": 'Yanýtla
If Label3.Caption = "" Then
Else
mesaji_ilet.Caption = Form1.Label20.Caption & " Nolu Mesaj Cevaplanýyor..." ' Mesaj Numarasýný Ver
Dim cizgi
cizgi = "________________________________________________________"
mesaji_ilet.Text1.Text = Label3.Caption & ";"
mesaji_ilet.Text3.Text = "Ynt:> " & Label6.Caption
mesaji_ilet.RichTextBox1.Text = vbCrLf & vbCrLf & cizgi & vbCrLf & "Mesajý Gönderen: " & Label3.Caption & vbCrLf & "Mesaj Gönderim Tarihi: " & Label4.Caption & vbCrLf & cizgi & vbCrLf & RichTextBox1.Text
mesaji_ilet.Show
End If
'**********************************************************
Case "c" ' ilet
On Error GoTo hata
If Label3.Caption = "" Then
Else
mesaji_ilet.Caption = Form1.Label20.Caption & " Nolu Mesaj Cevaplanýyor..."  ' Mesaj Numarasýný Ver
'mesaji_ilet.Text1.Text = Label3.Caption & ";"
mesaji_ilet.Text3.Text = Label6.Caption
mesaji_ilet.Show
End If
Exit Sub
hata:
MsgBox "Ýleti Yapacaðýnýz Mesajýný Seçmediniz."
Exit Sub
'**********************************************************
Case "d" ' Sil
If Label3.Caption = "" Then
'MsgBox "Mesaj Yok."
Else
'Mesajý Silmeden önce Mutlaka seçmelisin.?
Command1_Click
End If
'**********************************************************
Case "e" 'Gönder Al
Refresh
gonder_al.Show
'**********************************************************
Case "f" 'Bul
Frame4.Visible = True
Text5.SetFocus
'**********************************************************
Case "g"
'
'---------------------------------------------------
Case "h"
'
'---------------------------------------------------
Case "i" 'Kapat
End
'---------------------------------------------------
End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
Case "a"
kullanici_bilgileri.Show
Case "b"
Baglan.Image3.Visible = False
Baglan.Image4.Visible = False
Baglan.Label8.Caption = "Bu Uygulamayý Kapatmak için Üzerine Çift Týklayýn..."
Baglan.Show
Case "c"
kullanici_ekle.Show
Case "d"
'Kullanýcý Düzenle
kullanici_ekle.Command3.Visible = True
kullanici_ekle.Command4.Visible = True
kullanici_ekle.Command5.Visible = True
kullanici_ekle.Command1.Enabled = False
kullanici_ekle.Check2.Value = 0
kullanici_ekle.Label14.Caption = "Dikkat : Kullanýcý Adýný Deðiþtiremezsiniz.."
kullanici_ekle.Label12.Caption = "" 'Ýd Numarasýný Gizle
kullanici_ekle.Text2.Enabled = False 'Kullanýcý Adý Deðiþtirilemesin.
kullanici_ekle.Show
' YARDIM VE PROGAM HAKKINDA BILGILERI BU BOLUMDE GOSTERTECEZ LOE
Case "e"
'Yardým Konularý
'MsgBox "Yardým Konularýný Görüntülemek için Yardým.hlp dosyasýný Satýn almanýz Gerekir.", vbInformation, "Bilgi Mesajý"
MsgBox "Ivedi Mesaj 1.0 Yardým Konusu Ýçermiyor.", vbInformation, "Bilgi Mesajý"

Case "f"
'Program Hakkýnda
program_hakkinda.Show
Case "g"
'Versiyon
MsgBox program_hakkinda.Label2.Caption, vbInformation, "Bilgi"
Case "h"
MsgBox "Update Bilgilerine Ulaþýlamadý.", vbCritical, "Bilinmeyen Uygulama Hatasý"
End Select
End Sub


Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Label16.Caption = ButtonMenu.Text
uygulama_calistir.Label4.Caption = ButtonMenu.Text 'Uygulama Ekranýna Dosya Adýný Göster
End Sub

Private Sub Treeview1_Click()
If Form1.ListView1.ListItems.Count = "0" Then 'Eðer Hiç Mesaj Yok ise
Form1.Text9.Visible = True
Else
Form1.Text9.Visible = False
End If

End Sub

Private Sub Treeview1_DblClick()
If Text8.Visible = False Or Picture3.Visible = False Then
'Eðer Ýkiside Gizli Ýse
Text8.Visible = True
Picture3.Visible = True
Else
Text8.Visible = False
Picture3.Visible = False
End If
End Sub


Private Sub Treeview1_NodeClick(ByVal Node As MSComctlLib.Node)
Label18.Caption = Treeview1.Nodes.Item(Treeview1.SelectedItem.Index).Text
End Sub


Private Sub tummesajlar_Click()
tum_mesajlar_listesi.Show
End Sub

Private Sub yatay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Res As Long
yatay.BackColor = vbBlack
ReleaseCapture
On Error Resume Next
Res = SendMessage(yatay.hWnd, WM_SYSCOMMAND, 61458, 0)
yatay.BackColor = vbButtonFace
If yatay.Top < 1500 Then yatay.Top = 3400 ' 2400
If yatay.Top > Form1.Height - 1000 Then yatay.Top = Form1.Height - 3000
ListView1.Height = yatay.Top - 1030 ' 1410
Shape1.Height = yatay.Top - 1030
Frame1.Top = yatay.Top - 20
yatay.Width = Frame1.Width
yatay.Left = Frame1.Left
RichTextBox1.Height = Form1.Height - ListView1.Height - 2790 ' 3080
RichTextBox1.Top = yatay.Top + 720 ' 720
End Sub

Private Sub yeni_kullanici_Click()
kullanici_ekle.Show
End Sub



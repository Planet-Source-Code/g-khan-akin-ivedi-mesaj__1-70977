VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form yeni_mesaj_gonder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yeni Mesaj"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13095
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   10350
      ScaleHeight     =   1095
      ScaleWidth      =   2355
      TabIndex        =   51
      Top             =   5670
      Width           =   2355
   End
   Begin MSComDlg.CommonDialog cmd2 
      Left            =   765
      Top             =   6030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1680
      Left            =   45
      TabIndex        =   25
      Top             =   675
      Visible         =   0   'False
      Width           =   13020
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   10665
         TabIndex        =   45
         Top             =   1350
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox Text5 
         Height          =   735
         Left            =   11475
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   44
         Text            =   "Form2.frx":06EA
         Top             =   180
         Visible         =   0   'False
         Width           =   1365
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10845
         Top             =   765
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
               Picture         =   "Form2.frx":06F2
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   570
         Left            =   12105
         TabIndex        =   38
         Top             =   1035
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1005
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageIndex      =   1
               Style           =   5
               Object.Width           =   1e-4
            EndProperty
         EndProperty
      End
      Begin VB.Label Label27 
         BackColor       =   &H80000010&
         BackStyle       =   0  'Transparent
         Caption         =   "...."
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
         Left            =   7020
         TabIndex        =   49
         Top             =   765
         Width           =   3795
      End
      Begin VB.Label Label26 
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
         Left            =   6525
         TabIndex        =   48
         Top             =   765
         Width           =   465
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FF0000&
         Caption         =   "Toolbar Hareketleri"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   10710
         TabIndex        =   41
         Top             =   495
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label25 
         BackColor       =   &H000000FF&
         Caption         =   "Ekte Gelen Dosyalar"
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
         Height          =   285
         Left            =   10665
         TabIndex        =   40
         Top             =   180
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label24 
         BackColor       =   &H8000000C&
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
         Left            =   3060
         TabIndex        =   39
         Top             =   495
         Width           =   2580
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000010&
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
         Left            =   6210
         TabIndex        =   37
         Top             =   1305
         Width           =   3435
      End
      Begin VB.Label Label21 
         Caption         =   "Okunma Tarihi:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4860
         TabIndex        =   36
         Top             =   1305
         Width           =   1905
      End
      Begin VB.Label Label20 
         BackColor       =   &H8000000C&
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
         Left            =   2250
         TabIndex        =   35
         Top             =   1305
         Width           =   2490
      End
      Begin VB.Label Label19 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
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
         Left            =   2250
         TabIndex        =   34
         Top             =   1035
         Width           =   8250
      End
      Begin VB.Label Label18 
         BackColor       =   &H8000000C&
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
         Left            =   2250
         TabIndex        =   33
         Top             =   765
         Width           =   4155
      End
      Begin VB.Label Label17 
         BackColor       =   &H8000000C&
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
         Left            =   2250
         TabIndex        =   32
         Top             =   495
         Width           =   1365
      End
      Begin VB.Label Label16 
         BackColor       =   &H8000000C&
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
         Left            =   2250
         TabIndex        =   31
         Top             =   225
         Width           =   2850
      End
      Begin VB.Label Label15 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Gönderim Tarihi:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   30
         Top             =   1305
         Width           =   1590
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Konu:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   29
         Top             =   1035
         Width           =   870
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Kime:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   28
         Top             =   765
         Width           =   870
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Kimden:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   27
         Top             =   495
         Width           =   870
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Mesaj No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   26
         Top             =   225
         Width           =   870
      End
   End
   Begin RichTextLib.RichTextBox imza 
      Height          =   645
      Left            =   10125
      TabIndex        =   15
      Top             =   6795
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   1138
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"Form2.frx":0A0C
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
   Begin VB.TextBox Text4 
      BackColor       =   &H00C0FFFF&
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
      Left            =   495
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Eklenen Dosyalar [ Sonradan Göndermek Ýstemediðiniz Dosyalarý Metin Alanýndan Silebilirsiniz.]"
      Top             =   1980
      Visible         =   0   'False
      Width           =   12570
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   225
      Top             =   6030
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Mesaj Göndereceðiniz Kiþiye Dosya Ekle"
      FileName        =   "*.*"
      Filter          =   "Tüm Dosyalar ( *.* )"
      FontSize        =   9
   End
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   45
      TabIndex        =   8
      Top             =   2340
      Width           =   13020
      Begin VB.CommandButton Command3 
         Caption         =   "print"
         Height          =   240
         Left            =   4860
         TabIndex        =   43
         Top             =   225
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton sil 
         Caption         =   "sil"
         Height          =   240
         Left            =   5715
         TabIndex        =   42
         Top             =   225
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton Command1 
         Caption         =   "bilgi"
         Height          =   240
         Left            =   2340
         TabIndex        =   24
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "mesaj"
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
         Left            =   1530
         TabIndex        =   18
         Top             =   225
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Yanýtlayýn"
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
         Left            =   11430
         TabIndex        =   12
         Top             =   225
         Width           =   1185
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Bilgilendirme"
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
         Left            =   7065
         TabIndex        =   11
         Top             =   225
         Width           =   1500
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Ýnceleyin"
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
         Left            =   3780
         TabIndex        =   10
         Top             =   180
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Acil"
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
         Left            =   630
         TabIndex        =   9
         Top             =   180
         Width           =   1140
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   11115
         Picture         =   "Form2.frx":0A8B
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   6705
         Picture         =   "Form2.frx":28FD
         Top             =   225
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   270
         Picture         =   "Form2.frx":476F
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   3510
         Picture         =   "Form2.frx":65E1
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label9 
         Caption         =   "Mesaj No:"
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
         Left            =   10395
         TabIndex        =   22
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "KN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9360
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Kayýt No:"
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
         Left            =   8505
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         Caption         =   "MN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11340
         TabIndex        =   23
         Top             =   180
         Visible         =   0   'False
         Width           =   1005
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
      Left            =   495
      TabIndex        =   7
      Top             =   1575
      Width           =   12570
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
      Left            =   495
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1170
      Width           =   12570
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
      Left            =   495
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   765
      Width           =   12570
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -450
      TabIndex        =   1
      Top             =   585
      Width           =   14055
   End
   Begin MSComctlLib.ImageList renkli 
      Left            =   630
      Top             =   6525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   7575
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2117
            MinWidth        =   2117
            Text            =   "Kullanýcý:"
            TextSave        =   "Kullanýcý:"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Þu anda Mesaj Gönderecek Kiþi"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4480
            MinWidth        =   4480
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Tarih:"
            TextSave        =   "Tarih:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Saat:"
            TextSave        =   "Saat:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   1764
            MinWidth        =   1764
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2187
            MinWidth        =   2187
            Text            =   "Mesaj ID:"
            TextSave        =   "Mesaj ID:"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   4480
            MinWidth        =   4480
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList renksiz 
      Left            =   45
      Top             =   6525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":8453
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":8B4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":90E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":9539
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":98D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":9FCD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A6C7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4290
      Left            =   225
      TabIndex        =   19
      Top             =   3105
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   7567
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"Form2.frx":AC61
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1005
      ButtonWidth     =   2434
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "renksiz"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Gönder      "
            Key             =   "a"
            Object.ToolTipText     =   "Mesajý Gönder"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Yanýtla      "
            Key             =   "b"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Ekle      "
            Key             =   "c"
            Object.ToolTipText     =   "Mesajýnýza Dosya Ekleyin"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Yazdýr      "
            Key             =   "d"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Sil      "
            Key             =   "e"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "      Kullanýcýlar      "
            Key             =   "f"
            Object.ToolTipText     =   "Kapat"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "     Kapat     "
            Key             =   "g"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   12555
         Picture         =   "Form2.frx":ACE0
         ScaleHeight     =   420
         ScaleWidth      =   420
         TabIndex        =   47
         Top             =   90
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H8000000E&
      Height          =   4560
      Left            =   45
      ScaleHeight     =   4500
      ScaleWidth      =   12915
      TabIndex        =   50
      Top             =   2925
      Width           =   12975
   End
   Begin VB.Label Label6 
      Caption         =   "mesajid"
      Height          =   240
      Left            =   4635
      TabIndex        =   17
      Top             =   6525
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "id"
      Height          =   285
      Left            =   4635
      TabIndex        =   16
      Top             =   6210
      Width           =   1230
   End
   Begin VB.Label Label4 
      Caption         =   "Ek"
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
      TabIndex        =   14
      Top             =   2025
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label Label3 
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
      Height          =   240
      Left            =   45
      TabIndex        =   4
      ToolTipText     =   "Konu Ýçeriðini Buraya Yazýn.."
      Top             =   1620
      Width           =   420
   End
   Begin VB.Label Label2 
      Caption         =   "Bilgi:"
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
      TabIndex        =   3
      ToolTipText     =   "Bilgi Vermek Ýstediðiniz Kiþi ve Kiþileri Buradan Seçin.."
      Top             =   1215
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Kime:"
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
      Height          =   240
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Mesaj Göndermek Ýstediðiniz Kiþi veya Kiþileri Buradan Seçin."
      Top             =   810
      Width           =   420
   End
   Begin VB.Menu dosya 
      Caption         =   "Dosya"
      Begin VB.Menu yeni 
         Caption         =   "Yeni"
         Begin VB.Menu yeni_mesaj 
            Caption         =   "Mesaj"
            Shortcut        =   ^N
         End
      End
      Begin VB.Menu cizgi21 
         Caption         =   "-"
      End
      Begin VB.Menu mesaji_gonder 
         Caption         =   "Mesajý Gönder"
         Shortcut        =   ^M
      End
      Begin VB.Menu cizgi19 
         Caption         =   "-"
      End
      Begin VB.Menu mnukaydet 
         Caption         =   "Mesajý Kaydet"
         Shortcut        =   ^S
      End
      Begin VB.Menu cizgi20 
         Caption         =   "-"
      End
      Begin VB.Menu daha_sonra_gonder 
         Caption         =   "Daha Sonra Gönder"
      End
      Begin VB.Menu cizgi1 
         Caption         =   "-"
      End
      Begin VB.Menu mesaj_iptal 
         Caption         =   "Bu Mesajý Ýptal Et"
      End
      Begin VB.Menu cizgi2 
         Caption         =   "-"
      End
      Begin VB.Menu yazdir 
         Caption         =   "Yazdýr"
         Begin VB.Menu mesaji_yazdir 
            Caption         =   "Mesajý Olarak Yazdýr"
            Shortcut        =   ^Y
         End
         Begin VB.Menu cizgi 
            Caption         =   "-"
         End
         Begin VB.Menu ekrani_yazdir 
            Caption         =   "Form Olarak Yazdýr"
            Shortcut        =   ^E
         End
      End
      Begin VB.Menu cizgi3 
         Caption         =   "-"
      End
      Begin VB.Menu Kapat 
         Caption         =   "Kapat"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "yeni_mesaj_gonder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private degisiklik_kontrol As Boolean

Private Sub kim_gonderdi()
'Mesajý Gönderen Kiþinin Bilgileri Okunuyor.
'On Error Resume Next
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from kullanicilar WHERE kullanici_adi ='" & Label17.Caption & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
                           imza.Text = suz![adi_soyadi] & vbCrLf & suz![aciklama]
suz.Close
On Error Resume Next
Dim resim
resim = programayarlari.Text4.Text & Label17.Caption & ".jpg"
Picture3.Picture = LoadPicture(resim)
Else
End If
End Sub
Private Sub sorgula()
'Ýncelenen Kaydý Mesaj Numarasýna Göre Sorgula
'Önemli Uyarý
'Bu Deðerleri Olan Anahtarlar Hata veriyor bunu önlemek için _
on error resume  next kodunu kullanmak gerekir.
On Error Resume Next
Dim conn As New ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim suz As New Recordset
suz.Open "Select * from mesajlar WHERE id ='" & Label16.Caption & "'", conn, adOpenKeyset, adLockOptimistic
If suz.RecordCount <> 0 Then
                           Label17.Caption = suz![kimden]
                           Label18.Caption = suz![kime]
                           Label19.Caption = suz![konu]
                           Label20.Caption = suz![gonderim_tarihi]
                           Label22.Caption = suz![okundu_tarih]
                           Label24.Caption = suz![txtbilgi]
                           Label25.Caption = suz![atac]
                           Label27.Caption = suz![bilgi]
                           Check1.Value = suz![acil]
                           Check2.Value = suz![inceleyin]
                           Check3.Value = suz![bilgilendirme]
                           Check4.Value = suz![yanitlayin]
                           RichTextBox1 = suz![mesaj]
                           
suz.Close
Else
'MsgBox "Mesaj Okuma Hatasý..", vbCritical, "Hata Kodu (C400)"
End If
'Ýmza Yükleniyor.
'***********************************************************************************
 ' yeni_mesaj_gonder.imza.Text = yeni_mesaj_gonder.Label24.Caption & vbCrLf & programayarlari.Text3.Text
'***********************************************************************************
End Sub

Private Sub Command1_Click()
On Error GoTo hata
If Text1.Text = "" Or Text3.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!" & vbCrLf & "Kime ; Konu ; Mesaj ;", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = yeni_mesaj_gonder.Text2.Text ' Deðiþken Satýr.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then
Label8.Caption = Label8.Caption + 1 ' Mesaj Numarasý
rs.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew ' 1 DEN 16 YA KADAR GERI SAYIM ISLEMINCE BILINMEYEN BIR DENKLEM
                    rs!id = Label8.Caption
                    rs!mesajid = Label10.Caption
                    rs!kimden = StatusBar1.Panels.Item(2).Text
                    rs!txtbilgi = StatusBar1.Panels.Item(3).Text
                    rs!kime = se_next_to
                    rs!konu = Text3.Text
                    rs!gonderim_tarihi = StatusBar1.Panels.Item(5).Text & StatusBar1.Panels.Item(7).Text
                    rs!okundu = "0"
                    rs!silindi = "0"
                    rs!gonderilen = "1" 'Eðer Sürekli Sen Gönderiyosan Gönderildi Olur.
                    rs!gonderilmedi = "0"
                    rs!Ek = Text4.Text
                    rs!acil = Check1.Value
                    rs!inceleyin = Check2.Value
                    rs!bilgilendirme = Check3.Value
                    rs!yanitlayin = Check4.Value
                    rs!okundu_tarih = "0"
                    rs!mesaj = "Bu Ýletiyi Bilgi Mesajý Olarak Aldýnýz" & RichTextBox1.Text
rs.Update
rs.Close
End If
Loop
'MsgBox " '' " & Text3.Text & " ''" & vbCrLf & vbCrLf & "    Mesajýnýz Baþarý ile Gönderildi...", vbInformation, "Tamamlandý."
gelen_mesajlar
okunmamis_mesajlar
Unload Me
End If
Exit Sub
hata:
'MsgBox "Bilinmeyen Bir Sistem Hatasý Meydana Geldi.Lütfen Sistem Yöneticinize Baþvurun.", vbCritical, "Hata"
Exit Sub
End Sub

Private Sub Command2_Click()

If Text1.Text = "" Or Text3.Text = "" Or RichTextBox1.Text = "" Then
MsgBox "Boþ Bilgi Giriþi Yapýyorsunuz.!" & vbCrLf & "Kime ; Konu ; Mesaj ;", vbCritical, "Hata Kodu (C300)"
Else
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
dosya_gonder
'Önce Dosya Transfer Ýþlemini Baþlatmamýz Gerekiyor.
'Eðer Dosya Kopyalamada Hata Meydana Gelirse Böylelikle Mesaj Defalarca
'Kayýt Edilmez........
Dim se_next_to As String
Dim se_mail_to As String
se_email_to = yeni_mesaj_gonder.Text1.Text ' Deðiþken Satýr.
Do While InStr(1, se_email_to, ";") <> 0
npos = InStr(1, se_email_to, ";")
se_next_to = Trim(Left(se_email_to, npos - 1))
se_email_to = Trim(Mid(se_email_to, npos + 1))
If se_next_to <> "" Then

yeni_mesaj_gonder.Label8.Caption = yeni_mesaj_gonder.Label8.Caption + 1 ' Mesaj Numarasý

rs.Open "Select * from mesajlar", conn, adOpenKeyset, adLockPessimistic
rs.AddNew ' 1 DEN 16 YA KADAR GERI SAYIM ISLEMINCE BILINMEYEN BIR DENKLEM
                    rs!id = yeni_mesaj_gonder.Label8.Caption
                    rs!mesajid = yeni_mesaj_gonder.Label10.Caption
                    rs!kimden = yeni_mesaj_gonder.StatusBar1.Panels.Item(2).Text
                    rs!txtbilgi = "(" & StatusBar1.Panels.Item(3).Text & ")"
                    rs!kime = se_next_to
                    rs!konu = yeni_mesaj_gonder.Text3.Text
                    rs!gonderim_tarihi = yeni_mesaj_gonder.StatusBar1.Panels.Item(5).Text & StatusBar1.Panels.Item(7).Text
                    rs!okundu = "0"
                    rs!silindi = "0"
                    rs!gonderilen = "1" 'Eðer Sürekli Sen Gönderiyosan Gönderildi Olur.
                    rs!gonderilmedi = "0"
                    
                    If Text4.Text = "" Then
                    rs!Ek = "0" 'Eðer Ek Yoksa 0 Deðeri
                    rs!atac = "0" 'Ek Yoksa Dosya Kaydý "0" Olacak
                    Else
                    rs!Ek = "1" 'Eðer Ek Warsa 1 Deðeri
                    rs!atac = yeni_mesaj_gonder.Text4.Text
                    End If
                    rs!acil = yeni_mesaj_gonder.Check1.Value
                    rs!inceleyin = yeni_mesaj_gonder.Check2.Value
                    rs!bilgilendirme = yeni_mesaj_gonder.Check3.Value
                    rs!yanitlayin = yeni_mesaj_gonder.Check4.Value
                    rs!okundu_tarih = "okunmadý"
                    rs!mesaj = yeni_mesaj_gonder.RichTextBox1.Text
                    
                    
                    
rs.Update
rs.Close
End If
Loop
MsgBox " '' " & Text3.Text & " ''" & vbCrLf & vbCrLf & "    Mesajýnýz Baþarý ile Gönderildi...", vbInformation, "Tamamlandý."
Command1_Click ' Bilgi Gödnermek Ýstediðiniz Kiþilere Gidiyor.
gelen_mesajlar
okunmamis_mesajlar
sol_menu
Unload Me
End If
'Exit Sub
'hata:
'MsgBox "Bilinmeyen Bir Sistem Hatasý Meydana Geldi.Lütfen Sistem Yöneticinize Baþvurun." & vbCrLf & "Çok Fazla Ek Göndermeye Çalýþtýnýz.", vbCritical, "Hata"
'Exit Sub
End Sub


Private Sub Command3_Click()
'Yazýcý
Dim HorizontalMargin As Long, VerticalMargin As Long
Dim soru
soru = MsgBox("Kullandýðýnýz Sürücü : " & Printer.DeviceName, vbInformation + vbYesNo, "Yazýcý Seçimi")
'
If soru = vbNo Then GoTo bitir
Printer.ScaleMode = vbMillimeters
HorizontalMargin = (230 - Printer.ScaleWidth) / 2
VerticalMargin = (297 - Printer.ScaleHeight) / 2
HorizontalMargin = 5 + HorizontalMargin
VerticalMargin = 5 + VerticalMargin
'
Printer.FontName = "Arial TUR"
Printer.FontSize = 10
Printer.FontBold = True
Printer.FontItalic = True
Printer.FontUnderline = False
Printer.FontStrikethru = False
Printer.ForeColor = RGB(0, 0, 2)
Printer.FillStyle = 1
'
Printer.Print Space(75) & "ÝÇ HABERLEÞME FORMU"
Printer.CurrentY = VerticalMargin + 25
'
Printer.FontName = "Arial TUR"
Printer.FontSize = 10
Printer.FontBold = False
Printer.FontItalic = False
Printer.FontUnderline = False
Printer.FontStrikethru = False
Printer.ForeColor = RGB(0, 0, 0)
Printer.FillStyle = 1
'
Printer.Print Label11 & " : " & Label16.Caption ' Mesaj No
Printer.CurrentY = VerticalMargin + 30
Printer.Line (16, 16)-(15, 18)

'
Printer.Print Label12 & " : " & Label17.Caption & "   " & Label24.Caption 'Kimden
Printer.CurrentY = VerticalMargin + 33
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print Label13 & " : " & Label18.Caption 'Kime
Printer.CurrentY = VerticalMargin + 39
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print Label14 & " : " & Label19.Caption 'Konu
Printer.CurrentY = VerticalMargin + 42
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print Label15 & " : " & Label20.Caption & Space(20) & Label21 & " : " & Label22.Caption
Printer.CurrentY = VerticalMargin + 45
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print Check1.Caption & ":   " & "[" & Check1.Value & "]" & Space(10) & Check2.Caption & ":   " & "[" & Check2.Value & "]" & Space(10) & Check3.Caption & ":   " & "[" & Check3.Value & "]" & Space(10) & Check4.Caption & ":   " & "[" & Check4.Value & "]"
Printer.CurrentY = VerticalMargin + 51
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print "" ' Boþluk vermek için.
Printer.Print "------------------------------- Mesaj -------------------------------"
Printer.CurrentY = VerticalMargin + 63
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print RichTextBox1.Text 'Mesaj
Printer.CurrentY = VerticalMargin + 66
Printer.Print "" ' Boþluk vermek için.
'
Printer.Print "------------------------------- Mesaj Sonu -------------------------"
Printer.CurrentY = VerticalMargin + 69
Printer.Print "" ' Boþluk vermek için.
Printer.EndDoc

bitir:
End Sub

Private Sub daha_sonra_gonder_Click()
MsgBox "Bu Mesaj Daha Sonra Gönderilemez.!", vbCritical, "Uyarý"

End Sub

Private Sub ekrani_yazdir_Click()
yazici_ekrani.Label6.Caption = Label17.Caption & " " & Label24.Caption 'Kimden
yazici_ekrani.Label1.Caption = "00" & Label16.Caption     'Mesaj Id
yazici_ekrani.Label7.Caption = Label18.Caption      'Kime
yazici_ekrani.Label8.Caption = Label27.Caption      'Bilgi
yazici_ekrani.Label9.Caption = Label19.Caption      'Konu
yazici_ekrani.Label10.Caption = RichTextBox1.Text   'Mesaj
yazici_ekrani.Label11.Caption = Label20.Caption     'Tarih
yazici_ekrani.Label14.Caption = program_hakkinda.Label7.Caption 'Firma
yazici_ekrani.Check1.Value = Check1.Value       'Acil
yazici_ekrani.Check2.Value = Check4.Value       'Yanýtlayýn
yazici_ekrani.Check3.Value = Check3.Value       'Bilgilendirme
yazici_ekrani.Check4.Value = Check2.Value       'Ýnceleyin
yazici_ekrani.Text1.Text = imza.Text
yazici_ekrani.Picture1.Picture = Picture3.Picture
yazici_ekrani.Show

End Sub

Private Sub Form_Activate()
sorgula
kim_gonderdi
End Sub

Private Sub Form_Load()
StatusBar1.Panels.Item(2).Text = Form1.kullanici_adi.Caption
StatusBar1.Panels.Item(3).Text = Form1.Label8.Caption
StatusBar1.Panels.Item(5).Text = Date & " "
StatusBar1.Panels.Item(7).Text = Time
'Ýmzayý Göster
'yeni_mesaj_gonder.imza.Text = Space(7) & Form1.Label8.Caption & vbCrLf & Form1.Label9.Caption
kontor ' mesaj Numarasý Olusturmak için...

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &H80000012
Label2.ForeColor = &H80000012
End Sub

Private Sub Form_Unload(Cancel As Integer)
'K A P A T  ( AMA ONCE KONTROL ET )
'If degisiklik_kontrol Then
'If MsgBox("Mesaj Göderilmedi. Bu Mesajý Saklamak Ýstiyormusunuz?", vbInformation + vbYesNo, "Mesaj Kayýt") = vbYes Then
'MsgBox "Evet Dedi."
'degisiklik_kontrol = False
'Else
'Unload Me ' Hayýr dedi.
'End If
'End If
'Eðer Hiç Biþ Yapmadýyda Direk olacak pencereyi Kapat
'Unload Me

End Sub


Private Sub Kapat_Click()
Unload Me
End Sub

Private Sub Label1_Click()
kisi_sec.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HFF&

End Sub


Private Sub Label2_Click()
kisi_sec.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &HFF&
End Sub


Private Sub Label23_Change()
If Label23.Caption = "Tümünü kaydet" Then
ek_kaydet.Label3.Caption = yeni_mesaj_gonder.Label25.Caption
ek_kaydet.Show
Else
End If
'Navigate Me, Label17.Caption
End Sub

Private Sub mesaj_iptal_Click()
Unload Me
End Sub

Private Sub mesaji_gonder_Click()
Command2_Click
End Sub

Private Sub mesaji_yazdir_Click()
Command3_Click
End Sub

Private Sub mnukaydet_Click()
On Error Resume Next
'Richtextbox ý texte kopyala
Text5.Text = RichTextBox1.Text
Dim sFile As String
With cmd2
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
Print #1, Text6.Text
Print #1, "Mesaj No: " & Label16.Caption
Print #1, Text6.Text
Print #1, "Kimden: " & Label17.Caption & " " & Label24.Caption
Print #1, Text6.Text
Print #1, "Kime: " & Label18.Caption
Print #1, Text6.Text
Print #1, "Gönderim Tarihi : " & Label20.Caption
Print #1, Text6.Text
Print #1, "Okunma Tarihi :" & Label22.Caption
Print #1, Text6.Text
Print #1, "Konu :" & Label19.Caption
Print #1, Text6.Text
Print #1, "****************************Mesajý***************************"
Print #1, Text6.Text
Print #1, Text5.Text
Print #1, Text6.Text
Print #1, "****************************Bitti***************************"
Print #1, Text6.Text
Print #1, Text6.Text
Print #1, "Bu Mesajý " & Date; " Tarihinde Dosya Olarak Arþivlediniz."
Close #1
MsgBox " Kayýtlar baþarý ile aktarýlmýþtýr..", vbInformation, "Tamamlandý."
End With
End Sub

Private Sub RichTextBox1_Change()
degisiklik_kontrol = True
End Sub

Private Sub sil_Click()
Dim soru
soru = MsgBox("Bu Mesaj Silinecek Eminmisiniz.?", vbQuestion + vbYesNo, "Dikkat")
If soru = vbYes Then
If Label16.Caption = "" Then
MsgBox "Silinecek Mesajý Seçmediniz.", vbCritical, "Uyarý"
Else
Form1.Text3.Text = "1"
Dim conn As New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.Jet.OleDB.4.0;Data Source=" & programayarlari.Text1.Text & "\veri.mdb"
conn.Open
Dim rs As New ADODB.Recordset
rs.Open " select * from mesajlar where id = '" & Label16.Caption & " '", conn, adOpenKeyset, adLockPessimistic
If rs.RecordCount <> 0 Then
                                rs![silindi] = Form1.Text3.Text
                                rs.Update
                                rs.Close
End If
Form1.Text1.Text = ""
Form1.Text2.Text = ""
Form1.Text3.Text = ""
Form1.Text4.Text = ""
Form1.Label17.Caption = ""
'Sabitleri Yenilemek Ýçin Tekrar Caðýr.
gelen_mesajlar 'Ýþlem Birtince Gelen Mesajlarý Yenile
sol_menu 'Sol taraftaki Menüyüde...
MsgBox "Mesajýnýz Silindi.", vbInformation, "Silindi"
gelen_mesajlar 'Ýþlem Birtince Gelen Mesajlarý Yenile
okunmamis_mesajlar
Unload Me
End If
End If 'sil

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
'Eðer Kiþi Adý El iþe girilirse sonuna noktali virgül koy
Text1.Text = Text1.Text & ";"
Text2.SetFocus
End If
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text3.SetFocus
End If

End Sub


Private Sub Text3_Change()
yeni_mesaj_gonder.Caption = Text3.Text
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.SetFocus
End If
End Sub


Private Sub Text4_DblClick()
'Eðer Ekler Bölümüne Ç,ft Týknýrsa Litesi Sil.
Text4.Text = ""
Text4.Visible = False
Label4.Visible = False
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then 'entere basýnca
RichTextBox1.SetFocus
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "a"
'Once Metni Hazýrla ve Gideccek Sekilde Kalsýn.
'------------------------------------------------------------
Command2_Click ' Mesaj Gondermke için Command_Butonu Týklýyor
'------------------------------------------------------------
Case "b" ' Yanýtla.
Dim cizgi
cizgi = "------------------------Ýleti Bilgisi--------------------"
'Yanýtla...
'Gerekli Bilgileri Aktarýyoruz...
mesaji_ilet.Caption = "Mesaj Ýletiliyor."
mesaji_ilet.Text1.Text = yeni_mesaj_gonder.Label17.Caption & ";"
mesaji_ilet.Text3.Text = "Ynt:> " & yeni_mesaj_gonder.Label19.Caption
mesaji_ilet.RichTextBox1.Text = vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & cizgi & vbCrLf & vbCrLf & yeni_mesaj_gonder.RichTextBox1.Text
mesaji_ilet.Show
'
Case "c"
'Ekle
cmd1.CancelError = True
On Error GoTo hata
cmd1.Action = 1
Text4.Text = Text4.Text + cmd1.FileTitle & ";"
Text4.Visible = True 'Gizli Olanlarý Göster
Label4.Visible = True ' Gizli Olanlarý Göster
Exit Sub
hata:
Exit Sub
Case "d" 'Yazdýr.
Dim soru
soru = MsgBox("Yazdýrma Ýþlemi Baþlayacak Eminmisiniz.!", vbInformation + vbYesNo, "Yazdýr.")
If soru = vbYes Then
ekrani_yazdir_Click
Else
End If
'Command3_Click
Case "e"
sil_Click
Case "f"
kisi_sec.Show
Case "g"
Unload Me
End Select

End Sub

Private Sub Toolbar2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Label23.Caption = ButtonMenu.Text
End Sub


Private Sub yeni_mesaj_Click()
'Ayný Fordan Yaratmak ( Ayný Anda Mesaj Gönderebilmek Ýçin)
Dim yeni_mesaj_gonder As New yeni_mesaj_gonder
'yeni_mesaj_gonder.imza.Text = Space(7) & Form1.Label8.Caption & vbCrLf & Form1.Label9.Caption
'
yeni_mesaj_gondert
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(3).Visible = False
'yeni_mesaj_gonder.Toolbar1.Buttons.Item(4).Visible = False

'
yeni_mesaj_gonder.Show
End Sub
Private Sub yeni_mesaj_gondert()
Dim resim
yeni_mesaj_gonder.Toolbar1.Buttons.Item(3).Visible = False
yeni_mesaj_gonder.Toolbar1.Buttons.Item(4).Visible = False
'Eðer Resim Yok Ýse
On Error GoTo hata
resim = programayarlari.Text4.Text & Form1.kullanici_adi & ".jpg"
yeni_mesaj_gonder.Picture3.Picture = LoadPicture(resim)
yeni_mesaj_gonder.imza.Text = Label8.Caption & vbCrLf & Label9.Caption
'yeni_mesaj_gonder.Show
Exit Sub
hata:
yeni_mesaj_gonder.imza.Text = Label8.Caption & vbCrLf & Label9.Caption
yeni_mesaj_gonder.Picture3.Print "Ýmzanýz Yok."
Exit Sub
End Sub

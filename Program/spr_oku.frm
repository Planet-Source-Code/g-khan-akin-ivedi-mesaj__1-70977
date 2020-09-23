VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form spr_mesaj_oku 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ivedi Mesaj Wiever 1.0 Build 110"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "spr_oku.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   9045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "yazdýr"
      Height          =   330
      Left            =   630
      TabIndex        =   4
      Top             =   6795
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "aç"
      Height          =   330
      Left            =   180
      TabIndex        =   3
      Top             =   6795
      Visible         =   0   'False
      Width           =   420
   End
   Begin MSComDlg.CommonDialog cmd1 
      Left            =   1260
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   9015
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   675
      Top             =   5400
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
            Picture         =   "spr_oku.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "spr_oku.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "spr_oku.frx":0CE6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   1005
      ButtonWidth     =   953
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aç"
            Key             =   "a"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Yazdýr"
            Key             =   "b"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Kapat"
            Key             =   "c"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6585
      Left            =   0
      TabIndex        =   0
      Top             =   675
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11615
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"spr_oku.frx":13E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "spr_mesaj_oku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim sFile As String
  '  If ActiveForm Is Nothing Then LoadNewDoc
      With cmd1
        .DialogTitle = "Aç"
        .CancelError = False

        .Filter = "TXT Dosya Türü (*.txt)|*.txt"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    spr_mesaj_oku.RichTextBox1.LoadFile sFile
    spr_mesaj_oku.Caption = sFile
End Sub

Private Sub Command2_Click()
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

        
          spr_mesaj_oku.RichTextBox1.SelPrint .hDC
        'End If
    End With
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "a" ' Aç
Command1_Click
Case "b" 'Yazdýr
Command2_Click
Case "c" 'Kapat
Unload Me
End Select
End Sub

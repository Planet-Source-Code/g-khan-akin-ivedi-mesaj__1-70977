VERSION 5.00
Begin VB.Form yazici_ekrani 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   6195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11460
   Icon            =   "logo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   8865
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   5265
      Width           =   2445
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   8865
      ScaleHeight     =   1050
      ScaleWidth      =   2445
      TabIndex        =   18
      Top             =   4140
      Width           =   2445
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1935
      Top             =   4995
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H8000000E&
      Caption         =   "Ýnceleyin"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9585
      TabIndex        =   15
      Top             =   1440
      Width           =   1725
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H8000000E&
      Caption         =   "Bilgilendirme"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9585
      TabIndex        =   12
      Top             =   1125
      Width           =   1725
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H8000000E&
      Caption         =   "Yanýtlayýn"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9585
      TabIndex        =   11
      Top             =   810
      Width           =   1725
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H8000000E&
      Caption         =   "Acil"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9585
      TabIndex        =   10
      Top             =   495
      Width           =   1725
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
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
      Height          =   285
      Left            =   3645
      TabIndex        =   17
      Top             =   5850
      Width           =   3480
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "KY-4-036 (R1)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   5895
      Width           =   1995
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÝÇ HABERLEÞME FORMU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2250
      TabIndex        =   14
      Top             =   90
      Width           =   7125
   End
   Begin VB.Label Label11 
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
      Left            =   9270
      TabIndex        =   13
      Top             =   1800
      Width           =   2085
   End
   Begin VB.Label Label10 
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
      Height          =   3750
      Left            =   405
      TabIndex        =   9
      Top             =   2295
      Width           =   10950
   End
   Begin VB.Line Line1 
      X1              =   -1125
      X2              =   11430
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Shape Shape1 
      Height          =   6180
      Left            =   0
      Top             =   0
      Width           =   11445
   End
   Begin VB.Label Label9 
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
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   1710
      Width           =   7935
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
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1395
      Width           =   7935
   End
   Begin VB.Label Label7 
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
      Height          =   555
      Left            =   1080
      TabIndex        =   6
      Top             =   810
      Width           =   7935
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
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Top             =   495
      Width           =   7935
   End
   Begin VB.Label Label5 
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
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   1710
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   1395
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
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
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Top             =   810
      Width           =   1005
   End
   Begin VB.Label Label2 
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
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9675
      TabIndex        =   0
      Top             =   45
      Width           =   1635
   End
End
Attribute VB_Name = "yazici_ekrani"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
'yazici_ekrani.PrintForm
'MsgBox "Ýç Haberleþme Formu Yazdýrýldý.", vbInformation, "Bitti"
'Timer1.Enabled = False
'Unload Me
End Sub

Private Sub Form_Click()
MsgBox "Bu Formu Yazdýrmak için [ F2 ] Forman Çýkmak Ýçin [ ESC ] Tuþuna Basýn.", vbInformation, "Bilgi."

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'******************************************************
'F2 Tuþuna Basýlýrsa Yeni Mesaj
If KeyCode = vbKeyF2 Then
yazici_ekrani.PrintForm
MsgBox "Ýç Haberleþme Formu Yazdýrýldý.", vbInformation, "Bitti"
Unload Me
End If
'******************************************************
'******************************************************
'F2 Tuþuna Basýlýrsa Yeni Mesaj
If KeyCode = vbKeyEscape Then
Unload Me
End If
'******************************************************
End Sub

Private Sub Label10_Click()
MsgBox "Bu Formu Yazdýrmak için [ F2 ] Forman Çýkmak Ýçin [ ESC ] Tuþuna Basýn.", vbInformation, "Bilgi."
End Sub



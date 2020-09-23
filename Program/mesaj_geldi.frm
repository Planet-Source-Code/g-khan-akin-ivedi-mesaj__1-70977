VERSION 5.00
Begin VB.Form mesaj_geldii 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Yeni Mesajýnýz War...."
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   15
   ClientWidth     =   15360
   Icon            =   "mesaj_geldi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   330
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer5 
      Interval        =   5000
      Left            =   9855
      Top             =   45
   End
   Begin VB.CommandButton Command1 
      Caption         =   "dýkla"
      Height          =   240
      Left            =   11655
      TabIndex        =   3
      Top             =   45
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   9405
      Top             =   45
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   8955
      Top             =   45
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8505
      Top             =   45
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8055
      Top             =   45
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Yeni mesajýnýz var."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   3690
      TabIndex        =   2
      Top             =   90
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3150
      TabIndex        =   1
      Top             =   90
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   15075
      Picture         =   "mesaj_geldi.frx":06EA
      Top             =   45
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   90
      Picture         =   "mesaj_geldi.frx":0A74
      Top             =   0
      Width           =   360
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   1305
      TabIndex        =   0
      Top             =   90
      Width           =   2130
   End
End
Attribute VB_Name = "mesaj_geldii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40


Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long)

Private Sub Command1_Click()
'5 sn sonra formu gizle
mesaj_geldii.Height = mesaj_geldii.Height - 30
If mesaj_geldii.Height = 90 Then
Timer4.Enabled = False
Unload Me
End If
End Sub

Private Sub Form_Activate()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE _
Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_DblClick()
Form1.WindowState = 2

End Sub


Private Sub Form_Load()
Label2.Caption = Form1.Label19.Caption
Label1.Caption = Form1.Label8.Caption
mesaj_geldii.Top = 10 ' en üstte acýl
End Sub

Private Sub Image2_Click()
Unload Me
End Sub

Private Sub Label3_DblClick()
Form1.WindowState = 2
End Sub


Private Sub Timer1_Timer()
' + olarak
Image1.Left = Image1.Left + 10
If Image1.Left = 800 Then
Timer1.Enabled = False 'iþlemi durdur.
Timer2.Enabled = True
Else
Timer1.Enabled = True 'iþlemi durdur.
End If
End Sub
Private Sub Timer2_Timer()
'   - olarak
Image1.Left = Image1.Left - 10
If Image1.Left = 90 Then
Timer2.Enabled = False 'iþlemi durdur.
Timer1.Enabled = True
Else
Timer2.Enabled = True 'iþlemi durdur.
End If
End Sub

Private Sub Timer4_Timer()
'Formu Kapatmaya baþla
Command1_Click
End Sub


Private Sub Timer5_Timer()
'5 sn sonra form yawas yawas kapanmaya baslayacak
Timer4.Enabled = True
Timer5.Enabled = False
End Sub



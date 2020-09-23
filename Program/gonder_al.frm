VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form gonder_al 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mesajlar Alýnýyor..."
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "gonder_al.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   2070
      TabIndex        =   3
      Top             =   2475
      Width           =   1275
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4095
      Top             =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tekrar Dene"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   2745
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   240
      Left            =   180
      TabIndex        =   0
      Top             =   855
      Visible         =   0   'False
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " Kontrol Ediliyor."
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
      Left            =   180
      TabIndex        =   1
      Top             =   540
      Width           =   1680
   End
End
Attribute VB_Name = "gonder_al"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Counter As Integer
   Dim Workarea(2850) As String
   ProgressBar1.Min = LBound(Workarea)
   ProgressBar1.Max = UBound(Workarea)
   ProgressBar1.Visible = True

'Set the Progress's Value to Min.
   ProgressBar1.Value = ProgressBar1.Min

'Loop through the array.
   For Counter = LBound(Workarea) To UBound(Workarea)
      'Set initial values for each item in the array.
      Workarea(Counter) = "Tamamlandý." & Counter
      ProgressBar1.Value = Counter
   Next Counter
   ProgressBar1.Visible = False
   ProgressBar1.Value = ProgressBar1.Min
Unload Me
End Sub

Private Sub Command2_Click()
gelen_mesajlar
Form1.Text8.Visible = False
sol_menu
Form1.Text8.Refresh
Form1.Label18.Caption = ""
Command1_Click
Unload Me
Form1.Text8.Visible = True
End Sub

Private Sub Timer1_Timer()
Command2_Click
End Sub



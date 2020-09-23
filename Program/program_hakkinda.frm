VERSION 5.00
Begin VB.Form program_hakkinda 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   LinkTopic       =   "Form2"
   ScaleHeight     =   6255
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Tamam"
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
      Left            =   3690
      TabIndex        =   5
      Top             =   5670
      Width           =   1320
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "1652-2548-2598-5285-5254"
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
      Height          =   240
      Left            =   180
      TabIndex        =   9
      Top             =   4050
      Width           =   4740
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Lisans Numarasý:"
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
      Left            =   180
      TabIndex        =   8
      Top             =   3735
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Gökhan AKIN - 549 329 04 23"
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
      Height          =   240
      Left            =   180
      TabIndex        =   7
      Top             =   3420
      Width           =   4740
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lisanslý Kullanýcý:"
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
      Left            =   180
      TabIndex        =   6
      Top             =   3105
      Width           =   1905
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000001&
      Height          =   4155
      Left            =   5040
      TabIndex        =   4
      Top             =   1890
      Width           =   105
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000001&
      Height          =   4200
      Left            =   0
      TabIndex        =   3
      Top             =   1890
      Width           =   60
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000001&
      Caption         =   "Label3"
      Height          =   330
      Left            =   -1080
      TabIndex        =   2
      Top             =   6030
      Width           =   6495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ver.1.0 Build 101"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1935
      TabIndex        =   1
      Top             =   2700
      Width           =   1545
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   270
      Picture         =   "program_hakkinda.frx":0000
      Top             =   2340
      Width           =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ýç Haberleþme"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   810
      TabIndex        =   0
      Top             =   2385
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   1920
      Left            =   0
      Picture         =   "program_hakkinda.frx":06EA
      Top             =   0
      Width           =   5130
   End
End
Attribute VB_Name = "program_hakkinda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub


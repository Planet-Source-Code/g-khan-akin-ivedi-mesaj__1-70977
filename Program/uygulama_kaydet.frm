VERSION 5.00
Begin VB.Form uygulama_calistir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uyarý"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   Icon            =   "uygulama_kaydet.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -315
      TabIndex        =   3
      Top             =   3330
      Width           =   5415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Kapat"
      Height          =   375
      Left            =   3420
      TabIndex        =   2
      Top             =   3510
      Width           =   1320
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Tamam"
      Height          =   375
      Left            =   2070
      TabIndex        =   1
      Top             =   3510
      Width           =   1320
   End
   Begin VB.Label Label5 
      Caption         =   "Yinede çalýþtýrmak istiyormusunuz?"
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
      Left            =   450
      TabIndex        =   7
      Top             =   1845
      Width           =   4110
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   3510
      Picture         =   "uygulama_kaydet.frx":06EA
      Stretch         =   -1  'True
      Top             =   90
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "........................................................................."
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
      Left            =   1215
      TabIndex        =   6
      Top             =   2835
      Width           =   3345
   End
   Begin VB.Label Label3 
      Caption         =   "Çalýþtýrýlacak Dosya:"
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
      TabIndex        =   5
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Uygulama Çalýþtýrýlacak..."
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
      Left            =   90
      TabIndex        =   4
      Top             =   360
      Width           =   2625
   End
   Begin VB.Label Label1 
      Caption         =   $"uygulama_kaydet.frx":643C
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   450
      TabIndex        =   0
      Top             =   1170
      Width           =   4110
   End
End
Attribute VB_Name = "uygulama_calistir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SW_SHOW = 1

Private Declare Function ShellExecute Lib _
"shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, _
ByVal lpOperation As String, _
ByVal lpFile As String, _
ByVal lpParameters As String, _
ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Sub Command1_Click()
'Seçilen Dosyayý Temp Klasörünün içeine Gönder ve Ordan Çalýþtýr.
Navigate Me, Label4.Caption
End Sub

Private Sub Command2_Click()
Unload Me
End Sub



Private Sub Navigate(frm As Form, ByVal WebPageURL As String)
     Dim kasif As Long
     kasif = ShellExecute(frm.hwnd, "open", WebPageURL, "", "", SW_SHOW)


End Sub


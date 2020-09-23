VERSION 5.00
Begin VB.Form kullanici_bilgileri 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kullanýcý Bilgileri"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "kullanici_bilgileri.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Kapat"
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
      Left            =   3375
      TabIndex        =   13
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   -270
      TabIndex        =   12
      Top             =   4050
      Width           =   5775
   End
   Begin VB.Frame Frame1 
      Height          =   60
      Left            =   -45
      TabIndex        =   7
      Top             =   2835
      Width           =   5460
   End
   Begin VB.Label Label16 
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
      Left            =   1350
      TabIndex        =   18
      Top             =   2385
      Width           =   3165
   End
   Begin VB.Label Label15 
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
      Left            =   1350
      TabIndex        =   17
      Top             =   2070
      Width           =   3165
   End
   Begin VB.Label Label14 
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
      Left            =   1350
      TabIndex        =   16
      Top             =   1755
      Width           =   3165
   End
   Begin VB.Label Label13 
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
      Left            =   1350
      TabIndex        =   15
      Top             =   1440
      Width           =   3165
   End
   Begin VB.Label Label12 
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
      Left            =   1350
      TabIndex        =   14
      Top             =   1125
      Width           =   3165
   End
   Begin VB.Label Label11 
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
      Left            =   1845
      TabIndex        =   11
      Top             =   3510
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Sisteme Kayýt Tarihiniz:"
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
      TabIndex        =   10
      Top             =   3510
      Width           =   1680
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Left            =   3375
      TabIndex        =   9
      Top             =   3105
      Width           =   915
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Þimdiye Kadar Alýdýðýnýz Toplam Mesaj Sayýsý:"
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
      TabIndex        =   8
      Top             =   3105
      Width           =   3210
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Durumu:"
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
      Left            =   225
      TabIndex        =   6
      Top             =   2385
      Width           =   1050
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Açýklama:"
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
      Left            =   225
      TabIndex        =   5
      Top             =   2070
      Width           =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Grubu:"
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
      Left            =   225
      TabIndex        =   4
      Top             =   1755
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Adý Soyadý:"
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
      Left            =   225
      TabIndex        =   3
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý ID:"
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
      Left            =   225
      TabIndex        =   2
      Top             =   1125
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4050
      Picture         =   "kullanici_bilgileri.frx":06EA
      Stretch         =   -1  'True
      Top             =   135
      Width           =   690
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kullanýcý Bilgileri"
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
      Left            =   135
      TabIndex        =   1
      Top             =   270
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "kullanici_bilgileri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
'Kullnýcý Bilgileri
Label12.Caption = Baglan.Label12.Caption ' Kullanýcý Kodu
Label13.Caption = Baglan.Label4.Caption  ' Kullanýcý Adý Soyadý
Label14.Caption = Baglan.Label13.Caption ' Grubu
Label15.Caption = Baglan.Label9.Caption  ' Açýklama
Label16.Caption = Baglan.Label11.Caption ' Kullanýcý Durumu
Label11.Caption = Baglan.Label10.Caption ' Sisteme Kayýt Tarihi
 
End Sub

Private Sub Label16_Change()
If Label16.Caption = "1" Then
'Eðer Deðer 1 ise aktif demek
Label16.Caption = Label16.Caption & " Aktif Durumda"
Else
'Label16.Caption = Label16.Caption & " Aktif Durumda Deðil"
End If
End Sub


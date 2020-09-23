VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form tum_mesajlar_listesi 
   Caption         =   "Tüm Mesajlar"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10365
   Icon            =   "tum_mesajlar_listesi.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8190
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ListView liste 
      Height          =   9690
      Left            =   90
      TabIndex        =   0
      Top             =   630
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   17092
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "tum_mesajlar_listesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

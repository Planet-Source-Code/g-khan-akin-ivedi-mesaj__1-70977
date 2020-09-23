Attribute VB_Name = "ApiModulleri"
Public Const WM_SYSCOMMAND = &H112
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Function sol_menu()
'Sol Menü yü Yükle
Form1.Treeview1.Nodes.Clear
okunmamis_mesajlar
arsiv_listesi
temizlikci
With Form1.Treeview1
 .BorderStyle = 1
   Dim nodX As Node
   Set nodX = .Nodes.Add(, , "d", "Mesajlar Menusu", 1, 1)
   Set nodX = .Nodes.Add("d", tvwChild, "d8", "Yerel Mesajlar", 2, 2)
   Set nodX = .Nodes.Add("d8", tvwChild, , "Gelen Mesajlar", 1, 1)
   Set nodX = .Nodes.Add("d8", tvwChild, , "Giden Mesajlar", 2, 2)
   'Set nodX = .Nodes.Add("d8", tvwChild, , "Gitmeyen Mesajlar", 3, 3)
   Set nodX = .Nodes.Add("d8", tvwChild, , "Silinmiþ Mesajlar", 4, 4)
   Set nodX = .Nodes.Add("d8", tvwChild, , "Arþiv", 5, 5)
Form1.Text8.Text = "(" & Form1.arsiv.ListItems.Count & ")" 'Arþiv Sayýsý"
nodX.EnsureVisible
End With
End Function
Function temizlikci()
Form1.Label3.Caption = ""
Form1.Label4.Caption = ""
Form1.Label6.Caption = ""
Form1.Toolbar2.Visible = False
Form1.RichTextBox1.Text = ""

End Function

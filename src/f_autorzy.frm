VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} f_autorzy 
   Caption         =   "TETRIS - autorzy"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "f_autorzy.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "f_autorzy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Powrót do menu
Private Sub Image_cofnij_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
    a_menu.Show
End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} a_menu 
   Caption         =   "TETRIS - menu g³ówne"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   OleObjectBlob   =   "a_menu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "a_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Lista wyboru poziomu gry
Private Sub UserForm_Initialize()
    a_menu.poziom.List = Array("Latwy", "Sredni", "Trudny")
    a_menu.poziom.ListIndex = 0
End Sub
'Uruchomienie gry
Private Sub Label_start_game_Click()
    Me.Hide
    Call UruchomGre
End Sub
'W³¹czenie UF ranking
Private Sub Label_ranking_Click()
    Me.Hide
    c_ranking.Show
End Sub
'W³¹czenie UF instrukcja
Private Sub Label_instrukcja_Click()
    Me.Hide
    d_instrukcja.Show
End Sub
'W³¹czenie UF autorzy
Private Sub Label_autorzy_Click()
    Me.Hide
    f_autorzy.Show
End Sub

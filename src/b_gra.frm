VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} b_gra 
   Caption         =   "TETRIS "
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "b_gra.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "b_gra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'sterowanie przyciskami z poziomu UF
Private Sub Image_lewo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WykonajRuch -1, 0, 0 'o 1 w lewo w poziomie
End Sub
Private Sub Image_prawo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WykonajRuch 1, 0, 0 'o 1 w prawo w poziomie
End Sub
Private Sub Image_dol_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WykonajRuch 0, 1, 0 'o 1 w dó³ w pionie
End Sub
Private Sub Image_rotacja_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    WykonajRuch 0, 0, 1 'obrót o 1
End Sub
Private Sub Image_szybkie_spadanie_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SzybkiSpadek
End Sub
Private Sub Image_pauza_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PrzelaczPauze
End Sub

Private Sub UserForm_Click()

End Sub

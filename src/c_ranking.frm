VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} c_ranking 
   Caption         =   "TETRIS - ranking"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6765
   OleObjectBlob   =   "c_ranking.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "c_ranking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RankingDane As Variant 'Tablica na nick, wynik i poziom
'Wyœwietlnie danych w rankingu
Private Sub UserForm_Activate()
    Dim i As Integer
    Dim IloscRekordow As Long
    Call Wczytaj_Ranking(RankingDane)
    'Sortowanie je¿eli s¹ dane
    If IsArray(RankingDane) Then
        Call Sortowanie_Babelkowe_Ranking(RankingDane)
        IloscRekordow = UBound(RankingDane, 1)
    Else
        IloscRekordow = 0
    End If
    'Wyœwietlenie wyników dla 5 miejsc
    For i = 1 To 5
        If i <= IloscRekordow Then
            'Je¿eli jest rekord to wpiwywanie danych
            Me.Controls("Label_nick" & i).Caption = RankingDane(i, 1)
            Me.Controls("Label_wynik" & i).Caption = Format(RankingDane(i, 2), "000000")
            Me.Controls("Label_poziom" & i).Caption = RankingDane(i, 3)
        Else
            'Je¿eli nie ma rekordów to --- lub 000000
            Me.Controls("Label_nick" & i).Caption = "---"
            Me.Controls("Label_wynik" & i).Caption = "000000"
            Me.Controls("Label_poziom" & i).Caption = "---"
        End If
    Next i
End Sub
Sub Wczytaj_Ranking(Lista As Variant)
    Dim ws As Worksheet
    Dim OstatniWiersz As Long
    Dim i As Long
    Set ws = Worksheets("ranking")
    'Znalezienie ostatniego zajêtego wiersza w kolumnie B nick
    OstatniWiersz = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    'Je¿eli nie ma rekordów w arkuszu tworzy pust¹ tablicê
    If OstatniWiersz < 2 Then
        Lista = Empty
        Exit Sub
    End If
    'Tworzenie tablicy, gdzie bêd¹ wiersze: 1 nick, 2 wynik, 3 poziom
    ReDim TablicaTymczasowa(1 To OstatniWiersz - 1, 1 To 3)
    For i = 2 To OstatniWiersz
        TablicaTymczasowa(i - 1, 1) = ws.Cells(i, 2).Value ' z kolumna B nick
        TablicaTymczasowa(i - 1, 2) = ws.Cells(i, 3).Value ' z kolumna C wynik
        TablicaTymczasowa(i - 1, 3) = ws.Cells(i, 4).Value ' z kolumna D poziom
    Next i
    Lista = TablicaTymczasowa
End Sub
Sub Sortowanie_Babelkowe_Ranking(Lista As Variant)
    Dim Pomocnicza As Variant
    Dim i As Long, j As Long, k As Integer
    For i = LBound(Lista, 1) To UBound(Lista, 1) - 1
        For j = i + 1 To UBound(Lista, 1)
            If Val(Lista(i, 2)) < Val(Lista(j, 2)) Then
                'Zamiana miejscami
                For k = 1 To 3
                    Pomocnicza = Lista(i, k)
                    Lista(i, k) = Lista(j, k)
                    Lista(j, k) = Pomocnicza
                Next k
            End If
        Next j
    Next i
End Sub
'powrot do menu g³ównego
Private Sub Image_cofnij_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
    a_menu.Show
End Sub


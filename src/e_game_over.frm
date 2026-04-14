VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} e_game_over 
   ClientHeight    =   2730
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5370
   OleObjectBlob   =   "e_game_over.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "e_game_over"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    'Sprawdzamy, czy gracz wcisn¹³ enter
    If KeyCode = 13 Then
        Dim NickGracza As String
        NickGracza = TextBox1.Text
        
        'jeœli gracz nic nie wpisa³, prosimy o nick
        If Trim(NickGracza) = "" Then
            MsgBox "Musisz podaæ swój nick, aby zapisaæ wynik!", vbExclamation, "Brak nicku"
            Exit Sub
        End If
        
        'Odwo³anie do rankingu
        Dim ArkRanking As Worksheet
        Set ArkRanking = ThisWorkbook.Sheets("ranking")
        
        'Znajdujemy pierwszy pusty wiersz w kolumnie nick
        Dim PustyWiersz As Long
        PustyWiersz = ArkRanking.Cells(ArkRanking.Rows.Count, "B").End(xlUp).Row + 1
        
        'Wpisujemy dane do tabelki
        ArkRanking.Cells(PustyWiersz, 1).Value = PustyWiersz - 1     ' Nr wiersza
        ArkRanking.Cells(PustyWiersz, 2).Value = NickGracza          ' Nick z pola tekstowego
        ArkRanking.Cells(PustyWiersz, 3).Value = Wynik               ' Punkty
        ArkRanking.Cells(PustyWiersz, 4).Value = a_menu.poziom.Text  ' Wybrany poziom z menu
        ArkRanking.Cells(PustyWiersz, 5).Value = Date                ' Dzisiejsza data
        
        'Powrót do menu
        Unload Me 'Zamyka menu i czyœci pamiêæ
        a_menu.Show
    End If
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Attribute VB_Name = "Module1_siatka"
'Wstawienie 10x20 formantów typu label jako siatka do gry (w trybie projektowania)
Sub GenerujSiatkeTetris()
'deklaracje zmiennych
    Dim UF As VBComponent
    Dim Klocek As MSForms.Label
    Dim Nr As Long
    Dim Wiersz As Long, Kolumna As Long
    Dim NrLinii As Long
'konfiguracja siatki
    Const Rozmiar As Integer = 20 ' szerokoœæ i wysokoœæ klocka
    Const MaxWierszy As Integer = 20
    Const MaxKolumn As Integer = 10
    Const MarginesLewy As Integer = 20
    Const MarginesGorny As Integer = 15
    Dim NazwaFormularza As String: NazwaFormularza = "b_gra"
' Wczytanie formularza
    On Error Resume Next
    Set UF = ThisWorkbook.VBProject.VBComponents(NazwaFormularza)
    On Error GoTo 0
' Czyszczenie
    Dim i As Long
    For i = UF.Designer.Controls.Count - 1 To 0 Step -1
        UF.Designer.Controls.Remove UF.Designer.Controls(i).Name
    Next i
    UF.CodeModule.DeleteLines 1, UF.CodeModule.CountOfLines
' Wstawianie formantów 10x20
    Nr = 1
    For Wiersz = 1 To MaxWierszy
    For Kolumna = 1 To MaxKolumn
' Wstawienie Labeli
    Set Klocek = UF.Designer.Controls.Add("Forms.Label.1")
    With Klocek
    .Name = "b_" & Wiersz & "_" & Kolumna
    .Caption = ""
    .Height = Rozmiar
    .Width = Rozmiar
    .Left = (Kolumna - 1) * Rozmiar
    .Top = (Wiersz - 1) * Rozmiar
    .Left = MarginesLewy + (Kolumna - 1) * Rozmiar
    .Top = MarginesGorny + (Wiersz - 1) * Rozmiar
    .BackColor = vbWhite
    .BorderStyle = fmBorderStyleSingle
    .BorderColor = RGB(220, 220, 220)
    End With
    Nr = Nr + 1
    Next Kolumna
    Next Wiersz
'wyswietlenie formularza
    VBA.UserForms.Add(NazwaFormularza).Show
End Sub

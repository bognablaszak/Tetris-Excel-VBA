Attribute VB_Name = "Module2_gra"
'Zmienne globalne
Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Plansza(1 To 20, 1 To 10) As Integer
Public AktualnyKlocek As Integer 'Typ klocka
Public Rotacja As Integer        'Stan obrotu
Public PozX As Integer           'Pozycja klocka w poziomie
Public PozY As Integer           'Pozycja klocka w pionie
Public CzyPauza As Boolean       'Pauzowanie gry
'Wszystkie mozliwe ksztalty
Public Ksztalty(1 To 7, 1 To 4, 1 To 4, 1 To 4) As Integer
Public GraDziala As Boolean
Public Wynik As Long
'Kolory dla poszczególnych numerów klocków (0 to puste tlo)
Public Kolory(0 To 7) As Long
'przycisk start
Sub WlaczMenuGlowne()
    a_menu.Show
End Sub
Sub BudowaGry()
    Dim Wiersz As Integer, Kolumna As Integer
    
    'reset parametrów
    CzyPauza = False
    Wynik = 0
    
    'Aktualizacja interfejsu (wyœwietlenie wartoœci pocz¹tkowych)
    b_gra.lb_wynik.Caption = "0"
    b_gra.lb_poziom.Caption = a_menu.poziom.Text
    
    'Czyszczenie planszy (wype³nienie zerami)
    For Wiersz = 1 To 20
        For Kolumna = 1 To 10
            Plansza(Wiersz, Kolumna) = 0
        Next Kolumna
    Next Wiersz
    
    'Kolory
    Kolory(0) = vbWhite                  'Puste pole
    Kolory(1) = RGB(0, 255, 255)         'Jasnoniebieski
    Kolory(2) = RGB(0, 0, 255)           'Ciemnoniebieski
    Kolory(3) = RGB(255, 165, 0)         'Pomarañczowy
    Kolory(4) = RGB(255, 255, 0)         '¯ó³ty
    Kolory(5) = RGB(0, 255, 0)           'Zielony
    Kolory(6) = RGB(128, 0, 128)         'Fioletowy
    Kolory(7) = RGB(255, 0, 0)           'Czerwony
    
    ' £adujemy geometrie do pamiêci
    BudowaKsztaltow
End Sub
'Definiowanie wygladu kazdego klocka
Private Sub BudowaKsztaltow()
    Erase Ksztalty 'Wyzerowanie pamiêci
    
    '0 to puste, 1 to fragment klocka, spacja oddziela wiersze siatki 4x4
    'I
    DefiniowanieKlocka 1, 1, "0000 1111 0000 0000"
    DefiniowanieKlocka 1, 2, "0100 0100 0100 0100"
    DefiniowanieKlocka 1, 3, "0000 1111 0000 0000"
    DefiniowanieKlocka 1, 4, "0100 0100 0100 0100"
    
    'J
    DefiniowanieKlocka 2, 1, "1000 1110 0000 0000"
    DefiniowanieKlocka 2, 2, "0110 0100 0100 0000"
    DefiniowanieKlocka 2, 3, "0000 1110 0010 0000"
    DefiniowanieKlocka 2, 4, "0100 0100 1100 0000"
    
    'L
    DefiniowanieKlocka 3, 1, "0010 1110 0000 0000"
    DefiniowanieKlocka 3, 2, "0100 0100 0110 0000"
    DefiniowanieKlocka 3, 3, "0000 1110 1000 0000"
    DefiniowanieKlocka 3, 4, "1100 0100 0100 0000"
    
    'O (Kwadrat, wiec kazda rotacja wygl¹da tak samo)
    Dim r As Integer
    For r = 1 To 4
        DefiniowanieKlocka 4, r, "0000 0110 0110 0000"
    Next r
    
    'S
    DefiniowanieKlocka 5, 1, "0110 1100 0000 0000"
    DefiniowanieKlocka 5, 2, "0100 0110 0010 0000"
    DefiniowanieKlocka 5, 3, "0110 1100 0000 0000"
    DefiniowanieKlocka 5, 4, "0100 0110 0010 0000"
    
    'T
    DefiniowanieKlocka 6, 1, "0100 1110 0000 0000"
    DefiniowanieKlocka 6, 2, "0100 0110 0100 0000"
    DefiniowanieKlocka 6, 3, "0000 1110 0100 0000"
    DefiniowanieKlocka 6, 4, "0100 1100 0100 0000"
    
    'Z
    DefiniowanieKlocka 7, 1, "1100 0110 0000 0000"
    DefiniowanieKlocka 7, 2, "0010 0110 0100 0000"
    DefiniowanieKlocka 7, 3, "1100 0110 0000 0000"
    DefiniowanieKlocka 7, 4, "0010 0110 0100 0000"
End Sub
'Zamiana znaków na tablice 4x4
Private Sub DefiniowanieKlocka(Typ As Integer, Rot As Integer, Wzor As String)
    Dim w As Integer, k As Integer
    'macierz 4x4
    For w = 1 To 4
        For k = 1 To 4
            ' Odczytujemy znak przeskakuj¹c spacje (wzor: (w-1)*5 + k)
            If Mid(Wzor, (w - 1) * 5 + k, 1) = "1" Then
                Ksztalty(Typ, Rot, w, k) = Typ
            End If
        Next k
    Next w
End Sub
'Aktualizuje wygl¹d na podstawie tablicy w pamiêci
Sub UpdateEkranu()
    Dim w As Integer, k As Integer
    Dim NazwaKontrolki As String
    Dim KolorTla As Long
    
    'Rysuje klocki z pamieci (te co spadly)
    For w = 1 To 20
        For k = 1 To 10
            'Generuje nazwe kontrolki
            NazwaKontrolki = "b_" & w & "_" & k
            'Pobieranie koloru z palety
            KolorTla = Kolory(Plansza(w, k))
            b_gra.Controls(NazwaKontrolki).BackColor = KolorTla
        Next k
    Next w
    
    'Rysowanie aktywnego klocka
    If AktualnyKlocek > 0 Then
        Dim w_kl As Integer, k_kl As Integer
        Dim RzeczywisteY As Integer, RzeczywisteX As Integer
        'Przeszukujemy nasza siatke 4x4 aktualnego klocka
        For w_kl = 1 To 4
            For k_kl = 1 To 4
                If Ksztalty(AktualnyKlocek, Rotacja, w_kl, k_kl) > 0 Then
                    RzeczywisteY = PozY + w_kl - 1
                    RzeczywisteX = PozX + k_kl - 1
                    'Rysujemy tylko, jeœli klocek mieœci siê w kadrze
                    If RzeczywisteY >= 1 And RzeczywisteY <= 20 And RzeczywisteX >= 1 And RzeczywisteX <= 10 Then
                        NazwaKontrolki = "b_" & RzeczywisteY & "_" & RzeczywisteX
                        b_gra.Controls(NazwaKontrolki).BackColor = Kolory(AktualnyKlocek)
                    End If
                End If
            Next k_kl
        Next w_kl
    End If
    
    'odswiezenie formularza na ekranie, zeby program nie zamarzl
    DoEvents
End Sub
'Sprawdza, czy przysz³y ruch klocka jest w ogóle mo¿liwy
Public Function CzyRuchDozwolony(TestX As Integer, TestY As Integer, TestRot As Integer) As Boolean
    Dim w As Integer, k As Integer
    Dim RzeczywisteY As Integer, RzeczywisteX As Integer
    
    For w = 1 To 4
        For k = 1 To 4
            'Sprawdzamy tylko te kratki w siatce 4x4, gdzie fizycznie znajduje siê fragment klocka
            If Ksztalty(AktualnyKlocek, TestRot, w, k) > 0 Then
                RzeczywisteY = TestY + w - 1
                RzeczywisteX = TestX + k - 1
                'Sprawdzenie, czy klocek wychodzi poza plansze
                If RzeczywisteX < 1 Or RzeczywisteX > 10 Or RzeczywisteY > 20 Then
                    CzyRuchDozwolony = False
                    Exit Function
                End If
                'Sprawdzenie czy uderza w ju¿ zamro¿one klocki na planszy
                If RzeczywisteY >= 1 Then
                    If Plansza(RzeczywisteY, RzeczywisteX) > 0 Then
                        CzyRuchDozwolony = False
                        Exit Function
                    End If
                End If
            End If
        Next k
    Next w
    'Ruch jest mozliwy
    CzyRuchDozwolony = True
End Function
'Wpisuje klocek do tablicy Plansza na sta³e i losuje nowy
Sub ZamrozKlocek()
    Dim w As Integer, k As Integer
    Dim RzeczywisteY As Integer, RzeczywisteX As Integer
    
    For w = 1 To 4
        For k = 1 To 4
            If Ksztalty(AktualnyKlocek, Rotacja, w, k) > 0 Then
                RzeczywisteY = PozY + w - 1
                RzeczywisteX = PozX + k - 1
                'Zapisujemy numer koloru klocka do g³ównej planszy na sta³e
                If RzeczywisteY >= 1 And RzeczywisteY <= 20 And RzeczywisteX >= 1 And RzeczywisteX <= 10 Then
                    Plansza(RzeczywisteY, RzeczywisteX) = AktualnyKlocek
                End If
            End If
        Next k
    Next w
    
    SprawdzLinie
    
    'Losujemy nowy klocek (liczba od 1 do 7)
    AktualnyKlocek = Int((7 * Rnd) + 1)
    Rotacja = 1
    PozX = 4
    PozY = 1
    
    'Sprawdzamy czy ten nowy wylosowany klocek ma wolne miejsce
    If CzyRuchDozwolony(PozX, PozY, Rotacja) = False Then
        GraDziala = False
        b_gra.Hide
        On Error Resume Next
        e_game_over.lb_twoj_wynik.Caption = Wynik
        On Error GoTo 0
        e_game_over.Show
    End If
End Sub
Sub UruchomGre()
    'Uruchamia losowoœæ w Excelu, ¿ebyœmy nie mieli zawsze tych samych klocków
    Randomize
    '£adujemy czyst¹ planszê i klocki
    BudowaGry
    'Ustawiamy losowy startowy klocek z góry
    AktualnyKlocek = Int((7 * Rnd) + 1)
    Rotacja = 1
    PozX = 4
    PozY = 1
    'Pokazujemy okno gry w trybie pozwalajacym na dzialanie kodu w tle
    b_gra.Show vbModeless
    GraDziala = True
    
    Dim CzasStart As Single
    Dim Predkosc As Single
    
    'Odczytujemy poziom z menu i ustalamy czas spadania
    Select Case a_menu.poziom.Text
        Case "£atwy"
            Predkosc = 0.8  ' Wolne spadanie
        Case "Œredni"
            Predkosc = 0.5  ' Standardowe tempo
        Case "Trudny"
            Predkosc = 0.15 ' Bardzo szybko - wyzwanie!
        Case Else
            Predkosc = 0.5  ' Zabezpieczenie na wypadek b³êdu
    End Select
    
    Do While GraDziala
        If Not CzyPauza Then 'gra bez pauzy
            'Opadanie klocka
            If CzyRuchDozwolony(PozX, PozY + 1, Rotacja) = True Then
                PozY = PozY + 1
            Else 'klocek zatrzymuje sie, jesli nie moze spasc nizej
                ZamrozKlocek
            End If
        End If
        
        'odœwie¿enie ekranu
        UpdateEkranu
        
        'stoper
        CzasStart = Timer
        Do While Timer < CzasStart + Predkosc
            DoEvents
            SprawdzKlawiature 'reaguje na klawiature w czasie spadku
        Loop
    Loop
End Sub
'Sprawdza i usuwa pe³ne linie, dodaj¹c punkty
Sub SprawdzLinie()
    Dim w As Integer, k As Integer, w2 As Integer
    Dim PelnaLinia As Boolean
    Dim UsunieteLinie As Integer
    UsunieteLinie = 0
    
    'Skanujemy od dolu do gory plansze
    For w = 20 To 1 Step -1
        PelnaLinia = True
        'Sprawdzamy czy w danym wierszu s¹ jakieœ puste luki
        For k = 1 To 10
            If Plansza(w, k) = 0 Then
                PelnaLinia = False 'dziura w wierszu
                Exit For
            End If
        Next k
        
        'Jeœli ca³a linia jest pe³na klocków
        If PelnaLinia = True Then
            UsunieteLinie = UsunieteLinie + 1
            'Zsuwamy wszystko co jest powy¿ej o jedno piêtro w dó³
            For w2 = w To 2 Step -1
                For k = 1 To 10
                    Plansza(w2, k) = Plansza(w2 - 1, k)
                Next k
            Next w2
            
            'Najwy¿szy wiersz na wszelki wypadek czyœcimy bo zjecha³ w dó³
            For k = 1 To 10
                Plansza(1, k) = 0
            Next k
            
            'Wiersze zjechaly wiec sprawdzamy jeszcze raz
            w = w + 1
        End If
    Next w
    
    'Jeœli usunêliœmy jakieœ linie, aktualizujemy wynik
    If UsunieteLinie > 0 Then
        'Punktacja 100pkt za ka¿d¹ liniê
        Wynik = Wynik + (UsunieteLinie * 100)
        'wyswietlamy wynik
        On Error Resume Next
        b_gra.lb_wynik.Caption = Wynik
        On Error GoTo 0
    End If
End Sub
'Natychmiastowe przeniesienie klocka na dó³ planszy - Hard Drop
Sub SzybkiSpadek()
    If Not GraDziala Or CzyPauza Then Exit Sub
    'Przesuwanie klocka w dó³ tak d³ugo jak to mo¿liwe
    Do While CzyRuchDozwolony(PozX, PozY + 1, Rotacja)
        PozY = PozY + 1
    Loop
    'Odœwie¿enie ekranu i blokowanie klocka po opadniêciu na sam dó³
    UpdateEkranu
    ZamrozKlocek
End Sub
'W³¹czanie/wy³¹czanie pauzy
Sub PrzelaczPauze()
    If Not GraDziala Then Exit Sub
    'Odwrócenie znaczenia pauza
    CzyPauza = Not CzyPauza
    'Aktualizacja napisu UF
    If CzyPauza Then
        b_gra.Caption = "TETRIS - PAUZA (Kliknij pauzê, aby wznowiæ)"
    Else
        b_gra.Caption = "TETRIS"
    End If
End Sub
'Tworzenie opóŸnienia w dzia³aniu kodu
Private Sub PauzaDlaKlawisza(Sekundy As Single)
    Dim t As Single: t = Timer
    Do While Timer < t + Sekundy
        DoEvents
    Loop
End Sub

'Sterowanie klockiem przez klawisze i przyciski
Sub WykonajRuch(dx As Integer, dy As Integer, dRot As Integer)
    Dim NoweX As Integer, NoweY As Integer, NowaRotacja As Integer
    'Obliczenie nowych wspó³rzêdnych
    NoweX = PozX + dx
    NoweY = PozY + dy
    NowaRotacja = Rotacja + dRot
    
    'Obs³uga rotacji
    If NowaRotacja > 4 Then NowaRotacja = 1
    If NowaRotacja < 1 Then NowaRotacja = 4 ' Na wypadek rotacji w drug¹ stronê
    
    'Sprawdzenie czy nowe miejsce jest dostêne
    If CzyRuchDozwolony(NoweX, NoweY, NowaRotacja) Then
        PozX = NoweX
        PozY = NoweY
        Rotacja = NowaRotacja
        UpdateEkranu
    End If
End Sub
'Sprawdza stan klawiszy na klawiaturze
Sub SprawdzKlawiature()
    Dim Kliknieto As Boolean
    'Sprawdzenie klawisza pauzy (P)
    If GetAsyncKeyState(vbKeyP) < 0 Then 'zwraca wartoœæ ujemn¹ przy wciœniêtym klawiszu
        PrzelaczPauze
        PauzaDlaKlawisza 0.3
        Exit Sub
    End If
    
    'Je¿eli pauza to koniec procedury
    If CzyPauza = True Then Exit Sub
    Kliknieto = False
    
    'Strza³ka w lewo/prawo - ruch lewo/prawo
    If GetAsyncKeyState(vbKeyLeft) < 0 Then
        WykonajRuch -1, 0, 0
        Kliknieto = True
    ElseIf GetAsyncKeyState(vbKeyRight) < 0 Then
        WykonajRuch 1, 0, 0
        Kliknieto = True
        
    'Strza³ka w górê - obrót
    ElseIf GetAsyncKeyState(vbKeyUp) < 0 Then
        WykonajRuch 0, 0, 1
        Kliknieto = True
        
    'Strza³ka w Dó³ - przyspiesza opadanie (soft drop)
    ElseIf GetAsyncKeyState(vbKeyDown) < 0 Then
        WykonajRuch 0, 1, 0
        Kliknieto = True
        PauzaDlaKlawisza 0.05
        Exit Sub
        
    'Spacja - natychmiastowe l¹dowanie (hard drop)
    ElseIf GetAsyncKeyState(vbKeySpace) < 0 Then
        SzybkiSpadek
        PauzaDlaKlawisza 0.2 'du¿a pauza przed nowym klockiem
        Exit Sub
    End If
    
    'Krótka pauza po wykonaniu ruchu lewo/prawo/obrót
    If Kliknieto Then
        PauzaDlaKlawisza 0.1
    End If
End Sub

Sub ZatrzymajGre()
    GraDziala = False
End Sub

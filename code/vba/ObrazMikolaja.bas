Attribute VB_Name = "ObrazMikolaja"
Option Explicit

'--------------------------------------------------------------------
' CEL MODU£U
' Na podstawie procentu w komórce (0–1, np. 0,25 = 25%) wybieramy:
'   • jednego Miko³aja (nazwa zaczyna siê na "d", np. d3)
'   • jeden pasek postêpu (nazwa zaczyna siê na "p", np. p3)
' i pokazujemy tylko te dwa obrazki, a resztê ukrywamy.
'--------------------------------------------------------------------

'=== KONFIGURACJA ===
Private Const TBL_SHEET As String = "tajne zapiski elfów"   ' Arkusz z tabel¹ steruj¹c¹ (procent + nazwy obrazków)
Private Const CAL_SHEET As String = "kalendarz"             ' Arkusz z kalendarzem (aktywny)

Private Const HDR_PERCENT As String = "ProcentDocelowy"     ' Nag³ówek kolumny z wartoœci¹ procentow¹ (0–1)
Private Const HDR_STAGE As String = "EtapMikolaja"          ' Nag³ówek kolumny z nazw¹ obrazka Miko³aja (np. d5)
Private Const HDR_BAR_STAGE As String = "EtapPaska"         ' Nag³ówek kolumny z nazw¹ obrazka paska (np. p5)

Private Const PERCENT_CELL As String = "D27"                ' Komórka z bie¿¹cym procentem (0–1)
Private Const EPS As Double = 0.000001                      ' EPS to ma³a „tolerancja” przy porównaniu liczb: wartoœci
                                                            ' wygl¹daj¹ce tak samo w komórkach mog¹ w pamiêci ró¿niæ siê
                                                            ' o mikrou³amki (np. 0,30000000000000004), wiêc jeœli ró¿nica
                                                            ' miêdzy dwiema liczbami jest mniejsza ni¿ EPS, traktujemy je
                                                            ' jako równe.

'--------------------------------------------------------------------
' G³ówna procedura:
' 1) Pobiera bie¿¹cy procent z PERCENT_CELL.
' 2) W tabeli znajduje wiersz z ProcentDocelowy najbli¿szym bie¿¹cej wartoœci (z tolerancj¹ EPS).
' 3) Z tego wiersza pobiera nazwy: Miko³aja (d*) i paska (p*).
' 4) Na arkuszu z obrazkami ukrywa wszystkie d*/p*, a pokazuje tylko wybrane dwa.
'--------------------------------------------------------------------
Public Sub UpdateSantaByPercent()
    Dim wsTbl As Worksheet, wsCal As Worksheet
    Set wsTbl = ThisWorkbook.Worksheets(TBL_SHEET)  ' zmienna dla arkusza z tabel¹ steruj¹c¹
    Set wsCal = ThisWorkbook.Worksheets(CAL_SHEET)  ' zmienna dla arkusza z obrazkami

    ' 1) Bie¿¹cy procent (0–1)
        Dim targetPct As Double
        targetPct = wsTbl.Range(PERCENT_CELL).Value2

    ' 2) Numery kolumn wg nag³ówków w wierszu 1
        Dim cPct As Long, cSanta As Long, cBar As Long
        cPct = FindHeader(wsTbl, HDR_PERCENT)
        cSanta = FindHeader(wsTbl, HDR_STAGE)
        cBar = FindHeader(wsTbl, HDR_BAR_STAGE)
        If cPct = 0 Or cSanta = 0 Or cBar = 0 Then Exit Sub

    ' 3) ZnajdŸ wiersz w tabeli, którego ProcentDocelowy jest najbli¿ej bie¿¹cej wartoœci (targetPct)
        Dim lastRow As Long            ' numer ostatniego wiersza z danymi w kolumnie ProcentDocelowy
        Dim r As Long                  ' licznik pêtli (po którym wierszu w³aœnie idziemy)
        Dim bestRow As Long            ' numer „najlepszego” (najbli¿szego) wiersza znalezionego do tej pory
        Dim bestDiff As Double         ' najmniejsza ró¿nica znaleziona do tej pory
        Dim d As Double                ' ró¿nica w bie¿¹cym wierszu
        Dim rowPct As Double           ' wartoœæ ProcentDocelowy w bie¿¹cym wierszu
    
        ' Ustal, gdzie koñcz¹ siê dane w kolumnie procentów (¿eby nie jechaæ po pustych wierszach)
        lastRow = wsTbl.Cells(wsTbl.Rows.Count, cPct).End(xlUp).Row
    
        ' Na start: „nie mamy jeszcze faworyta”, ustaw bardzo du¿¹ ró¿nicê i brak wiersza
        bestDiff = 1000000000#
        bestRow = 0
    
        ' PrzejdŸ przez wszystkie wiersze danych (zak³adamy, ¿e nag³ówki s¹ w wierszu 1)
        For r = 2 To lastRow
        rowPct = wsTbl.Cells(r, cPct).Value2   ' odczytaj procent z tego wiersza (0–1)
        d = Abs(rowPct - targetPct)            ' policz, jak bardzo ró¿ni siê od bie¿¹cej wartoœci
    
        ' Jeœli ta ró¿nica jest mniejsza ni¿ dotychczasowa to zapamiêtaj ten wiersz jako „najbli¿szy”
        If d < bestDiff Then
            bestDiff = d
            bestRow = r
        End If
    
        ' Jeœli trafiliœmy praktycznie w punkt, dalej ju¿ nie szukamy
        If d <= EPS Then Exit For
        Next r
    
        ' Gdyby z jakiegoœ powodu nic nie znaleziono (pusta kolumna itp.), koñczymy bez b³êdu
        If bestRow = 0 Then Exit Sub

    ' 4) Z wybranego wiersza tabeli pobierz NAZWY DWÓCH OBRAZKÓW, które maj¹ zostaæ pokazane na arkuszu:
    '    • santaName – Miko³aj (nazwy zaczynaj¹ siê na "d", np. d7)
    '    • barName   – pasek postêpu (nazwy zaczynaj¹ siê na "p", np. p7)
    '    To musz¹ byæ dok³adnie takie same nazwy, jak nazwy kszta³tów na arkuszu „kalendarz”.

        Dim santaName As String, barName As String
        santaName = Trim$(CStr(wsTbl.Cells(bestRow, cSanta).Value))   ' np. "d7"
        barName = Trim$(CStr(wsTbl.Cells(bestRow, cBar).Value))       ' np. "p7"

    ' Gdyby któraœ nazwa by³a pusta, przerwij – nie ma co pokazywaæ
        If santaName = "" Or barName = "" Then
            'MsgBox "Brak nazwy obrazka w tabeli dla wiersza " & bestRow, vbExclamation
            Exit Sub
        End If

    ' 5) Prze³¹cz widocznoœæ obrazków na arkuszu z obrazkami
        Dim shp As Shape, nm As String, pf As String
        Dim prevUpd As Boolean

        ' Zapamiêtaj aktualne ustawienie odœwie¿ania ekranu i wy³¹cz je na czas zmian,
        ' ¿eby unikn¹æ migotania podczas hurtowego ukrywania/pokazywania kszta³tów.
        prevUpd = Application.ScreenUpdating
        Application.ScreenUpdating = False
        
        On Error GoTo CleanExit   ' Jeœli wydarzy siê b³¹d w pêtli, przeskocz do CleanExit (przywróci odœwie¿anie)
        
        ' PrzejdŸ po KA¯DYM obrazku na arkuszu „kalendarz”
        For Each shp In wsCal.Shapes
            nm = shp.Name                      ' pe³na nazwa kszta³tu, np. "d7", "p7", "tlo"
            pf = LCase$(Left$(nm, 1))          ' pierwsza litera nazwy („d” albo „p” decyduje o grupie)
            
            Select Case pf
                Case "d"
                    ' Grupa „Miko³aje”: poka¿ TYLKO tego, którego nazwa = santaName; pozosta³e ukryj
                    shp.Visible = (nm = santaName)
                Case "p"
                    ' Grupa „paski”: poka¿ TYLKO ten, którego nazwa = barName; pozosta³e ukryj
                    shp.Visible = (nm = barName)
                ' Inne kszta³ty (np. t³o, napisy), nie ruszamy ich widocznoœci
            End Select
        Next shp
        
CleanExit:
        ' Niezale¿nie od tego, czy by³ b³¹d, czy nie, przywróæ poprzednie ustawienie odœwie¿ania ekranu.
        Application.ScreenUpdating = prevUpd
        End Sub

'--------------------------------------------------------------------
' Funkcja pomocnicza:
' Zwraca numer kolumny, w której w wierszu 1 znajduje siê dok³adnie dany nag³ówek;
' w przeciwnym razie zwraca 0.
'--------------------------------------------------------------------
Private Function FindHeader(ws As Worksheet, headerText As String) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        If Len(c.Value2) = 0 Then Exit For
        If StrComp(CStr(c.Value2), headerText, vbTextCompare) = 0 Then
            FindHeader = c.Column
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

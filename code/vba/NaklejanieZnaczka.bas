Attribute VB_Name = "NaklejanieZnaczka"
Option Explicit

'------------------------------------------------------------
' CEL:
' Po kliknięciu w miejsce na znaczek pocztowy:
'  1) znajdź w tabeli wiersz z nazwą tego obrazka,
'  2) wpisz w kolumnie potwierdzeń "DONE",
'  3) ukryj kliknięty obrazek (odsłoni znaczek).
'------------------------------------------------------------

'== KONFIGURACJA ==
Const CONTROL_TABLE_SHEET_NAME As String = "tajne zapiski elfów"   ' Arkusz z tabelą sterującą
Const HEADER_PICTURE_NAME As String = "NazwaObrazka"               ' Nagłówek kolumny z nazwami obrazków
Const HEADER_CONFIRM_CELL As String = "KomorkaPotwierdzenia"       ' Nagłówek kolumny do wpisania "DONE"

'== GŁÓWNE MAKRO: klik -> "DONE" w tabeli -> ukrycie obrazka ==
Sub HideAndMarkDone()
    Dim clickedShapeName As String                 ' nazwa klikniętego obrazka/kształtu
    Dim controlTableSheet As Worksheet             ' arkusz z tabelą sterującą
    Dim sheetWhereClickOccurred As Worksheet       ' arkusz z kalendarzem (aktywny)
    Dim pictureNameColIndex As Long                ' numer kolumny z nagłówkiem "NazwaObrazka"
    Dim confirmColIndex As Long                    ' numer kolumny z nagłówkiem "KomorkaPotwierdzenia"
    Dim foundPictureRowCell As Range               ' komórka w kolumnie z nazwami, gdzie znaleziono nasz obrazek

    ' 1) Nazwa klikniętego obiektu (Excel przekazuje ją automatycznie)
    clickedShapeName = Application.Caller

    ' 2) Ustal referencje do arkuszy
    Set controlTableSheet = ThisWorkbook.Worksheets(CONTROL_TABLE_SHEET_NAME) ' wskazanie arkusza z tabelą sterującą
    Set sheetWhereClickOccurred = ActiveSheet                                 ' wskazanie aktywnego arkusza

    ' 3) Pobierz numery kolumn wg nagłówków w wierszu 1
    pictureNameColIndex = FindHeaderColumn(controlTableSheet, HEADER_PICTURE_NAME)      ' wskazanie numeru kolumny z nagłówkiem "NazwaObrazka"
    confirmColIndex = FindHeaderColumn(controlTableSheet, HEADER_CONFIRM_CELL)          ' wskazanie numeru kolumny z nagłówkiem "KomorkaPotwierdzenia"
    If pictureNameColIndex = 0 Or confirmColIndex = 0 Then Exit Sub                     ' brak wymaganych nagłówków

    ' 4) Znajdź wiersz odpowiadający nazwie klikniętego obrazka
    Set foundPictureRowCell = controlTableSheet.Columns(pictureNameColIndex).Find( _
                                What:=clickedShapeName, LookIn:=xlValues, _
                                LookAt:=xlWhole, MatchCase:=False)
    If foundPictureRowCell Is Nothing Then Exit Sub                                     ' nie znaleziono – zakończ

    ' 5) Zaznacz wykonanie: wpisz "DONE" w kolumnie potwierdzeń tego samego wiersza
    controlTableSheet.Cells(foundPictureRowCell.Row, confirmColIndex).Value = "DONE"

    ' 6) Ukryj kliknięty obrazek
    sheetWhereClickOccurred.Shapes(clickedShapeName).Visible = msoFalse
End Sub

'== POMOCNICZA: zwraca numer kolumny z podanym nagłówkiem w wierszu 1 (albo 0, gdy brak) ==
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim headerCell As Range
    Set headerCell = ws.Rows(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerCell Is Nothing Then
        FindHeaderColumn = headerCell.Column
    Else
        FindHeaderColumn = 0
    End If
End Function


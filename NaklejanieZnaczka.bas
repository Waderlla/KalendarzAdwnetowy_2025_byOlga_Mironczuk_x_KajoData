Attribute VB_Name = "NaklejanieZnaczka"
Option Explicit

'------------------------------------------------------------
' CEL:
' Po klikniêciu w miejsce na znaczek pocztowy:
'  1) znajdŸ w tabeli wiersz z nazw¹ tego obrazka,
'  2) wpisz w kolumnie potwierdzeñ "DONE",
'  3) ukryj klikniêty obrazek (ods³oni znaczek).
'------------------------------------------------------------

'== KONFIGURACJA ==
Const CONTROL_TABLE_SHEET_NAME As String = "tajne zapiski elfów"   ' Arkusz z tabel¹ steruj¹c¹
Const HEADER_PICTURE_NAME As String = "NazwaObrazka"               ' Nag³ówek kolumny z nazwami obrazków
Const HEADER_CONFIRM_CELL As String = "KomorkaPotwierdzenia"       ' Nag³ówek kolumny do wpisania "DONE"

'== G£ÓWNE MAKRO: klik -> "DONE" w tabeli -> ukrycie obrazka ==
Sub HideAndMarkDone()
    Dim clickedShapeName As String                 ' nazwa klikniêtego obrazka/kszta³tu
    Dim controlTableSheet As Worksheet             ' arkusz z tabel¹ steruj¹c¹
    Dim sheetWhereClickOccurred As Worksheet       ' arkusz z kalendarzem (aktywny)
    Dim pictureNameColIndex As Long                ' numer kolumny z nag³ówkiem "NazwaObrazka"
    Dim confirmColIndex As Long                    ' numer kolumny z nag³ówkiem "KomorkaPotwierdzenia"
    Dim foundPictureRowCell As Range               ' komórka w kolumnie z nazwami, gdzie znaleziono nasz obrazek

    ' 1) Nazwa klikniêtego obiektu (Excel przekazuje j¹ automatycznie)
    clickedShapeName = Application.Caller

    ' 2) Ustal referencje do arkuszy
    Set controlTableSheet = ThisWorkbook.Worksheets(CONTROL_TABLE_SHEET_NAME) ' wskazanie arkusza z tabel¹ steruj¹c¹
    Set sheetWhereClickOccurred = ActiveSheet                                 ' wskazanie aktywnego arkusza

    ' 3) Pobierz numery kolumn wg nag³ówków w wierszu 1
    pictureNameColIndex = FindHeaderColumn(controlTableSheet, HEADER_PICTURE_NAME)      ' wskazanie numeru kolumny z nag³ówkiem "NazwaObrazka"
    confirmColIndex = FindHeaderColumn(controlTableSheet, HEADER_CONFIRM_CELL)          ' wskazanie numeru kolumny z nag³ówkiem "KomorkaPotwierdzenia"
    If pictureNameColIndex = 0 Or confirmColIndex = 0 Then Exit Sub                     ' brak wymaganych nag³ówków

    ' 4) ZnajdŸ wiersz odpowiadaj¹cy nazwie klikniêtego obrazka
    Set foundPictureRowCell = controlTableSheet.Columns(pictureNameColIndex).Find( _
                                What:=clickedShapeName, LookIn:=xlValues, _
                                LookAt:=xlWhole, MatchCase:=False)
    If foundPictureRowCell Is Nothing Then Exit Sub                                     ' nie znaleziono – zakoñcz

    ' 5) Zaznacz wykonanie: wpisz "DONE" w kolumnie potwierdzeñ tego samego wiersza
    controlTableSheet.Cells(foundPictureRowCell.Row, confirmColIndex).Value = "DONE"

    ' 6) Ukryj klikniêty obrazek
    sheetWhereClickOccurred.Shapes(clickedShapeName).Visible = msoFalse
End Sub

'== POMOCNICZA: zwraca numer kolumny z podanym nag³ówkiem w wierszu 1 (albo 0, gdy brak) ==
Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim headerCell As Range
    Set headerCell = ws.Rows(1).Find(What:=headerText, LookIn:=xlValues, LookAt:=xlWhole)
    If Not headerCell Is Nothing Then
        FindHeaderColumn = headerCell.Column
    Else
        FindHeaderColumn = 0
    End If
End Function


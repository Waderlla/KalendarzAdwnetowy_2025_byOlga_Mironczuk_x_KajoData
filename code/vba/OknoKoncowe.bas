Attribute VB_Name = "OknoKoncowe"
Option Explicit

'--------------------------------------------------------------------
' CEL MODU£U (f*)
' Na podstawie procentu w komórce (0–1, np. 0,25) wybieramy jeden obrazek
' z kolumny "Final" (nazwa zaczyna siê na "f", np. f3) i pokazujemy tylko jego.
' Pozosta³e f* ukrywamy. Innych kszta³tów nie dotykamy.
'--------------------------------------------------------------------

'=== KONFIGURACJA ===
Private Const TBL_SHEET As String = "tajne zapiski elfów"   ' Arkusz z tabel¹ steruj¹c¹
Private Const CAL_SHEET As String = "kalendarz"             ' Arkusz z kalendarzem

Private Const HDR_PERCENT As String = "ProcentDocelowy"     ' Kolumna z wartoœci¹ procentow¹ (0–1)
Private Const HDR_FINAL As String = "Final"                 ' Kolumna z nazw¹ obrazka f* (np. f5)

Private Const PERCENT_CELL As String = "D27"                ' Komórka z bie¿¹cym procentem (0–1)
Private Const EPS As Double = 0.000001                      ' Tolerancja porównania liczb
'===========================================================

Public Sub UpdateFinalByPercent()
    Dim wsTbl As Worksheet, wsCal As Worksheet
    Set wsTbl = ThisWorkbook.Worksheets(TBL_SHEET)
    Set wsCal = ThisWorkbook.Worksheets(CAL_SHEET)

    ' 0) Bramka daty – poka¿ f* dopiero OD dnia z D28.
    '    Jeœli D28 nie jest prawdziw¹ dat¹ (tekst albo puste) – bramkê ignorujemy.
    Dim startVal As Variant
    startVal = wsTbl.Range("D28").Value2

    If IsDate(startVal) Then
        ' porównuj TYLKO czêœæ daty (bez czasu)
        If Date < DateValue(CDate(startVal)) Then
            Dim shp0 As Shape
            Dim prevUpd0 As Boolean
            prevUpd0 = Application.ScreenUpdating
            Application.ScreenUpdating = False
            For Each shp0 In wsCal.Shapes
                If LCase$(Left$(shp0.Name, 1)) = "f" Then shp0.Visible = False
            Next shp0
            Application.ScreenUpdating = prevUpd0
            Exit Sub
        End If
    End If

    ' 1) Bie¿¹cy procent
    Dim targetPct As Double
    targetPct = wsTbl.Range(PERCENT_CELL).Value2

    ' 2) Numery kolumn po nag³ówkach
    Dim cPct As Long, cFin As Long
    cPct = FindHeader(wsTbl, HDR_PERCENT)
    cFin = FindHeader(wsTbl, HDR_FINAL)
    If cPct = 0 Or cFin = 0 Then Exit Sub

    ' 3) Szukamy wiersza, którego ProcentDocelowy jest najbli¿ej targetPct
    Dim lastRow As Long, r As Long
    Dim bestRow As Long, bestDiff As Double
    Dim d As Double, rowPct As Double

    lastRow = wsTbl.Cells(wsTbl.Rows.Count, cPct).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    bestDiff = 1E+30
    bestRow = 0

    For r = 2 To lastRow
        rowPct = wsTbl.Cells(r, cPct).Value2
        d = Abs(rowPct - targetPct)
        If d < bestDiff Then
            bestDiff = d
            bestRow = r
        End If
        If d <= EPS Then Exit For ' trafienie "prawie w punkt"
    Next r

    If bestRow = 0 Then Exit Sub

    ' 4) Nazwa obrazka f* z kolumny "Final"
    Dim fName As String
    fName = Trim$(CStr(wsTbl.Cells(bestRow, cFin).Value))
    If Len(fName) = 0 Then Exit Sub

    ' 5) Prze³¹cz widocznoœæ tylko dla grupy f*
    Dim shp As Shape, nm As String
    Dim prevUpd As Boolean
    prevUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo CleanExit

    For Each shp In wsCal.Shapes
        nm = shp.Name
        If LCase$(Left$(nm, 1)) = "f" Then
            shp.Visible = (nm = fName)
        End If
    Next shp

CleanExit:
    Application.ScreenUpdating = prevUpd
End Sub

'--------------------------------------------------------------------
' Funkcja pomocnicza: zwraca numer kolumny, której nag³ówek (wiersz 1)
' jest równy podanemu tekstowi (bez rozró¿niania wielkoœci liter).
' Gdy brak — zwraca 0.
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

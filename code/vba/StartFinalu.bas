Attribute VB_Name = "StartFinalu"
Option Explicit

'============== KONFIG STARTERA ==============
Private Const TBL_SHEET      As String = "tajne zapiski elfów"
Private Const CAL_SHEET      As String = "kalendarz"
Private Const DEADLINE_CELL  As String = "D28"                  ' data startu pokazywania f*
'=============================================

' GŁÓWNE WEJŚCIE: URUCHAMIAJ TĘ PROCEDURĘ zamiast UpdateFinalByPercent
Public Sub RunFinalWithGate()
    Dim wsTbl As Worksheet, wsCal As Worksheet
    Set wsTbl = ThisWorkbook.Worksheets(TBL_SHEET)
    Set wsCal = ThisWorkbook.Worksheets(CAL_SHEET)

    ' 1) Jeśli dziń < data z D28 › ukryj wszystkie f* i zakończ
    If GateBlocks(wsTbl, wsCal) Then Exit Sub

    ' 2) W przeciwnym wypadku › odpal moduł
    Call UpdateFinalByPercent
End Sub

' --- BRAMKA DATY: True = ma blokowaæ (ukrywa f* i zatrzymuje) ---
Private Function GateBlocks(wsTbl As Worksheet, wsCal As Worksheet) As Boolean
    Dim startDt As Date, ok As Boolean
    ok = TryGetDate(wsTbl.Range(DEADLINE_CELL).Value2, startDt)
    If ok Then
        ' Porównujemy tylko czêœæ daty (bez czasu)
        If CLng(Date) < CLng(startDt) Then
            HideAllF wsCal
            GateBlocks = True
            Exit Function
        End If
    End If
    GateBlocks = False
End Function

' --- Ukrywa wszystkie kształty, których nazwa zaczyna się na "f" (także w grupach) ---
Private Sub HideAllF(wsCal As Worksheet)
    Dim shp As Shape
    Dim prevUpd As Boolean
    prevUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo ExitHere

    For Each shp In wsCal.Shapes
        If LCase$(Left$(shp.Name, 1)) = "f" Then
            shp.Visible = False
        ElseIf shp.Type = msoGroup Then
            Dim g As Shape
            For Each g In shp.GroupItems
                If LCase$(Left$(g.Name, 1)) = "f" Then g.Visible = False
            Next g
        End If
    Next shp

ExitHere:
    Application.ScreenUpdating = prevUpd
End Sub

' --- Bezpieczne pobranie samej daty (obsługuje liczby i teksty) ---
Private Function TryGetDate(ByVal rawVal As Variant, ByRef outDate As Date) As Boolean
    On Error GoTo Fail
    If IsDate(rawVal) Or IsNumeric(rawVal) Or Len(Trim$(CStr(rawVal))) > 0 Then
        Dim d As Date: d = CDate(rawVal)
        outDate = DateSerial(Year(d), Month(d), Day(d)) ' sam dzieñ
        TryGetDate = True
        Exit Function
    End If
Fail:
    TryGetDate = False
End Function


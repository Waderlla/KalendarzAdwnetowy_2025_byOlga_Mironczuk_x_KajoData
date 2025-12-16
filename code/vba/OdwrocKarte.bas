Attribute VB_Name = "OdwrocKarte"
Option Explicit

'=== KONFIG ===
Private Const TBL_SHEET As String = "tajne zapiski elfów"   ' tabela steruj¹ca
Private Const CAL_SHEET As String = "kalendarz"             ' arkusz z kszta³tami

Private Const HDR_START As String = "PoczatkowaData"        ' data startu (dzieñ odkrycia)
Private Const HDR_BACK  As String = "TylKarty"              ' nazwa obrazka tk* (np. tk7)

' Ukryj tk*, dla których PoczatkowaData <= dzisiaj. Pozosta³e poka¿.
Public Sub OdwrocKarte()
    Dim wsT As Worksheet, wsC As Worksheet
    Dim cStart As Long, cBack As Long, lastRow As Long, r As Long
    Dim today As Date: today = Date

    Set wsT = ThisWorkbook.Worksheets(TBL_SHEET)
    Set wsC = ThisWorkbook.Worksheets(CAL_SHEET)

    cStart = FindHeaderSafe(wsT, HDR_START)
    cBack = FindHeaderSafe(wsT, HDR_BACK)
    If cStart = 0 Or cBack = 0 Then Exit Sub

    lastRow = wsT.Cells(wsT.Rows.Count, cStart).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim toHide As Collection, toShow As Collection
    Set toHide = New Collection
    Set toShow = New Collection

    Dim nm As String, v As Variant, d As Date
    For r = 2 To lastRow
        nm = Trim$(CStr(wsT.Cells(r, cBack).Value))
        If Len(nm) > 0 Then
            v = wsT.Cells(r, cStart).Value
            If IsDate(v) Then
                d = DateValue(CDate(v))
                If d <= today Then
                    toHide.Add nm     ' dzieñ nadszed³ -> ukryj ty³
                Else
                    toShow.Add nm     ' jeszcze przed startem -> poka¿ ty³
                End If
            End If
        End If
    Next r

    ' Prze³¹cz widocznoœæ tylko dla tk*
    Dim shp As Shape
    Dim prevUpd As Boolean: prevUpd = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For Each shp In wsC.Shapes
        If LCase$(Left$(shp.Name, 2)) = "tk" Then
            If InColl(toHide, shp.Name) Then
                shp.Visible = msoFalse
            ElseIf InColl(toShow, shp.Name) Then
                shp.Visible = msoTrue
            End If
        End If
    Next shp

    Application.ScreenUpdating = prevUpd
End Sub

'== helpery ==
Private Function FindHeaderSafe(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim lastCol As Long, c As Range, txt As String, want As String
    want = LCase$(Trim$(headerText))
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Function
    For Each c In ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Cells
        txt = Replace$(LCase$(Trim$(CStr(c.Value2))), ChrW(160), " ")
        If txt = Replace$(want, ChrW(160), " ") Then
            FindHeaderSafe = c.Column
            Exit Function
        End If
    Next c
End Function

Private Function InColl(col As Collection, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = key Then InColl = True: Exit Function
    Next i
End Function

Attribute VB_Name = "ZnakX"
Option Explicit

'================= KONFIG =================
Private Const TBL_SHEET As String = "tajne zapiski elfów"   ' arkusz z tabelą
Private Const CAL_SHEET As String = "kalendarz"             ' arkusz z obrazkami

' Nag³ówki kolumn w tabeli (wiersz 1)
Private Const HDR_END  As String = "KoncowaData"            ' data końcowa dla wiersza
Private Const HDR_X    As String = "X"                      ' nazwa obrazka x* (np. x7)
Private Const HDR_CONF As String = "KomorkaPotwierdzenia"   ' komórka potwierdzenia (DONE)

'================================================
' POKAŻ x*, dla których:
'   KoncowaData < dzisiaj  ORAZ  KomorkaPotwierdzenia <> "DONE"
' Resztę ukryj.
'================================================
Public Sub ShowX_UpToToday_KeepVisible()
    Dim wsT As Worksheet, wsC As Worksheet
    Dim cEnd As Long, cX As Long, cConf As Long
    Dim lastRow As Long, r As Long
    Dim today As Date: today = Date

    Set wsT = ThisWorkbook.Worksheets(TBL_SHEET)
    Set wsC = ThisWorkbook.Worksheets(CAL_SHEET)

    cEnd = FindHeader(wsT, HDR_END)
    cX = FindHeader(wsT, HDR_X)
    cConf = FindHeader(wsT, HDR_CONF)        ' może być 0 jeśli kolumny nie ma
    If cEnd = 0 Or cX = 0 Then Exit Sub

    lastRow = wsT.Cells(wsT.Rows.Count, cEnd).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim toShow As Collection, toHide As Collection
    Set toShow = New Collection
    Set toHide = New Collection

    Dim nm As String, dt As Variant, dtOnly As Date
    Dim conf As String

    For r = 2 To lastRow
        nm = Trim$(CStr(wsT.Cells(r, cX).Value))
        If Len(nm) > 0 Then

            ' 1) Jeœli DONE -> zawsze ukryj
            If cConf > 0 Then
                conf = UCase$(Trim$(CStr(wsT.Cells(r, cConf).Value)))
                If conf = "DONE" Then
                    toHide.Add nm
                    GoTo NextRow
                End If
            End If

            ' 2) Bramka daty
            dt = wsT.Cells(r, cEnd).Value
            If IsDate(dt) Then
                dtOnly = DateValue(CDate(dt))
                If dtOnly < today Then
                    toShow.Add nm
                Else
                    toHide.Add nm
                End If
            End If
        End If
NextRow:
    Next r

    ' Przełącz tylko kształty x* na arkuszu kalendarza
    Dim shp As Shape
    Application.ScreenUpdating = False
    For Each shp In wsC.Shapes
        If LCase$(Left$(shp.Name, 1)) = "x" Then
            If InColl(toShow, shp.Name) Then
                shp.Visible = msoTrue
            ElseIf InColl(toHide, shp.Name) Then
                shp.Visible = msoFalse
            End If
        End If
    Next shp
    Application.ScreenUpdating = True
End Sub

'== pomocnicza: numer kolumny po nagłówku (0, gdy brak)
Private Function FindHeader(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim c As Range, txt As String, want As String
    want = LCase$(Trim$(headerText))
    For Each c In ws.Rows(1).Cells
        If Len(c.Value2) = 0 Then Exit For
        txt = Replace$(LCase$(Trim$(CStr(c.Value2))), ChrW(160), " ")
        If txt = Replace$(want, ChrW(160), " ") Then
            FindHeader = c.Column
            Exit Function
        End If
    Next c
    FindHeader = 0
End Function

'== pomocnicza: czy element jest w kolekcji (dokładne dopasowanie)
Private Function InColl(col As Collection, ByVal key As String) As Boolean
    Dim i As Long
    For i = 1 To col.Count
        If col(i) = key Then InColl = True: Exit Function
    Next i
End Function



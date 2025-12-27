VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit                           ' Wymusza jawne deklarowanie wszystkich zmiennych.
                                          ' Dziêki temu ka¿da zmienna musi byæ wczeœniej zadeklarowana instrukcj¹ Dim.
                                          ' Chroni to przed b³êdami literówek.

Private Sub Workbook_Open()               ' Procedura zdarzeniowa skoroszytu (ThisWorkbook).
                                          ' Uruchamia siê AUTOMATYCZNIE w momencie otwarcia pliku Excel (po w³¹czeniu makr).
                                          ' Nie wymaga ¿adnego rêcznego uruchamiania. Excel sam j¹ wywo³a po otwarciu pliku.

    MasterRefresh "Open"                  ' Wywo³anie naszego g³ównego makra steruj¹cego (MasterRefresh).
                                          ' Przekazuje do niego informacjê tekstow¹ ("Open") jako „powód uruchomienia”.
End Sub                                   ' Koniec procedury Workbook_Open. Po jej zakoñczeniu Excel dokoñczy otwieranie skoroszytu.

Attribute VB_Name = "RUNNER"
Option Explicit

' G£ÓWNY RUNNER – tu wywo³ujesz WSZYSTKIE swoje makra.
Public Sub MasterRefresh(Optional ByVal reason As String = "")
    Static busy As Boolean
    If busy Then Exit Sub
    busy = True

    On Error GoTo Clean
    Application.EnableEvents = False


     ObrazMikolaja.UpdateSantaByPercent
     RunFinalWithGate
     ShowX_UpToToday_KeepVisible
     OdwrocKarte.OdwrocKarte
    ' ...dowolne kolejne...


Clean:
    Application.EnableEvents = True
    busy = False
End Sub

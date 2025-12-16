# Modules overview

Ten dokument opisuje rolę modułów VBA wyeksportowanych do repozytorium oraz pokazuje, jak łączą się w jedną aplikację.

## Skąd startuje aplikacja

### ThisWorkbook.cls
Rola: automatyczny start aplikacji po otwarciu pliku.  
Główne zdarzenie: Workbook_Open()  
Co robi: wywołuje MasterRefresh("Open"), czyli startowy refresh całej logiki i interfejsu.

---

## Główna logika sterująca

### RUNNER.bas
Rola: centralny runner aplikacji, odpowiedzialny za spójne uruchamianie logiki i synchronizację UI ze stanem.  
Główna procedura: Public Sub MasterRefresh(Optional ByVal reason As String = "")

Co robi:
- zabezpiecza przed wielokrotnym wejściem w odświeżanie
- tymczasowo wyłącza zdarzenia aplikacji
- uruchamia kluczowe moduły zależne od stanu i daty

Wywoływane procedury:
- ObrazMikolaja.UpdateSantaByPercent
- StartFinalu.RunFinalWithGate
- ZnakX.ShowX_UpToToday_KeepVisible
- OdwrocKarte.OdwrocKarte

---

## Stan aplikacji

Trwały stan aplikacji przechowywany jest w arkuszu sterującym „tajne zapiski elfów”.  
Arkusz ten nie jest widoczny ani dostępny dla użytkownika końcowego i pełni wyłącznie funkcję techniczną.

Logika aplikacji opiera się m.in. na następujących kolumnach:
- PoczatkowaData
- KoncowaData
- KomorkaPotwierdzenia
- EtapMikolaja
- EtapPaska
- ProcentDocelowy
- TylKarty
- Final
- X

---

## Moduły funkcjonalne

### OdwrocKarte.bas
Rola: kontrola widoczności „tyłów kart” w zależności od daty.  
Główna procedura: Public Sub OdwrocKarte()

Działanie:
- dla elementów, których PoczatkowaData jest mniejsza lub równa bieżącej dacie, ukrywa warstwę tylną
- pozostałe elementy pozostają widoczne

---

### ZnakX.bas
Rola: wizualne oznaczanie dni, które wygasły i nie zostały zatwierdzone.  
Główna procedura: Public Sub ShowX_UpToToday_KeepVisible()

Działanie:
- pokazuje znak X dla elementów, których KoncowaData minęła i które nie mają statusu DONE
- pozostałe oznaczenia ukrywa

---

### NaklejanieZnaczka.bas
Rola: trwałe zatwierdzenie zadania przez użytkownika.  
Główna procedura: Sub HideAndMarkDone()

Działanie:
- identyfikuje kliknięty element interfejsu
- zapisuje status DONE w arkuszu sterującym
- ukrywa kliknięty element, odsłaniając warstwę pod spodem

---

### ObrazMikolaja.bas
Rola: sterowanie etapami wizualnymi Mikołaja oraz paska postępu.  
Główna procedura: Public Sub UpdateSantaByPercent()

Działanie:
- odczytuje procent wykonania z komórki D27
- wybiera odpowiedni etap wizualny na podstawie ProcentDocelowy
- aktualizuje widoczność elementów graficznych

---

### StartFinalu.bas
Rola: bramka czasowa sterująca dostępem do finału.  
Główna procedura: Public Sub RunFinalWithGate()

Działanie:
- porównuje bieżącą datę z wartością w komórce D28
- przed osiągnięciem daty finału ukrywa wszystkie elementy finałowe
- po spełnieniu warunku uruchamia logikę finału

---

### OknoKoncowe.bas
Rola: obsługa warstwy finałowej aplikacji.  
Główna procedura: Public Sub UpdateFinalByPercent()

Działanie:
- steruje widocznością elementów finałowych na podstawie procentu wykonania
- korzysta z tych samych progów co logika paska postępu

---

## Przepływ działania (skrót)

1. Otwarcie pliku → ThisWorkbook.Workbook_Open  
2. Inicjalizacja → RUNNER.MasterRefresh  
3. Synchronizacja UI:
   - aktualizacja progresu
   - oznaczenia X
   - odwracanie kart
   - kontrola finału
4. Interakcje użytkownika → zapis stanu + aktualizacja UI

Dokument ma charakter techniczny i portfolio.

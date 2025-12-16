# Architecture

## Cel projektu

Kalendarz adwentowy w Excelu zaprojektowany jako mini-aplikacja działająca lokalnie. Interfejs oparty jest na elementach graficznych Excela, natomiast logika sterowana przez VBA reaguje na działania użytkownika, takie jak kliknięcia w okienka czy otwarcie pliku, oraz pilnuje reguł dostępności zadań.

Projekt został zaprojektowany tak, aby zachowywał się jak aplikacja stanowa, a nie statyczny arkusz.

## Główne założenia

- Aplikacja działa w pełni offline, bez połączenia z internetem.
- Interakcja użytkownika odbywa się przez kliknięcia w elementy UI.
- Dostępność okienek zależy od daty oraz zdefiniowanych reguł czasowych.
- Postęp użytkownika jest zapisywany trwale w arkuszu kontrolnym i nie znika po zamknięciu pliku.
- Cofnięcie postępu nie jest możliwe po jego zatwierdzeniu.

## Komponenty

### Warstwa UI (Excel)

- Elementy graficzne Excela pełniące rolę okienek, przycisków i warstw informacyjnych.
- Widoczność i wygląd elementów UI są dynamicznie sterowane przez VBA w zależności od stanu aplikacji.

### Warstwa logiki (VBA)

- Moduły `.bas` odpowiedzialne za poszczególne fragmenty funkcjonalności aplikacji.
- `ThisWorkbook.cls` jako punkt wejścia, uruchamiany automatycznie przy otwarciu pliku.

### Warstwa danych (stan aplikacji)

Trwały stan aplikacji przechowywany jest w arkuszu kontrolnym **„tajne zapiski elfów”**.  
Nazwa arkusza ma charakter fabularny, natomiast technicznie pełni on rolę centralnej tabeli sterującej logiką aplikacji.

Arkusz zawiera m.in. następujące informacje:
- identyfikator dnia / elementu wizualnego
- zakres dat, w których okienko jest aktywne
- informacje o zatwierdzeniu zadania
- etapy wizualne (karty, pasek postępu, elementy finałowe)

## Punkty wejścia (Entry points)

### Otwarcie pliku
`ThisWorkbook.Workbook_Open()` uruchamia procedurę `RUNNER.MasterRefresh()`, która synchronizuje interfejs z zapisanym stanem aplikacji.

### Akcje użytkownika
Kliknięcia w elementy UI wyzwalają makra przypisane do obiektów graficznych. VBA weryfikuje datę, zapisany stan oraz reguły czasowe.

## Uwagi

Dokumentacja ma charakter edukacyjny i portfolio.

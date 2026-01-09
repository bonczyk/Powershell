# PowerShell Scripts Collection

Kolekcja użytecznych skryptów PowerShell do różnych zadań administracyjnych.

## 📊 Generator Raportu Informacji o Komputerze

Zestaw skryptów do zbierania szczegółowych informacji o komputerze i generowania profesjonalnych raportów HTML.

### Pliki

- **`Get-ComputerInfo.ps1`** - Zbiera szczegółowe informacje o systemie
- **`New-ComputerReport.ps1`** - Generuje profesjonalny raport HTML
- **`Generate-ComputerReport.ps1`** - Wygodny skrypt łączący obie funkcjonalności

### Funkcjonalności

Skrypty zbierają następujące informacje:

- ✅ **Informacje o systemie operacyjnym** - nazwa, wersja, architektura, data instalacji
- ✅ **Informacje o BIOS** - producent, wersja, data wydania, numer seryjny
- ✅ **Procesory** - nazwa, producent, liczba rdzeni, częstotliwość, architektura
- ✅ **Pamięć RAM** - całkowita pojemność, wykorzystanie, szczegóły modułów
- ✅ **Dyski logiczne** - partycje, pojemność, wykorzystanie przestrzeni
- ✅ **Dyski fizyczne** - model, interfejs, rozmiar, status
- ✅ **Karty sieciowe** - konfiguracja IP, MAC, DHCP, DNS

### Użycie

#### Metoda 1: Prosty sposób (zalecany)

Użyj głównego skryptu `Generate-ComputerReport.ps1`:

```powershell
# Podstawowe użycie
.\Generate-ComputerReport.ps1

# Z automatycznym otwarciem w przeglądarce
.\Generate-ComputerReport.ps1 -OpenReport

# Z określeniem ścieżki wyjściowej
.\Generate-ComputerReport.ps1 -OutputPath "C:\Reports\MojRaport.html"

# Tryb verbose (szczegółowe informacje)
.\Generate-ComputerReport.ps1 -Verbose -OpenReport
```

#### Metoda 2: Osobne skrypty

Możesz także używać skryptów osobno:

```powershell
# Krok 1: Zbierz dane
$dane = .\Get-ComputerInfo.ps1

# Krok 2: Wygeneruj raport
.\New-ComputerReport.ps1 -ComputerData $dane -OpenReport
```

Lub w jednej linii używając pipeline:

```powershell
.\Get-ComputerInfo.ps1 | .\New-ComputerReport.ps1 -OpenReport
```

### Wymagania

- **System operacyjny:** Windows 7/Server 2008 R2 lub nowszy
- **PowerShell:** Wersja 5.1 lub nowsza
- **Uprawnienia:** Standardowe uprawnienia użytkownika (niektóre informacje mogą wymagać uprawnień administratora)

### Parametry

#### Generate-ComputerReport.ps1

| Parametr | Typ | Opis |
|----------|-----|------|
| `-OutputPath` | String | Opcjonalna ścieżka do pliku wyjściowego HTML |
| `-OpenReport` | Switch | Automatycznie otwiera raport w przeglądarce |
| `-Verbose` | Switch | Wyświetla szczegółowe informacje podczas wykonywania |

#### Get-ComputerInfo.ps1

Funkcja `Get-ComputerSystemInfo`:
- Brak wymaganych parametrów
- Zwraca obiekt PSCustomObject z wszystkimi zebranymi danymi

#### New-ComputerReport.ps1

Funkcja `New-ComputerHTMLReport`:

| Parametr | Typ | Wymagany | Opis |
|----------|-----|----------|------|
| `-ComputerData` | PSCustomObject | Tak | Dane zebrane przez Get-ComputerSystemInfo |
| `-OutputPath` | String | Nie | Ścieżka do pliku HTML |
| `-OpenReport` | Switch | Nie | Otwiera raport w przeglądarce |

### Przykłady

```powershell
# Przykład 1: Szybkie wygenerowanie raportu
.\Generate-ComputerReport.ps1 -OpenReport

# Przykład 2: Raport z własną nazwą pliku
.\Generate-ComputerReport.ps1 -OutputPath "C:\Reports\Serwer-$(Get-Date -Format 'yyyy-MM-dd').html"

# Przykład 3: Zbieranie danych do zmiennej do późniejszego użycia
$info = .\Get-ComputerInfo.ps1
# ... wykonaj inne operacje ...
.\New-ComputerReport.ps1 -ComputerData $info -OutputPath "raport.html"

# Przykład 4: Import jako funkcja
. .\Get-ComputerInfo.ps1
. .\New-ComputerReport.ps1
$data = Get-ComputerSystemInfo -Verbose
New-ComputerHTMLReport -ComputerData $data -OpenReport
```

### Format raportu HTML

Wygenerowany raport HTML zawiera:

- 📊 **Karty podsumowania** - szybki przegląd kluczowych informacji
- 📋 **Tabele szczegółowe** - wszystkie zebrane dane w przejrzystej formie
- 🎨 **Profesjonalny wygląd** - gradient, cienie, responsywny design
- 📱 **Responsywność** - dostosowuje się do różnych rozmiarów ekranu
- 🖨️ **Gotowy do druku** - można łatwo wydrukować lub zapisać jako PDF

### Struktura danych

Obiekt zwracany przez `Get-ComputerSystemInfo` zawiera:

```powershell
PSCustomObject @{
    CollectionDate   # Data i czas zebrania danych
    SystemInfo       # Podstawowe info o systemie
    BIOSInfo         # Informacje o BIOS
    Processors       # Lista procesorów
    Memory           # Podsumowanie pamięci RAM
    MemoryModules    # Szczegóły modułów RAM
    LogicalDisks     # Partycje/dyski logiczne
    PhysicalDisks    # Dyski fizyczne
    NetworkAdapters  # Karty sieciowe
}
```

## 🔧 Inne skrypty

### Compare-GPO.ps1
Skrypt do porównywania obiektów zasad grupy (GPO) w środowisku Active Directory.

### EdgeSeleniumTableToCSV.ps1
Automatyzacja pobierania danych z tabel na stronach internetowych do plików CSV.

### Set-DeploymentWorkWeekSchedule.ps1
Ustawianie harmonogramu wdrożeń na dni robocze tygodnia.

---

## 📝 Licencja

Skrypty są udostępniane "tak jak są" bez żadnych gwarancji.

## 👤 Autor

bonczyk

## 🤝 Współpraca

Wszelkie sugestie i pull requesty są mile widziane!

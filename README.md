# ps.bat

Hybrid batch/PowerShell launcher — runs PowerShell code directly from a `.bat` file.

Hybrydowy launcher batch/PowerShell — uruchamia kod PowerShell bezpośrednio z pliku `.bat`.

---

## How it works / Jak działa

### English

`ps.bat` is a **hybrid batch/PowerShell file** — a single file that is valid both as a Windows batch script (`.bat`) and as PowerShell code. This technique lets you double-click a `.bat` file and have it execute PowerShell code without needing a separate `.ps1` file.

#### Line-by-line breakdown

```bat
@(set "0=%~f0"^)#) & powershell -nop -c iex([io.file]::ReadAllText($env:0)) & exit/b
```

This is the key line — it is interpreted **differently** by CMD and by PowerShell:

**When CMD (batch) processes this line:**

1. `@` — suppresses echoing the command to the console.
2. `(set "0=%~f0"^)` — sets the environment variable `0` to the full path of the current batch file (`%~f0`). The caret `^` escapes the closing parenthesis so CMD does not treat it as the end of a group.
3. `#)` — in CMD this is just harmless text (not a valid command, but it does not cause an error in this context).
4. `& powershell -nop -c iex([io.file]::ReadAllText($env:0))` — launches PowerShell with no profile (`-nop`), and tells it to read the entire contents of the file (whose path is stored in the environment variable `0`) and execute it with `Invoke-Expression` (`iex`).
5. `& exit/b` — exits the batch file after PowerShell finishes, returning control to CMD.

**When PowerShell processes the same line (after reading the file):**

1. `@(set "0=%~f0"^)` — PowerShell tries to evaluate this as an array expression `@(...)`. The call to `set` produces no meaningful output, and then `#)` starts a **comment** (because `#` is the PowerShell comment character), so the rest of the line (`& powershell ...`) is ignored.
2. As a result, PowerShell **skips this line entirely** and proceeds to execute only the remaining PowerShell code below.

```powershell
Write-host 'same window'; pause
```

Writes `"same window"` to the console and waits for a key press. This runs **inside the original CMD window**.

```powershell
$_Paste_in_Powershell = { Write-host 'new window'; pause }
start powershell -args "-nop -c & {`n`n$($_Paste_in_Powershell-replace'"','\"')}"
```

1. Defines a script block `$_Paste_in_Powershell` that writes `"new window"` and pauses.
2. `start powershell` opens a **new PowerShell window** and passes the script block as an argument. The `-replace'"','\"'` escapes any double quotes inside the block so they survive the argument passing.
3. The commented-out `#-verb runas` at the end shows that you could uncomment it to run the new window **as Administrator** (elevated).

#### Summary

The net effect is:
- You double-click `ps.bat` in Windows Explorer.
- CMD runs the first line, which launches PowerShell and feeds it the file contents.
- PowerShell ignores the batch-specific first line (thanks to the `#` comment trick) and executes the rest as normal PowerShell code.
- The script first runs code in the same window, then opens a second PowerShell window.

---

### Polski

`ps.bat` to **hybrydowy plik batch/PowerShell** — pojedynczy plik, który jest poprawny zarówno jako skrypt wsadowy Windows (`.bat`), jak i jako kod PowerShell. Dzięki tej technice można kliknąć dwukrotnie na plik `.bat`, a ten uruchomi kod PowerShell bez potrzeby osobnego pliku `.ps1`.

#### Analiza linia po linii

```bat
@(set "0=%~f0"^)#) & powershell -nop -c iex([io.file]::ReadAllText($env:0)) & exit/b
```

To kluczowa linia — jest interpretowana **inaczej** przez CMD i przez PowerShell:

**Gdy CMD (batch) przetwarza tę linię:**

1. `@` — wyłącza wyświetlanie komendy w konsoli.
2. `(set "0=%~f0"^)` — ustawia zmienną środowiskową `0` na pełną ścieżkę bieżącego pliku batch (`%~f0`). Karetka `^` eskejpuje nawias zamykający, żeby CMD nie potraktował go jako koniec grupy.
3. `#)` — w CMD jest to po prostu nieszkodliwy tekst (nie jest prawidłową komendą, ale w tym kontekście nie powoduje błędu).
4. `& powershell -nop -c iex([io.file]::ReadAllText($env:0))` — uruchamia PowerShell bez ładowania profilu (`-nop`) i każe mu odczytać całą zawartość pliku (którego ścieżka jest zapisana w zmiennej środowiskowej `0`) oraz wykonać ją za pomocą `Invoke-Expression` (`iex`).
5. `& exit/b` — kończy działanie pliku batch po zakończeniu PowerShell, zwracając sterowanie do CMD.

**Gdy PowerShell przetwarza tę samą linię (po odczytaniu pliku):**

1. `@(set "0=%~f0"^)` — PowerShell próbuje to zinterpretować jako wyrażenie tablicowe `@(...)`. Wywołanie `set` nie daje sensownego wyniku, a następnie `#)` rozpoczyna **komentarz** (ponieważ `#` to znak komentarza w PowerShell), więc reszta linii (`& powershell ...`) jest ignorowana.
2. W rezultacie PowerShell **całkowicie pomija tę linię** i przechodzi do wykonania pozostałego kodu PowerShell poniżej.

```powershell
Write-host 'same window'; pause
```

Wypisuje `"same window"` w konsoli i czeka na naciśnięcie klawisza. Działa to **w tym samym oknie CMD**.

```powershell
$_Paste_in_Powershell = { Write-host 'new window'; pause }
start powershell -args "-nop -c & {`n`n$($_Paste_in_Powershell-replace'"','\"')}"
```

1. Definiuje blok skryptu `$_Paste_in_Powershell`, który wypisuje `"new window"` i czeka na klawisz.
2. `start powershell` otwiera **nowe okno PowerShell** i przekazuje blok skryptu jako argument. Wyrażenie `-replace'"','\"'` eskejpuje podwójne cudzysłowy wewnątrz bloku, aby przetrwały przekazywanie argumentów.
3. Zakomentowane `#-verb runas` na końcu pokazuje, że można je odkomentować, aby nowe okno uruchomiło się **jako Administrator** (z podwyższonymi uprawnieniami).

#### Podsumowanie

Efekt końcowy:
- Klikasz dwukrotnie `ps.bat` w Eksploratorze Windows.
- CMD wykonuje pierwszą linię, która uruchamia PowerShell i przekazuje mu zawartość pliku.
- PowerShell ignoruje pierwszą linię specyficzną dla batch (dzięki sztuczce z komentarzem `#`) i wykonuje resztę jako normalny kod PowerShell.
- Skrypt najpierw uruchamia kod w tym samym oknie, a następnie otwiera drugie okno PowerShell.

# Szkolny Asystent 365

Narzędzie w PowerShell dla administratorów szkolnych tenantów **Microsoft 365 Education (A1)**.
Działa w terminalu (menu sterowane strzałkami) na **Linuksie, macOS i Windows**.

Co potrafi:

| # | Funkcja | Co robi |
|---|---------|---------|
| 1 | Raport nieaktywnych skrzynek uczniów | Szuka kont uczniów, które nie logowały się do poczty od N dni |
| 2 | Blokowanie kont z raportu | Wyłącza logowanie kontom z raportu (tylko uczniom) |
| 3 | Odbieranie licencji | Zwalnia licencje zablokowanym kontom uczniów, z kopią zapasową |
| 4 | Pomoc: Microsoft Forms | Instrukcja przygotowania formularza zgłoszeń |
| 5 | Obsługa zgłoszeń z Forms | Masowy reset haseł i zakładanie nowych kont z pliku Excel |
| 6 | Raport HTML do druku | Karteczki z loginem i hasłem startowym dla wychowawców |
| 7 | Usuń pliki z hasłami | Kasuje wyniki i raport HTML po wydrukowaniu |
| 8 | Wyloguj i zamknij | Zamyka sesje Graph i Exchange Online |

---

## Spis treści

- [Wymagania](#wymagania)
- [Instalacja](#instalacja)
- [Uruchomienie](#uruchomienie)
- [Tryb próbny (-DryRun)](#tryb-próbny--dryrun)
- [Opis funkcji](#opis-funkcji)
- [Konfiguracja](#konfiguracja)
- [Pliki tworzone przez narzędzie](#pliki-tworzone-przez-narzędzie)
- [Cofanie zmian](#cofanie-zmian)
- [Bezpieczeństwo i RODO](#bezpieczeństwo-i-rodo)
- [Ograniczenia](#ograniczenia)
- [Rozwiązywanie problemów](#rozwiązywanie-problemów)

---

## Wymagania

- **PowerShell 7+** (`pwsh`). Windows PowerShell 5.1 nie jest obsługiwany.
- Moduły (instalowane automatycznie przy pierwszym uruchomieniu, w zakresie `CurrentUser`):
  - `Microsoft.Graph.Authentication`, `Microsoft.Graph.Users`, `Microsoft.Graph.Users.Actions`,
    `Microsoft.Graph.Identity.DirectoryManagement`
  - `ExchangeOnlineManagement`
  - `ImportExcel` (odczyt plików `.xlsx` bez Excela)
- Konto administratora w tenancie z rolami:
  - **User Administrator** – blokowanie kont, reset haseł, tworzenie kont, licencje,
  - **Exchange Administrator** (lub **Global Reader**) – odczyt statystyk skrzynek (opcja 1),
  - własna skrzynka pocztowa – jeśli chcesz wysyłać powiadomienia e-mail (opcja 5).
- Przy **pierwszym** logowaniu ktoś z rolą Global Administrator musi zatwierdzić uprawnienia
  aplikacji *Microsoft Graph Command Line Tools*:
  `User.ReadWrite.All`, `User-PasswordProfile.ReadWrite.All`, `Organization.Read.All`, `Mail.Send`.

## Instalacja

### Linux (Arch)

```bash
yay -S powershell-bin
```

### Linux (Debian/Ubuntu) i inne systemy

Zobacz instrukcję Microsoft: <https://learn.microsoft.com/powershell/scripting/install/installing-powershell>

### Windows

```powershell
winget install Microsoft.PowerShell
```

### Pobranie narzędzia

```bash
git clone https://github.com/wojcieh/szkolnyasystent365.git
cd szkolnyasystent365
```

Moduły doinstalują się same przy pierwszym starcie. Można je też zainstalować ręcznie:

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Users.Actions, Microsoft.Graph.Identity.DirectoryManagement, ExchangeOnlineManagement, ImportExcel -Scope CurrentUser
```

## Uruchomienie

```bash
pwsh ./szkolnyasystent365.ps1
```

Na Linuksie można też uruchomić plik bezpośrednio (`chmod +x` jest już ustawione):

```bash
./szkolnyasystent365.ps1
```

### Logowanie

Narzędzie loguje się **kodem urządzenia** – działa także na serwerze bez przeglądarki.
Przy starcie zobaczysz **dwa różne kody**, jeden po drugim:

1. Kod dla **Microsoft Graph**
2. Kod dla **Exchange Online**

Każdy z nich wpisz na stronie <https://login.microsoft.com/device> i zaloguj się kontem administratora.
Po zalogowaniu pojawi się menu z nazwą domeny tenanta.

Nawigacja: strzałki **GÓRA/DÓŁ**, wybór: **ENTER**.

## Tryb próbny (-DryRun)

```bash
pwsh ./szkolnyasystent365.ps1 -DryRun
```

W trybie próbnym narzędzie pokazuje, **co by zrobiło**, ale niczego nie zmienia w tenancie:
nie blokuje kont, nie odbiera licencji, nie resetuje haseł i nie tworzy kont.
Raport z opcji 1 jest generowany normalnie.

> **Zalecenie:** każdą większą operację (zwłaszcza pierwszą na danym tenancie) uruchom najpierw z `-DryRun`.

## Opis funkcji

### 1. Raport nieaktywnych skrzynek uczniów

Podajesz liczbę dni (np. `180`). Narzędzie:

1. pobiera konta i licencje z Entra ID oraz listę skrzynek z Exchange Online,
2. **odrzuca konta, które nie są kontami uczniów** (patrz [Kto jest uczniem](#kto-jest-uczniem)),
3. dla uczniów sprawdza `LastLogonTime` skrzynki,
4. zapisuje:
   - `NieaktywniUczniowie.csv` – uczniowie bez logowania dłużej niż N dni,
   - `PominieteKonta.csv` – konta pominięte wraz z powodem (warto przejrzeć!).

Raport jest punktem wyjścia dla opcji 2 i 3. **Przejrzyj go przed blokowaniem** – możesz
usunąć z pliku wiersze uczniów, których nie chcesz blokować.

### 2. Blokowanie kont z raportu

- Wczytuje `NieaktywniUczniowie.csv` i **ponownie weryfikuje** każde konto z aktualnym stanem Entra ID.
- Pomija konta, które nie istnieją, nie są kontami uczniów lub są już zablokowane.
- Pokazuje listę pominiętych i listę do zablokowania.
- Wymaga wpisania **`TAK`** (wielkimi literami).
- Zapisuje dziennik operacji w `logi/blokada_<data>.csv`.

### 3. Odbieranie licencji zablokowanym

- Działa **tylko na kontach już zablokowanych** (najpierw użyj opcji 2).
- Odbiera **wszystkie** licencje przypisane bezpośrednio do konta.
- Przed zmianą zapisuje kopię licencji w `logi/licencje_kopia_<data>.csv`.
- Wymaga wpisania **`USUN`**.
- Licencji przypisanych przez grupę nie da się odebrać tą drogą – trzeba usunąć ucznia z grupy.

### 4. Pomoc: Microsoft Forms

Wyświetla w terminalu instrukcję przygotowania formularza (skrót poniżej).

### 5. Obsługa zgłoszeń z Forms

Nauczyciele zgłaszają przez formularz Microsoft Forms prośby o **reset hasła** lub **nowe konto**
dla ucznia. Narzędzie przetwarza je hurtowo.

#### Przygotowanie formularza

Formularz musi mieć pytania, których nagłówki zawierają (wielkość liter bez znaczenia):

| Kolumna | Rozpoznawane nazwy | Przykład |
|---------|--------------------|----------|
| ID | `ID` (dodawane automatycznie przez Forms) | `17` |
| Typ zgłoszenia | `Czego dotyczy`, `Typ` | `Reset hasła` / `Utworzenie nowego konta` |
| Imię | `Imię`, `Imie` | `Jan` |
| Nazwisko | `Nazwisko` | `Kowalski` |
| Klasa | `Klasa`, `Oddział` | `2B` |
| E-mail nauczyciela | `Email`, `E-mail` (opcjonalnie) | adres zgłaszającego |

Typ zgłoszenia jest rozpoznawany po słowach: **reset / hasło** albo **nowe / utworzenie**.

#### Eksport odpowiedzi

1. Forms → zakładka **Odpowiedzi** → **Otwórz w programie Excel**.
2. Zapisz plik jako `Zgloszenia.xlsx` w folderze narzędzia (albo podaj inną ścieżkę przy uruchomieniu).

#### Jak narzędzie decyduje

| Sytuacja | Działanie |
|----------|-----------|
| Reset hasła, dane pasują do **dokładnie jednego** konta ucznia | Reset hasła, wymuszona zmiana przy logowaniu |
| Reset hasła, kilku uczniów o tych samych danych | Próba rozstrzygnięcia po klasie; jeśli się nie uda – **pominięcie** |
| Reset hasła, dane pasują tylko do konta nauczyciela / wyjątku | **Pominięcie** |
| Nowe konto, takie imię i nazwisko już istnieje | **Pominięcie** (nigdy nie resetuje istniejącego konta) |
| Nowe konto, brak dopasowania | Utworzenie konta `imie.nazwisko@domena` + licencja uczniowska |
| Puste imię/nazwisko lub zapis spoza alfabetu łacińskiego (np. cyrylica) | **Pominięcie** |

Pominięte zgłoszenia są opisane w terminalu i w dzienniku – wymagają obsługi ręcznej.
**Nie są** oznaczane jako przetworzone, więc po wyjaśnieniu można uruchomić opcję 5 ponownie.

Nowe konta dostają: login `imie.nazwisko` (bez polskich znaków; przy kolizji `imie.nazwisko2`),
stanowisko `Uczeń`, dział = klasa, lokalizację `PL` i licencję `STANDARDWOFFPACK_STUDENT`
(lub inną licencję uczniowską z wolnymi miejscami).

Hasła startowe mają postać `Kmpt-4829-Rbwa` – są losowe (generator kryptograficzny) i muszą być
zmienione przy pierwszym logowaniu.

Opcjonalnie narzędzie wysyła nauczycielowi e-mail z danymi ucznia – nadawcą jest zalogowany administrator.

Wyniki trafiają do `wyniki_zgloszen.csv`, a numery zrealizowanych zgłoszeń do `przetworzone_id.txt`
(dzięki temu przy kolejnym uruchomieniu nie zostaną przetworzone drugi raz).

### 6. Raport HTML do druku

Tworzy `Raport_SzkolnyAsystent.html` z tabelą i karteczkami do pocięcia (2 na szerokość strony).
Otwórz plik w przeglądarce, wpisz klasę w pole filtra, kliknij **Drukuj karteczki**.

### 7. Usuń pliki z hasłami

Po wydrukowaniu karteczek **usuń** `wyniki_zgloszen.csv` i `Raport_SzkolnyAsystent.html`
(wymaga wpisania `USUN`). Plik `przetworzone_id.txt` zostaje – nie zawiera haseł.

### 8. Wyloguj i zamknij

Zamyka sesje Microsoft Graph i Exchange Online. Przypomina, jeśli pliki z hasłami nadal istnieją.

## Konfiguracja

Ustawienia znajdują się na początku pliku `szkolnyasystent365.ps1`, w sekcji `KONFIGURACJA`.

### Kto jest uczniem

Konto jest traktowane jako konto ucznia, gdy:

1. **nie** pasuje do listy wyjątków (`$script:ExcludePattern`),
2. **nie** ma licencji nauczycielskiej (`*FACULTY*`),
3. **ma** licencję uczniowską (`*STUDENT*`) **albo** stanowisko (`JobTitle`) zaczynające się od „Uczeń”.

### Lista wyjątków

Konta pasujące do wzorca nigdy nie są blokowane, pozbawiane licencji ani resetowane.
Wzorzec jest sprawdzany na loginie **bez domeny** oraz na nazwie wyświetlanej:

```powershell
$script:ExcludePattern = '(?i)(^s\d+$|stacja|egzaminator|admin|sekretariat|biblioteka|dyrektor|^ucze[nń]\s?\d+$|\btest)'
```

Domyślnie wyklucza m.in. konta pracowni (`s1`…`s24`, `stacja…`), egzaminacyjne, administracyjne
i testowe. Dopisz własne konta, oddzielając je znakiem `|`, np. `...|\btest|pracownia|kiosk)`.

## Pliki tworzone przez narzędzie

Wszystkie pliki powstają w folderze narzędzia (niezależnie od bieżącego katalogu).
Na Linuksie/macOS mają uprawnienia `600` (tylko właściciel), folder `logi/` – `700`.

| Plik | Zawartość | Dane wrażliwe |
|------|-----------|---------------|
| `NieaktywniUczniowie.csv` | Raport z opcji 1 | dane osobowe |
| `PominieteKonta.csv` | Konta pominięte w raporcie i powód | dane osobowe |
| `wyniki_zgloszen.csv` | Wyniki opcji 5 | **hasła** |
| `Raport_SzkolnyAsystent.html` | Karteczki do druku | **hasła** |
| `przetworzone_id.txt` | Numery zrealizowanych zgłoszeń | – |
| `logi/sesja_<data>.log` | Dziennik operacji (bez haseł) | dane osobowe |
| `logi/blokada_<data>.csv` | Dziennik blokowania kont | dane osobowe |
| `logi/licencje_kopia_<data>.csv` | Kopia licencji sprzed odebrania | dane osobowe |

Wszystkie te pliki są w `.gitignore` i **nie trafiają do repozytorium**.

## Cofanie zmian

Po połączeniu (`Connect-MgGraph -Scopes User.ReadWrite.All`) w konsoli `pwsh`:

**Odblokowanie kont** z dziennika blokady:

```powershell
Import-Csv ./logi/blokada_20261117_100325.csv -Delimiter ';' |
    Where-Object wynik -eq 'OK' |
    ForEach-Object { Update-MgUser -UserId $_.id -AccountEnabled:$true }
```

**Przywrócenie licencji** z kopii:

```powershell
Import-Csv ./logi/licencje_kopia_20261117_100325.csv -Delimiter ';' | ForEach-Object {
    $add = @($_.skuIds -split ',' | ForEach-Object { @{ SkuId = $_ } })
    Set-MgUserLicense -UserId $_.id -AddLicenses $add -RemoveLicenses @()
}
```

Nazwy plików podmień na właściwe z folderu `logi/`.

## Bezpieczeństwo i RODO

- Pliki CSV, raporty i dzienniki zawierają **dane osobowe uczniów** – przechowuj je tylko tak długo, jak to konieczne.
- Pliki z **hasłami** usuwaj zaraz po wydrukowaniu (opcja 7).
- Nie wysyłaj raportów pocztą ani nie wrzucaj ich na dyski współdzielone.
- Nie dodawaj plików z danymi do repozytorium (chroni przed tym `.gitignore`).
- Hasła nigdy nie trafiają do dziennika `logi/sesja_*.log`.
- Wszystkie operacje zmieniające tenant wymagają potwierdzenia słowem (`TAK` / `USUN`).

## Ograniczenia

- **`LastLogonTime` dotyczy tylko skrzynki pocztowej.** Uczeń, który korzysta wyłącznie z Teams lub OneDrive,
  może zostać uznany za nieaktywnego. Zawsze przejrzyj raport przed blokowaniem.
- Opcja 1 odpytuje każdą skrzynkę osobno – przy ponad 1000 kont trwa to kilkanaście minut.
- Dopasowanie w opcji 5 wymaga imienia i nazwiska zapisanych alfabetem łacińskim.
- Konta z uszkodzonymi polskimi znakami w nazwie (np. `Szyma�ski`) mogą nie zostać odnalezione
  – warto je poprawić w centrum administracyjnym Microsoft 365.

## Rozwiązywanie problemów

| Objaw | Rozwiązanie |
|-------|-------------|
| Brak kodu logowania po komunikacie „Nawiązywanie połączenia z Microsoft Graph” | Zaktualizuj skrypt do najnowszej wersji (kod był ukrywany w starszej wersji) |
| `AADSTS65001` / prośba o zgodę administratora | Pierwsze logowanie musi wykonać Global Administrator i zatwierdzić uprawnienia |
| „Raport ma nieobsługiwany (stary) format” | Wygeneruj raport ponownie w opcji 1 |
| „Nie udało się rozpoznać wszystkich kluczowych kolumn” | Sprawdź nagłówki w pliku Excel – patrz tabela w opisie opcji 5 |
| `Insufficient privileges` przy resecie hasła | Konto ma rolę administracyjną – takie hasła resetuje się ręcznie |
| Nowe konto ma uwagę `BRAK LICENCJI` | Brak wolnych licencji uczniowskich – zwolnij je (opcje 2–3) i przypisz ręcznie |
| Nie wychodzą e-maile do nauczycieli | Konto administratora musi mieć skrzynkę pocztową w Exchange Online |

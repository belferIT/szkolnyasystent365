#!/usr/bin/env pwsh
#Requires -Version 7.0
<#
.SYNOPSIS
    Szkolny Asystent Microsoft 365 - narzędzie do zarządzania kontami uczniów (sektor Edukacja, A1).
.DESCRIPTION
    - Raport nieaktywnych skrzynek UCZNIÓW (Exchange Online + Entra ID),
    - blokowanie kont i odbieranie licencji z raportu (z kopią stanu sprzed zmian),
    - masowa obsługa zgłoszeń z Microsoft Forms (reset hasła / nowe konto),
    - wydruk karteczek z danymi dostępowymi dla wychowawców.
.PARAMETER DryRun
    Tryb próbny: skrypt pokazuje, co by zrobił, ale nie zmienia niczego w tenancie.
.EXAMPLE
    pwsh ./szkolnyasystent365.ps1 -DryRun
.NOTES
    Wymagane moduły: Microsoft.Graph (Authentication, Users, Identity.DirectoryManagement, Users.Actions),
    ExchangeOnlineManagement, ImportExcel. Wymaga PowerShell 7+.
#>
[CmdletBinding()]
param(
    [switch]$DryRun
)

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================================================
# KONFIGURACJA
# ============================================================================

# Wszystkie pliki robocze trzymamy obok skryptu, niezależnie od bieżącego katalogu.
$script:BaseDir          = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:InactiveCsv      = Join-Path $script:BaseDir 'NieaktywniUczniowie.csv'
$script:SkippedCsv       = Join-Path $script:BaseDir 'PominieteKonta.csv'
$script:ResultsCsv       = Join-Path $script:BaseDir 'wyniki_zgloszen.csv'
$script:ProcessedIdsFile = Join-Path $script:BaseDir 'przetworzone_id.txt'
$script:HtmlReport       = Join-Path $script:BaseDir 'Raport_SzkolnyAsystent.html'
$script:LogDir           = Join-Path $script:BaseDir 'logi'
$script:LogFile          = Join-Path $script:LogDir ("sesja_{0}.log" -f (Get-Date -Format 'yyyy-MM-dd'))

# Konta, których narzędzie NIGDY nie blokuje, nie pozbawia licencji ani nie resetuje
# (sprawdzane na loginie BEZ domeny oraz na nazwie wyświetlanej). Dopisz tu konta pracowni, egzaminacyjne itp.
$script:ExcludePattern = '(?i)(^s\d+$|stacja|egzaminator|admin|sekretariat|biblioteka|dyrektor|^ucze[nń]\s?\d+$|\btest)'

# Uprawnienia Graph (delegowane) - bez Directory.AccessAsUser.All.
$script:GraphScopes = @(
    'User.ReadWrite.All',
    'User-PasswordProfile.ReadWrite.All',
    'Organization.Read.All',
    'Mail.Send'
)

# --- STANY GLOBALNE SESJI ---
$global:CurrentTenantDomain = "BRAK POŁĄCZENIA"
$global:GraphConnected = $false
$global:ExoConnected = $false
$script:UsersById = @{}
$script:SkuNames = @{}
$script:Skus = @()

# ============================================================================
# FUNKCJE POMOCNICZE
# ============================================================================

function Initialize-WorkDir {
    if (-not (Test-Path -Path $script:LogDir)) {
        New-Item -ItemType Directory -Path $script:LogDir -Force | Out-Null
    }
    if ($IsLinux -or $IsMacOS) {
        & chmod 700 -- $script:LogDir 2>$null
    }
}

function Set-PrivatePermissions {
    # Pliki z danymi osobowymi i hasłami - dostęp tylko dla właściciela.
    param([Parameter(Mandatory = $true)][string]$Path)
    if (($IsLinux -or $IsMacOS) -and (Test-Path -Path $Path)) {
        & chmod 600 -- $Path 2>$null
    }
}

function Write-ColorLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][ValidateSet("info", "success", "warn", "error")][string]$Level
    )
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    switch ($Level) {
        "info"    { Write-Host "[$timestamp] [INFO]    $Message" -ForegroundColor Cyan }
        "success" { Write-Host "[$timestamp] [SUKCES]  $Message" -ForegroundColor Green }
        "warn"    { Write-Host "[$timestamp] [UWAGA]   $Message" -ForegroundColor Yellow }
        "error"   { Write-Host "[$timestamp] [BŁĄD]    $Message" -ForegroundColor Red }
    }
    # Ślad audytowy w pliku (bez haseł - hasła nigdy nie trafiają do komunikatów).
    try {
        if (Test-Path -Path $script:LogDir) {
            Add-Content -Path $script:LogFile -Value "[$timestamp] [$($Level.ToUpper())] $Message" -Encoding utf8
            Set-PrivatePermissions -Path $script:LogFile
        }
    }
    catch { }
}

function Wait-AnyKey {
    param([string]$Message = "Naciśnij dowolny klawisz, aby powrócić do menu...")
    Write-Host "`n$Message" -ForegroundColor Gray
    [System.Console]::ReadKey($true) | Out-Null
}

function Remove-PolishChars {
    # Normalizacja nazw do porównań i loginów: małe litery, bez diakrytyków, tylko [a-z0-9].
    # Uwaga: znaki spoza alfabetu łacińskiego (np. cyrylica) dają pusty wynik - wywołujący MUSI to sprawdzić.
    param([Parameter(Mandatory = $false)][string]$InputText)

    if ([string]::IsNullOrWhiteSpace($InputText)) {
        return ""
    }

    $normalized = $InputText.Trim().ToLowerInvariant().Replace('ł', 'l')
    $normalized = $normalized.Normalize([System.Text.NormalizationForm]::FormD)
    $builder = [System.Text.StringBuilder]::new()
    foreach ($ch in $normalized.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$builder.Append($ch)
        }
    }
    return ($builder.ToString() -replace '[^a-z0-9]', '')
}

function New-TempPassword {
    # Kryptograficznie losowe hasło startowe, np. "Kmpt-4829-Rbwa" (~48 bitów entropii,
    # spełnia wymogi złożoności Entra ID). Pominięte znaki mylące: I, l, O, 0, 1.
    $upper  = 'ABCDEFGHJKLMNPRSTUWXYZ'
    $lower  = 'abcdefghijkmnoprstuwxyz'
    $digits = '23456789'
    $pick = {
        param([string]$Set, [int]$Count)
        -join (1..$Count | ForEach-Object { $Set[[System.Security.Cryptography.RandomNumberGenerator]::GetInt32($Set.Length)] })
    }
    return "{0}{1}-{2}-{3}{4}" -f (& $pick $upper 1), (& $pick $lower 3), (& $pick $digits 4), (& $pick $upper 1), (& $pick $lower 3)
}

function ConvertTo-HtmlSafe {
    param([AllowNull()][object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-SafeCsvValue {
    # Ochrona przed wstrzyknięciem formuł przy otwieraniu CSV w Excelu.
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return "" }
    if ($Value -match '^[=+\-@\t\r]') { return "'" + $Value }
    return $Value
}

function Ensure-ModuleInstalled {
    param([Parameter(Mandatory = $true)][string]$ModuleName)

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-ColorLog -Message "Brak wymaganego modułu '$ModuleName'. Rozpoczynam instalację z PSGallery (Scope CurrentUser)..." -Level "warn"
        try {
            Install-Module -Name $ModuleName -Scope CurrentUser -Repository PSGallery -ErrorAction Stop
            Write-ColorLog -Message "Pomyślnie zainstalowano moduł '$ModuleName'." -Level "success"
        }
        catch {
            Write-ColorLog -Message "Nie udało się zainstalować modułu '$ModuleName': $($_.Exception.Message)" -Level "error"
            throw $_
        }
    }
}

# ============================================================================
# KATALOG UŻYTKOWNIKÓW I KLASYFIKACJA KONT
# ============================================================================

function Update-DirectoryCache {
    # Pobiera aktualne konta i licencje z Entra ID. Zwraca listę użytkowników.
    $script:Skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    $script:SkuNames = @{}
    foreach ($s in $script:Skus) {
        $script:SkuNames[[string]$s.SkuId] = $s.SkuPartNumber
    }

    $users = @(Get-MgUser -All -Property Id, DisplayName, GivenName, Surname, UserPrincipalName, Department, JobTitle, AccountEnabled, AssignedLicenses -ErrorAction Stop)
    $script:UsersById = @{}
    foreach ($u in $users) {
        $script:UsersById[[string]$u.Id] = $u
    }
    return $users
}

function Test-ExcludedAccount {
    param([Parameter(Mandatory = $true)]$User)
    $login = ([string]$User.UserPrincipalName).Split('@')[0]
    return ($login -match $script:ExcludePattern) -or ([string]$User.DisplayName -match $script:ExcludePattern)
}

function Get-AccountClass {
    # Zwraca 'uczen' albo powód, dla którego konto NIE jest traktowane jako konto ucznia.
    param([Parameter(Mandatory = $true)]$User)

    if (Test-ExcludedAccount -User $User) {
        return 'wykluczone (lista wyjątków)'
    }
    $skuNames = @($User.AssignedLicenses | ForEach-Object { $script:SkuNames[[string]$_.SkuId] } | Where-Object { $_ })
    if ($skuNames -match 'FACULTY') {
        return 'licencja nauczycielska (FACULTY)'
    }
    if ($skuNames -match 'STUDENT') {
        return 'uczen'
    }
    if ([string]$User.JobTitle -match '(?i)^ucze[nń]') {
        return 'uczen'
    }
    return 'brak licencji ucznia i stanowiska "Uczeń"'
}

function Import-InactiveReport {
    # Wczytuje raport z opcji 1. Zwraca wiersze albo nic (komunikat błędu jest wypisywany tutaj).
    if (-not (Test-Path -Path $script:InactiveCsv)) {
        Write-ColorLog -Message "Plik raportu '$($script:InactiveCsv)' nie istnieje. Wygeneruj go najpierw w Opcji 1." -Level "error"
        return
    }
    $rows = @(Import-Csv -Path $script:InactiveCsv -Delimiter ";" -Encoding utf8)
    if ($rows.Count -eq 0) {
        Write-ColorLog -Message "Raport '$($script:InactiveCsv)' jest pusty." -Level "warn"
        return
    }
    $cols = $rows[0].PSObject.Properties.Name
    if (('id' -notin $cols) -or ('userprincipalname' -notin $cols)) {
        Write-ColorLog -Message "Raport ma nieobsługiwany (stary) format - brak kolumn 'id' i 'userprincipalname'. Wygeneruj go ponownie w Opcji 1." -Level "error"
        return
    }
    return $rows
}

function Resolve-ReportTargets {
    # Weryfikuje wiersze raportu z AKTUALNYM stanem Entra ID: konto musi istnieć i być kontem ucznia.
    param([Parameter(Mandatory = $true)][object[]]$Rows)

    $targets = [System.Collections.Generic.List[object]]::new()
    $skipped = [System.Collections.Generic.List[object]]::new()

    foreach ($r in $Rows) {
        $u = $script:UsersById[[string]$r.id]
        if ($null -eq $u) {
            $skipped.Add([PSCustomObject]@{ UPN = $r.userprincipalname; Powod = 'konto nie istnieje w Entra ID' })
            continue
        }
        $class = Get-AccountClass -User $u
        if ($class -ne 'uczen') {
            $skipped.Add([PSCustomObject]@{ UPN = $u.UserPrincipalName; Powod = $class })
            continue
        }
        $targets.Add($u)
    }
    return [PSCustomObject]@{ Targets = $targets; Skipped = $skipped }
}

function Show-List {
    param([object[]]$Items, [string]$Header, [scriptblock]$Format, [int]$Max = 15)
    if (-not $Items -or $Items.Count -eq 0) { return }
    Write-Host "`n$Header ($($Items.Count)):" -ForegroundColor Yellow
    $Items | Select-Object -First $Max | ForEach-Object { Write-Host "   - $(& $Format $_)" }
    if ($Items.Count -gt $Max) {
        Write-Host "   ... i $($Items.Count - $Max) więcej" -ForegroundColor DarkGray
    }
}

# ============================================================================
# INICJALIZACJA I POŁĄCZENIE ZE ŚRODOWISKIEM
# ============================================================================

function Connect-SzkolnyTenant {
    Clear-Host
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "      INICJALIZACJA POŁĄCZENIA ZE ŚRODOWISKIEM MICROSOFT 365      " -ForegroundColor Cyan
    Write-Host "==================================================================" -ForegroundColor Cyan

    Ensure-ModuleInstalled -ModuleName "ImportExcel"
    Ensure-ModuleInstalled -ModuleName "Microsoft.Graph.Authentication"
    Ensure-ModuleInstalled -ModuleName "Microsoft.Graph.Users"
    Ensure-ModuleInstalled -ModuleName "Microsoft.Graph.Users.Actions"
    Ensure-ModuleInstalled -ModuleName "Microsoft.Graph.Identity.DirectoryManagement"
    Ensure-ModuleInstalled -ModuleName "ExchangeOnlineManagement"

    Write-ColorLog -Message "Nawiązywanie połączenia z Microsoft Graph API (DeviceCode)..." -Level "info"
    try {
        # Kod logowania Graph trafia do strumienia wyjścia - bez Out-Host zostałby połknięty
        # przez "$authSuccess = Connect-SzkolnyTenant" i skrypt czekałby bez komunikatu.
        Write-Host "Otwórz w przeglądarce podany adres i wpisz kod logowania:" -ForegroundColor Yellow
        Connect-MgGraph -UseDeviceAuthentication -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop | Out-Host
        $global:GraphConnected = $true
        Write-ColorLog -Message "Połączono z Microsoft Graph API jako $((Get-MgContext).Account)." -Level "success"
    }
    catch {
        Write-ColorLog -Message "Błąd logowania do Microsoft Graph: $($_.Exception.Message)" -Level "error"
        return $false
    }

    try {
        $org = Get-MgOrganization -Property VerifiedDomains -ErrorAction Stop | Select-Object -First 1
        $defaultDomainObj = $org.VerifiedDomains | Where-Object { $_.IsDefault -eq $true } | Select-Object -First 1
        if (-not $defaultDomainObj) {
            $defaultDomainObj = $org.VerifiedDomains | Select-Object -First 1
        }
        $global:CurrentTenantDomain = $defaultDomainObj.Name
        Write-ColorLog -Message "Rozpoznano aktywny tenant edukacyjny: $($global:CurrentTenantDomain)" -Level "success"
    }
    catch {
        Write-ColorLog -Message "Nie udało się ustalić domeny organizacji: $($_.Exception.Message)" -Level "warn"
        $global:CurrentTenantDomain = "NieznanaDomena"
    }

    Write-ColorLog -Message "Nawiązywanie połączenia z Exchange Online..." -Level "info"
    try {
        Write-Host "Exchange Online wymaga osobnego logowania - wpisz NOWY kod:" -ForegroundColor Yellow
        Connect-ExchangeOnline -Device -ShowProgress:$false -ShowBanner:$false -ErrorAction Stop | Out-Host
        $global:ExoConnected = $true
        Write-ColorLog -Message "Połączono z Exchange Online." -Level "success"
    }
    catch {
        Write-ColorLog -Message "Błąd logowania do Exchange Online: $($_.Exception.Message)" -Level "error"
        return $false
    }

    Start-Sleep -Seconds 2
    return $true
}

# ============================================================================
# TUI: MENU INTERAKTYWNE
# ============================================================================

function Invoke-InteractiveMenu {
    param(
        [Parameter(Mandatory = $true)][string[]]$Options,
        [Parameter(Mandatory = $true)][string]$Title
    )

    $selectedIndex = 0

    while ($true) {
        Clear-Host
        Write-Host "==================================================================" -ForegroundColor DarkCyan
        Write-Host "             ASYSTENT ADMINISTRATORA OFFICE 365 (A1)             " -ForegroundColor White
        Write-Host -NoNewline " Podłączony Tenant: " -ForegroundColor Gray
        Write-Host "$($global:CurrentTenantDomain)" -ForegroundColor Green
        if ($DryRun) {
            Write-Host " TRYB PRÓBNY (-DryRun): żadne zmiany nie zostaną zapisane w tenancie" -ForegroundColor Black -BackgroundColor Yellow
        }
        Write-Host "==================================================================" -ForegroundColor DarkCyan
        Write-Host " $Title" -ForegroundColor Yellow
        Write-Host "------------------------------------------------------------------" -ForegroundColor Gray

        for ($i = 0; $i -lt $Options.Count; $i++) {
            if ($i -eq $selectedIndex) {
                Write-Host "  [>] $($Options[$i])" -ForegroundColor Black -BackgroundColor Cyan
            }
            else {
                Write-Host "      $($Options[$i])" -ForegroundColor White
            }
        }

        Write-Host "------------------------------------------------------------------" -ForegroundColor Gray
        Write-Host " [Nawigacja: Strzałki GÓRA/DÓŁ, Zatwierdzenie: ENTER]" -ForegroundColor DarkGray

        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            ([System.ConsoleKey]::UpArrow) {
                $selectedIndex = if ($selectedIndex -gt 0) { $selectedIndex - 1 } else { $Options.Count - 1 }
            }
            ([System.ConsoleKey]::DownArrow) {
                $selectedIndex = if ($selectedIndex -lt ($Options.Count - 1)) { $selectedIndex + 1 } else { 0 }
            }
            ([System.ConsoleKey]::Enter) {
                return $selectedIndex
            }
        }
    }
}

# ============================================================================
# OPCJA 1: RAPORT NIEAKTYWNYCH SKRZYNEK UCZNIÓW
# ============================================================================

function Invoke-InactiveMailboxReport {
    Clear-Host
    Write-Host "--- GENEROWANIE RAPORTU NIEAKTYWNYCH SKRZYNEK UCZNIÓW (EXCHANGE ONLINE) ---" -ForegroundColor Cyan

    $daysInput = Read-Host "Podaj próg braku logowania w dniach (np. 180)"
    [int]$thresholdDays = 0
    if (-not [int]::TryParse($daysInput, [ref]$thresholdDays) -or $thresholdDays -lt 1) {
        Write-ColorLog -Message "Niepoprawna liczba dni. Użyto wartości domyślnej 180 dni." -Level "warn"
        $thresholdDays = 180
    }
    $cutoffDate = (Get-Date).AddDays(-$thresholdDays)

    try {
        Write-ColorLog -Message "Pobieranie kont i licencji z Entra ID..." -Level "info"
        [void](Update-DirectoryCache)

        Write-ColorLog -Message "Pobieranie listy skrzynek z Exchange Online..." -Level "info"
        $mailboxes = @(Get-ExoMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -Properties UserPrincipalName, DisplayName, ExternalDirectoryObjectId -ErrorAction Stop)
        $total = $mailboxes.Count
        Write-ColorLog -Message "Znaleziono $total skrzynek użytkowników. Analiza dotyczy wyłącznie kont uczniów." -Level "info"

        $inactiveUsers = [System.Collections.Generic.List[PSObject]]::new()
        $skippedUsers = [System.Collections.Generic.List[PSObject]]::new()
        $counter = 0

        foreach ($mbx in $mailboxes) {
            $counter++
            Write-Progress -Activity "Analiza logowań skrzynek pocztowych" -Status "Sprawdzanie: $($mbx.UserPrincipalName)" -PercentComplete (($counter / [Math]::Max($total, 1)) * 100)

            $user = $script:UsersById[[string]$mbx.ExternalDirectoryObjectId]
            if ($null -eq $user) {
                $skippedUsers.Add([PSCustomObject]@{ userprincipalname = $mbx.UserPrincipalName; displayname = $mbx.DisplayName; powod = 'brak konta w Entra ID' })
                continue
            }
            $class = Get-AccountClass -User $user
            if ($class -ne 'uczen') {
                $skippedUsers.Add([PSCustomObject]@{ userprincipalname = $mbx.UserPrincipalName; displayname = $mbx.DisplayName; powod = $class })
                continue
            }

            try {
                $stats = Get-ExoMailboxStatistics -Identity $mbx.UserPrincipalName -Properties LastLogonTime -ErrorAction Stop
                $lastLogon = $stats.LastLogonTime
            }
            catch {
                Write-ColorLog -Message "Brak statystyk dla $($mbx.UserPrincipalName): $($_.Exception.Message)" -Level "warn"
                continue
            }

            if ($null -eq $lastLogon -or $lastLogon -lt $cutoffDate) {
                $inactiveUsers.Add([PSCustomObject]@{
                    "id"                = $user.Id
                    "userprincipalname" = $user.UserPrincipalName
                    "displayname"       = $user.DisplayName
                    "klasa"             = $user.Department
                    "lastlogontime"     = if ($null -eq $lastLogon) { "Nigdy / Brak danych" } else { $lastLogon.ToString("yyyy-MM-dd HH:mm:ss") }
                    "status"            = "Nieaktywny"
                })
            }
        }
        Write-Progress -Activity "Analiza logowań skrzynek pocztowych" -Completed

        $inactiveUsers | Export-Csv -Path $script:InactiveCsv -NoTypeInformation -Delimiter ";" -Encoding utf8
        $skippedUsers | Export-Csv -Path $script:SkippedCsv -NoTypeInformation -Delimiter ";" -Encoding utf8
        Set-PrivatePermissions -Path $script:InactiveCsv
        Set-PrivatePermissions -Path $script:SkippedCsv

        Write-ColorLog -Message "Wygenerowano raport: $($script:InactiveCsv) (nieaktywni uczniowie: $($inactiveUsers.Count))" -Level "success"
        Write-ColorLog -Message "Konta pominięte (nie-uczniowie / wyjątki): $($skippedUsers.Count) - lista w $($script:SkippedCsv)" -Level "info"
        Write-Host "`nUWAGA: LastLogonTime dotyczy tylko skrzynki pocztowej. Uczeń korzystający wyłącznie z Teams/OneDrive" -ForegroundColor DarkYellow
        Write-Host "może zostać uznany za nieaktywnego. Przejrzyj raport przed blokowaniem kont." -ForegroundColor DarkYellow
    }
    catch {
        Write-ColorLog -Message "Wystąpił błąd podczas pobierania danych: $($_.Exception.Message)" -Level "error"
    }

    Wait-AnyKey
}

# ============================================================================
# OPCJA 2: BLOKOWANIE KONT Z RAPORTU
# ============================================================================

function Invoke-BlockInactiveAccounts {
    Clear-Host
    Write-Host "--- BLOKOWANIE KONT Z RAPORTU NIEAKTYWNYCH (MICROSOFT GRAPH) ---" -ForegroundColor Cyan

    $rows = @(Import-InactiveReport)
    if ($rows.Count -eq 0) { Wait-AnyKey; return }

    try {
        Write-ColorLog -Message "Weryfikacja kont z raportu z aktualnym stanem Entra ID..." -Level "info"
        [void](Update-DirectoryCache)
    }
    catch {
        Write-ColorLog -Message "Nie udało się pobrać kont z Entra ID: $($_.Exception.Message)" -Level "error"
        Wait-AnyKey; return
    }

    $resolved = Resolve-ReportTargets -Rows $rows
    $toBlock = @($resolved.Targets | Where-Object { $_.AccountEnabled -eq $true })
    $alreadyBlocked = $resolved.Targets.Count - $toBlock.Count

    Show-List -Items @($resolved.Skipped) -Header "Pominięte (nie zostaną zablokowane)" -Format { param($x) "$($x.UPN) - $($x.Powod)" }
    if ($alreadyBlocked -gt 0) {
        Write-ColorLog -Message "Już zablokowanych w raporcie: $alreadyBlocked (bez zmian)." -Level "info"
    }
    if ($toBlock.Count -eq 0) {
        Write-ColorLog -Message "Brak kont uczniów do zablokowania." -Level "info"
        Wait-AnyKey; return
    }
    Show-List -Items $toBlock -Header "Do zablokowania" -Format { param($x) "$($x.UserPrincipalName) ($($x.Department))" }

    if ($DryRun) {
        Write-ColorLog -Message "[DRY-RUN] Zablokowano by $($toBlock.Count) kont. Nic nie zmieniono." -Level "warn"
        Wait-AnyKey; return
    }

    $confirm = Read-Host "`nCZY NA PEWNO chcesz zablokować logowanie dla $($toBlock.Count) kont? Wpisz 'TAK', aby kontynuować"
    if ($confirm -cne "TAK") {
        Write-ColorLog -Message "Operacja anulowana przez administratora." -Level "info"
        Wait-AnyKey; return
    }

    $logPath = Join-Path $script:LogDir ("blokada_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $log = [System.Collections.Generic.List[object]]::new()
    $counter = 0
    $successCount = 0
    foreach ($u in $toBlock) {
        $counter++
        Write-Progress -Activity "Blokowanie kont użytkowników" -Status "Przetwarzanie: $($u.UserPrincipalName)" -PercentComplete (($counter / $toBlock.Count) * 100)
        $result = 'OK'
        try {
            Update-MgUser -UserId $u.Id -BodyParameter @{ accountEnabled = $false } -ErrorAction Stop
            $successCount++
        }
        catch {
            $result = "BŁĄD: $($_.Exception.Message)"
            Write-ColorLog -Message "Błąd blokowania $($u.UserPrincipalName): $($_.Exception.Message)" -Level "error"
        }
        $log.Add([PSCustomObject]@{ id = $u.Id; userprincipalname = $u.UserPrincipalName; accountEnabledPrzed = $u.AccountEnabled; wynik = $result; data = (Get-Date -Format 's') })
    }
    Write-Progress -Activity "Blokowanie kont użytkowników" -Completed

    $log | Export-Csv -Path $logPath -NoTypeInformation -Delimiter ";" -Encoding utf8
    Set-PrivatePermissions -Path $logPath
    Write-ColorLog -Message "Zablokowano $successCount z $($toBlock.Count) kont. Dziennik operacji: $logPath" -Level "success"
    Wait-AnyKey
}

# ============================================================================
# OPCJA 3: ODBIERANIE LICENCJI ZABLOKOWANYM
# ============================================================================

function Invoke-RevokeLicensesFromReport {
    Clear-Host
    Write-Host "--- ODBIERANIE LICENCJI ZABLOKOWANYM KONTOM (MICROSOFT GRAPH) ---" -ForegroundColor Cyan

    $rows = @(Import-InactiveReport)
    if ($rows.Count -eq 0) { Wait-AnyKey; return }

    try {
        Write-ColorLog -Message "Weryfikacja kont z raportu z aktualnym stanem Entra ID..." -Level "info"
        [void](Update-DirectoryCache)
    }
    catch {
        Write-ColorLog -Message "Nie udało się pobrać kont z Entra ID: $($_.Exception.Message)" -Level "error"
        Wait-AnyKey; return
    }

    $resolved = Resolve-ReportTargets -Rows $rows
    $skipped = [System.Collections.Generic.List[object]]::new($resolved.Skipped)
    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($u in $resolved.Targets) {
        if ($u.AccountEnabled -ne $false) {
            $skipped.Add([PSCustomObject]@{ UPN = $u.UserPrincipalName; Powod = 'konto nadal aktywne - najpierw zablokuj (Opcja 2)' })
        }
        elseif (@($u.AssignedLicenses).Count -gt 0) {
            $targets.Add($u)
        }
    }

    Show-List -Items @($skipped) -Header "Pominięte" -Format { param($x) "$($x.UPN) - $($x.Powod)" }
    if ($targets.Count -eq 0) {
        Write-ColorLog -Message "Brak zablokowanych kont uczniów z przypisanymi licencjami." -Level "info"
        Wait-AnyKey; return
    }
    Show-List -Items @($targets) -Header "Do odebrania licencji" -Format {
        param($x) "$($x.UserPrincipalName): " + ((@($x.AssignedLicenses) | ForEach-Object { $script:SkuNames[[string]$_.SkuId] }) -join ', ')
    }

    if ($DryRun) {
        Write-ColorLog -Message "[DRY-RUN] Odebrano by licencje $($targets.Count) kontom. Nic nie zmieniono." -Level "warn"
        Wait-AnyKey; return
    }

    $confirm = Read-Host "`nOdebrać WSZYSTKIE licencje $($targets.Count) zablokowanym kontom uczniów? Wpisz 'USUN', aby potwierdzić"
    if ($confirm -cne "USUN") {
        Write-ColorLog -Message "Operacja anulowana przez użytkownika." -Level "info"
        Wait-AnyKey; return
    }

    # Kopia zapasowa licencji PRZED zmianami - pozwala je przywrócić (Set-MgUserLicense -AddLicenses).
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backupPath = Join-Path $script:LogDir "licencje_kopia_$stamp.csv"
    $targets | ForEach-Object {
        [PSCustomObject]@{
            id                = $_.Id
            userprincipalname = $_.UserPrincipalName
            skuIds            = (@($_.AssignedLicenses).SkuId -join ',')
            skuNazwy          = ((@($_.AssignedLicenses) | ForEach-Object { $script:SkuNames[[string]$_.SkuId] }) -join ',')
        }
    } | Export-Csv -Path $backupPath -NoTypeInformation -Delimiter ";" -Encoding utf8
    Set-PrivatePermissions -Path $backupPath
    Write-ColorLog -Message "Zapisano kopię licencji: $backupPath" -Level "info"

    $counter = 0
    $modifiedCount = 0
    foreach ($u in $targets) {
        $counter++
        Write-Progress -Activity "Odbieranie licencji użytkownikom" -Status "Przetwarzanie: $($u.UserPrincipalName)" -PercentComplete (($counter / $targets.Count) * 100)
        try {
            $skuIdsToRemove = @(@($u.AssignedLicenses).SkuId)
            Set-MgUserLicense -UserId $u.Id -AddLicenses @() -RemoveLicenses $skuIdsToRemove -ErrorAction Stop | Out-Null
            $modifiedCount++
        }
        catch {
            Write-ColorLog -Message "Błąd odbierania licencji $($u.UserPrincipalName) (licencje grupowe usuwa się w grupie): $($_.Exception.Message)" -Level "error"
        }
    }
    Write-Progress -Activity "Odbieranie licencji użytkownikom" -Completed

    Write-ColorLog -Message "Zakończono. Odebrano licencje z $modifiedCount z $($targets.Count) kont." -Level "success"
    Wait-AnyKey
}

# ============================================================================
# OPCJA 4: POMOC / INSTRUKCJA FORMS
# ============================================================================

function Show-FormsHelpScreen {
    Clear-Host
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "       INSTRUKCJA OBSŁUGI ZGŁOSZEŃ FORMULARZY MICROSOFT FORMS     " -ForegroundColor White
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. STRUKTURA FORMULARZA MICROSOFT FORMS:" -ForegroundColor Yellow
    Write-Host "   Aby aplikacja prawidłowo zmapowała dane, formularz musi zawierać"
    Write-Host "   następujące pola w nagłówkach kolumn (wielkość liter nie ma znaczenia):"
    Write-Host "   - ID                  : Identyfikator odpowiedzi Forms (generowany auto)"
    Write-Host "   - Czego dotyczy / Typ : 'Reset hasła' lub 'Utworzenie nowego konta'"
    Write-Host "   - Imię                : Imię ucznia (alfabet łaciński)"
    Write-Host "   - Nazwisko            : Nazwisko ucznia (alfabet łaciński)"
    Write-Host "   - Klasa / Oddział     : np. 1A, 2B, 3Cp"
    Write-Host "   - Email / E-mail      : Adres e-mail nauczyciela zgłaszającego"
    Write-Host ""
    Write-Host "2. EKSPORT DANYCH:" -ForegroundColor Yellow
    Write-Host "   - Otwórz swój formularz w MS Forms -> Zakładka 'Odpowiedzi'"
    Write-Host "   - Kliknij 'Otwórz w programie Excel'"
    Write-Host "   - Zapisz pobrany plik jako 'Zgloszenia.xlsx' w folderze skryptu."
    Write-Host ""
    Write-Host "3. ZASADY BEZPIECZEŃSTWA:" -ForegroundColor Yellow
    Write-Host "   - Reset hasła wykonywany jest tylko, gdy dane wskazują DOKŁADNIE jedno konto ucznia."
    Write-Host "   - Kolizje (kilku uczniów o tych samych danych) i konta nie-uczniów są pomijane"
    Write-Host "     i wymagają ręcznej obsługi. Zgłoszenie 'nowe konto' nigdy nie resetuje hasła."
    Write-Host "   - Uruchom skrypt z parametrem -DryRun, aby zobaczyć efekt bez wprowadzania zmian."
    Write-Host ""
    Write-Host "4. MECHANIZM ANTY-DUPLIKATOWY:" -ForegroundColor Yellow
    Write-Host "   - Każde zrealizowane zgłoszenie odnotowuje swój numer ID w pliku"
    Write-Host "     'przetworzone_id.txt'. Pominięte zgłoszenia NIE są oznaczane - po ręcznym"
    Write-Host "     wyjaśnieniu można je przetworzyć ponownie."
    Write-Host ""
    Write-Host "5. WYDRUK KARTECZEK DLA WYCHOWAWCÓW:" -ForegroundColor Yellow
    Write-Host "   - Po zakończeniu procedury w Opcji 5, przejdź do Opcji 6."
    Write-Host "   - Po wydrukowaniu USUŃ pliki z hasłami (Opcja 7)."
    Write-Host ""
    Wait-AnyKey -Message "Naciśnij dowolny klawisz, aby powrócić do menu głównego..."
}

# ============================================================================
# OPCJA 5: OBSŁUGA ZGŁOSZEŃ Z FORMS (MASOWY RESET I TWORZENIE)
# ============================================================================

function New-NameIndex {
    # Indeks: znormalizowane "imie|nazwisko" (w obu kolejnościach) oraz "dn|<nazwa wyświetlana>" -> lista kont.
    param([object[]]$Users)
    $index = @{}
    foreach ($u in $Users) {
        Add-ToNameIndex -Index $index -User $u
    }
    return $index
}

function Add-ToNameIndex {
    param([hashtable]$Index, $User)
    $g = Remove-PolishChars -InputText $User.GivenName
    $s = Remove-PolishChars -InputText $User.Surname
    $d = Remove-PolishChars -InputText $User.DisplayName
    $keys = [System.Collections.Generic.HashSet[string]]::new()
    if ($g -and $s) {
        [void]$keys.Add("$g|$s")
        [void]$keys.Add("$s|$g")
    }
    if ($d) {
        [void]$keys.Add("dn|$d")
    }
    foreach ($k in $keys) {
        if (-not $Index.ContainsKey($k)) {
            $Index[$k] = [System.Collections.Generic.List[object]]::new()
        }
        $Index[$k].Add($User)
    }
}

function Find-UsersByName {
    # Dokładne dopasowanie (bez symboli wieloznacznych). Wymaga niepustego imienia i nazwiska.
    param([hashtable]$Index, [string]$First, [string]$Last)
    if ([string]::IsNullOrEmpty($First) -or [string]::IsNullOrEmpty($Last)) {
        return
    }
    $found = @{}
    foreach ($k in @("$First|$Last", "dn|$First$Last", "dn|$Last$First")) {
        if ($Index.ContainsKey($k)) {
            foreach ($u in $Index[$k]) {
                $found[[string]$u.Id] = $u
            }
        }
    }
    return @($found.Values)
}

function Get-StudentSku {
    $sku = $script:Skus | Where-Object { $_.SkuPartNumber -eq 'STANDARDWOFFPACK_STUDENT' } | Select-Object -First 1
    if ($null -eq $sku) {
        $sku = $script:Skus |
            Where-Object { $_.SkuPartNumber -match 'STUDENT' -and $_.SkuPartNumber -notmatch 'BENEFIT' } |
            Sort-Object { $_.PrepaidUnits.Enabled - $_.ConsumedUnits } -Descending |
            Select-Object -First 1
    }
    return $sku
}

function Add-ResultRow {
    param([System.Collections.Specialized.OrderedDictionary]$Row)
    $safe = [ordered]@{}
    foreach ($entry in $Row.GetEnumerator()) {
        $safe[$entry.Key] = ConvertTo-SafeCsvValue -Value ([string]$entry.Value)
    }
    [PSCustomObject]$safe | Export-Csv -Path $script:ResultsCsv -Append -Force -NoTypeInformation -Delimiter ';' -Encoding utf8
    Set-PrivatePermissions -Path $script:ResultsCsv
}

function Add-ProcessedId {
    param([string]$Id, [System.Collections.Generic.HashSet[string]]$Set)
    [void]$Set.Add($Id)
    Add-Content -Path $script:ProcessedIdsFile -Value $Id -Encoding utf8
}

function Invoke-ProcessFormsRequests {
    Clear-Host
    Write-Host "--- MASOWE PRZETWARZANIE ZGŁOSZEŃ (MICROSOFT FORMS / EXCEL) ---" -ForegroundColor Cyan
    if ($DryRun) {
        Write-Host "TRYB PRÓBNY: hasła nie zostaną zmienione, konta nie zostaną utworzone." -ForegroundColor Yellow
    }

    $defaultXlsx = Join-Path $script:BaseDir 'Zgloszenia.xlsx'
    $filePath = Read-Host "Podaj ścieżkę do pliku Excela [Domyślnie: $defaultXlsx]"
    if ([string]::IsNullOrWhiteSpace($filePath)) {
        $filePath = $defaultXlsx
    }

    if (-not (Test-Path -Path $filePath)) {
        Write-ColorLog -Message "Nie odnaleziono pliku: $filePath" -Level "error"
        Wait-AnyKey; return
    }

    $shouldSendMail = $false
    if (-not $DryRun) {
        $sendEmailChoice = Read-Host "Czy wysyłać powiadomienia e-mail do nauczycieli zgłaszających? (TAK/NIE) [Domyślnie: NIE]"
        $shouldSendMail = ($sendEmailChoice -eq "TAK")
    }

    $processedIds = [System.Collections.Generic.HashSet[string]]::new()
    if (Test-Path -Path $script:ProcessedIdsFile) {
        foreach ($l in (Get-Content -Path $script:ProcessedIdsFile -Encoding utf8)) {
            if (-not [string]::IsNullOrWhiteSpace($l)) {
                [void]$processedIds.Add($l.Trim())
            }
        }
    }

    Write-ColorLog -Message "Wczytywanie arkusza Excel..." -Level "info"
    try {
        $excelData = @(Import-Excel -Path $filePath)
    }
    catch {
        Write-ColorLog -Message "Błąd otwierania pliku Excel: $($_.Exception.Message)" -Level "error"
        Wait-AnyKey; return
    }

    if ($excelData.Count -eq 0) {
        Write-ColorLog -Message "Arkusz Excel jest pusty." -Level "warn"
        Wait-AnyKey; return
    }

    $propNames = $excelData[0].PSObject.Properties.Name
    $colId       = $propNames | Where-Object { $_ -match '(?i)^id$' } | Select-Object -First 1
    $colTyp      = $propNames | Where-Object { $_ -match '(?i)(czego dotyczy|typ)' } | Select-Object -First 1
    $colImie     = $propNames | Where-Object { $_ -match '(?i)(imi[eę])' } | Select-Object -First 1
    $colNazwisko = $propNames | Where-Object { $_ -match '(?i)nazwisko' } | Select-Object -First 1
    $colKlasa    = $propNames | Where-Object { $_ -match '(?i)(klasa|oddzia[lł])' } | Select-Object -First 1
    $colEmail    = $propNames | Where-Object { $_ -match '(?i)(e-?mail)' } | Select-Object -First 1

    if (-not ($colId -and $colTyp -and $colImie -and $colNazwisko -and $colKlasa)) {
        Write-ColorLog -Message "Nie udało się rozpoznać wszystkich kluczowych kolumn w arkuszu." -Level "error"
        Write-Host "Wykryte kolumny: ID=$colId, Typ=$colTyp, Imię=$colImie, Nazwisko=$colNazwisko, Klasa=$colKlasa, Email=$colEmail" -ForegroundColor Yellow
        Wait-AnyKey; return
    }

    Write-ColorLog -Message "Pobieranie bazy kont i licencji z Entra ID..." -Level "info"
    try {
        $allUsers = Update-DirectoryCache
    }
    catch {
        Write-ColorLog -Message "Błąd ładowania użytkowników: $($_.Exception.Message)" -Level "error"
        Wait-AnyKey; return
    }
    $nameIndex = New-NameIndex -Users $allUsers
    $upnSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($u in $allUsers) { [void]$upnSet.Add([string]$u.UserPrincipalName) }
    Write-ColorLog -Message "Zaindeksowano $($allUsers.Count) kont." -Level "success"

    $studentSku = Get-StudentSku
    if ($null -eq $studentSku) {
        Write-ColorLog -Message "Nie odnaleziono licencji uczniowskiej (STUDENT)! Nowe konta nie otrzymają licencji." -Level "warn"
    }
    else {
        $free = $studentSku.PrepaidUnits.Enabled - $studentSku.ConsumedUnits
        Write-ColorLog -Message "Licencja dla nowych kont: $($studentSku.SkuPartNumber) (wolnych: $free)" -Level "info"
        if ($free -le 0) {
            Write-ColorLog -Message "Brak wolnych licencji $($studentSku.SkuPartNumber) - nowe konta mogą zostać bez licencji." -Level "warn"
        }
    }

    $processedCounter = 0
    $totalRows = $excelData.Count
    $stats = @{ reset = 0; nowe = 0; pominiete = 0; bledy = 0 }

    foreach ($row in $excelData) {
        $processedCounter++
        $reqId = ([string]$row.$colId).Trim()
        if ([string]::IsNullOrWhiteSpace($reqId) -or $processedIds.Contains($reqId)) {
            continue
        }

        $rawTyp      = ([string]$row.$colTyp).Trim()
        $rawImie     = ([string]$row.$colImie).Trim()
        $rawNazwisko = ([string]$row.$colNazwisko).Trim()
        $rawKlasa    = ([string]$row.$colKlasa).Trim()
        $rawEmail    = if ($colEmail) { ([string]$row.$colEmail).Trim() } else { "" }

        Write-Progress -Activity "Przetwarzanie zgłoszeń z Forms" -Status "Wiersz $processedCounter / $totalRows : $rawImie $rawNazwisko" -PercentComplete (($processedCounter / $totalRows) * 100)

        $cleanImie = Remove-PolishChars -InputText $rawImie
        $cleanNazwisko = Remove-PolishChars -InputText $rawNazwisko

        if (-not $cleanImie -or -not $cleanNazwisko) {
            Write-ColorLog -Message "Pominięto ID $reqId : imię lub nazwisko puste albo zapisane poza alfabetem łacińskim ('$rawImie $rawNazwisko'). Obsłuż ręcznie." -Level "warn"
            $stats.pominiete++
            continue
        }

        $isReset = $rawTyp -match '(?i)(reset|has[lł])'
        $isNew = $rawTyp -match '(?i)(nowe|utworz)'
        if ($isReset -eq $isNew) {
            Write-ColorLog -Message "Pominięto ID $reqId : nie rozpoznano typu zgłoszenia ('$rawTyp')." -Level "warn"
            $stats.pominiete++
            continue
        }

        $candidates = @(Find-UsersByName -Index $nameIndex -First $cleanImie -Last $cleanNazwisko)

        # ŚCIEŻKA A: RESET HASŁA - tylko gdy dane wskazują dokładnie jedno konto ucznia
        if ($isReset) {
            $studentCandidates = @($candidates | Where-Object { (Get-AccountClass -User $_) -eq 'uczen' })
            if ($studentCandidates.Count -gt 1) {
                $cleanKlasa = Remove-PolishChars -InputText $rawKlasa
                $studentCandidates = @($studentCandidates | Where-Object { (Remove-PolishChars -InputText $_.Department) -eq $cleanKlasa })
            }

            if ($studentCandidates.Count -ne 1) {
                $reason = if ($candidates.Count -eq 0) { "nie znaleziono konta" }
                          elseif ($studentCandidates.Count -eq 0) { "brak pasującego konta ucznia (dopasowania: $(@($candidates.UserPrincipalName) -join ', '))" }
                          else { "kolizja - kilka kont: $(@($studentCandidates.UserPrincipalName) -join ', ')" }
                Write-ColorLog -Message "Pominięto ID $reqId ($rawImie $rawNazwisko, $rawKlasa): $reason. Obsłuż ręcznie." -Level "warn"
                $stats.pominiete++
                continue
            }

            $targetUser = $studentCandidates[0]
            $finalClass = if (-not [string]::IsNullOrWhiteSpace($targetUser.Department)) { $targetUser.Department } else { $rawKlasa }

            if ($DryRun) {
                Write-ColorLog -Message "[DRY-RUN] ID $reqId : zresetowano by hasło dla $($targetUser.UserPrincipalName)" -Level "info"
                $stats.reset++
                continue
            }

            $tempPassword = New-TempPassword
            try {
                $updateParams = @{
                    passwordProfile = @{
                        password                      = $tempPassword
                        forceChangePasswordNextSignIn = $true
                    }
                }
                Update-MgUser -UserId $targetUser.Id -BodyParameter $updateParams -ErrorAction Stop

                Add-ResultRow -Row ([ordered]@{
                    ID_Zgloszenia = $reqId; Akcja = 'Reset hasla'; Klasa = $finalClass; Imie = $rawImie; Nazwisko = $rawNazwisko
                    Login = $targetUser.UserPrincipalName; HasloStartowe = $tempPassword; Nauczyciel = $rawEmail; Uwagi = ''
                })
                Add-ProcessedId -Id $reqId -Set $processedIds
                $stats.reset++
                Write-ColorLog -Message "Zresetowano hasło dla: $($targetUser.UserPrincipalName) (ID $reqId)" -Level "success"

                if ($shouldSendMail -and -not [string]::IsNullOrWhiteSpace($rawEmail)) {
                    Send-NotificationEmail -ToEmail $rawEmail -StudentName "$rawImie $rawNazwisko" -StudentLogin $targetUser.UserPrincipalName -TempPassword $tempPassword -StudentClass $finalClass -ActionType "Reset hasła"
                }
            }
            catch {
                $stats.bledy++
                Write-ColorLog -Message "Błąd resetowania hasła dla ID $reqId ($($targetUser.UserPrincipalName)): $($_.Exception.Message)" -Level "error"
            }
            continue
        }

        # ŚCIEŻKA B: NOWE KONTO - nigdy nie resetuje istniejącego konta
        if ($candidates.Count -gt 0) {
            Write-ColorLog -Message "Pominięto ID $reqId : wniosek o nowe konto, ale istnieje już konto '$rawImie $rawNazwisko': $(@($candidates.UserPrincipalName) -join ', '). Zweryfikuj ręcznie." -Level "warn"
            $stats.pominiete++
            continue
        }

        $baseLogin = "$cleanImie.$cleanNazwisko"
        $finalUpn = "$baseLogin@$($global:CurrentTenantDomain)"
        $suffix = 2
        while ($upnSet.Contains($finalUpn)) {
            $finalUpn = "$baseLogin$suffix@$($global:CurrentTenantDomain)"
            $suffix++
        }

        if ($DryRun) {
            Write-ColorLog -Message "[DRY-RUN] ID $reqId : utworzono by konto $finalUpn ($rawKlasa)" -Level "info"
            [void]$upnSet.Add($finalUpn)
            $stats.nowe++
            continue
        }

        $tempPassword = New-TempPassword
        try {
            $newUserParams = @{
                accountEnabled    = $true
                displayName       = "$rawImie $rawNazwisko"
                givenName         = $rawImie
                surname           = $rawNazwisko
                userPrincipalName = $finalUpn
                mailNickname      = $finalUpn.Split('@')[0] -replace '[^a-zA-Z0-9.\-_]', ''
                department        = $rawKlasa
                jobTitle          = "Uczeń"
                usageLocation     = "PL"
                passwordProfile   = @{
                    password                      = $tempPassword
                    forceChangePasswordNextSignIn = $true
                }
            }
            $newUser = New-MgUser -BodyParameter $newUserParams -ErrorAction Stop
            [void]$upnSet.Add($finalUpn)

            $note = ''
            if ($null -ne $studentSku) {
                try {
                    Set-MgUserLicense -UserId $newUser.Id -AddLicenses @(@{ SkuId = $studentSku.SkuId }) -RemoveLicenses @() -ErrorAction Stop | Out-Null
                }
                catch {
                    $note = 'BRAK LICENCJI'
                    Write-ColorLog -Message "Konto $finalUpn utworzone, ale NIE przypisano licencji: $($_.Exception.Message)" -Level "warn"
                }
            }
            else {
                $note = 'BRAK LICENCJI'
            }

            $newEntry = [PSCustomObject]@{
                Id                = $newUser.Id
                DisplayName       = "$rawImie $rawNazwisko"
                GivenName         = $rawImie
                Surname           = $rawNazwisko
                UserPrincipalName = $finalUpn
                Department        = $rawKlasa
                JobTitle          = "Uczeń"
                AccountEnabled    = $true
                AssignedLicenses  = @()
            }
            Add-ToNameIndex -Index $nameIndex -User $newEntry

            Add-ResultRow -Row ([ordered]@{
                ID_Zgloszenia = $reqId; Akcja = 'Utworzenie konta'; Klasa = $rawKlasa; Imie = $rawImie; Nazwisko = $rawNazwisko
                Login = $finalUpn; HasloStartowe = $tempPassword; Nauczyciel = $rawEmail; Uwagi = $note
            })
            Add-ProcessedId -Id $reqId -Set $processedIds
            $stats.nowe++
            Write-ColorLog -Message "Utworzono nowe konto ucznia: $finalUpn (ID $reqId)" -Level "success"

            if ($shouldSendMail -and -not [string]::IsNullOrWhiteSpace($rawEmail)) {
                Send-NotificationEmail -ToEmail $rawEmail -StudentName "$rawImie $rawNazwisko" -StudentLogin $finalUpn -TempPassword $tempPassword -StudentClass $rawKlasa -ActionType "Nowe konto"
            }
        }
        catch {
            $stats.bledy++
            Write-ColorLog -Message "Błąd tworzenia konta dla ID $reqId ($rawImie $rawNazwisko): $($_.Exception.Message)" -Level "error"
        }
    }

    Write-Progress -Activity "Przetwarzanie zgłoszeń z Forms" -Completed
    Write-ColorLog -Message "Zakończono. Reset: $($stats.reset), nowe konta: $($stats.nowe), pominięte (do ręcznej obsługi): $($stats.pominiete), błędy: $($stats.bledy)." -Level "success"
    if (-not $DryRun -and (Test-Path -Path $script:ResultsCsv)) {
        Write-ColorLog -Message "Wyniki z hasłami: $($script:ResultsCsv) - usuń po wydruku (Opcja 7)." -Level "warn"
    }
    Wait-AnyKey
}

function Send-NotificationEmail {
    param(
        [string]$ToEmail,
        [string]$StudentName,
        [string]$StudentLogin,
        [string]$TempPassword,
        [string]$StudentClass,
        [string]$ActionType
    )

    $eAction = ConvertTo-HtmlSafe $ActionType
    $eName = ConvertTo-HtmlSafe $StudentName
    $eClass = ConvertTo-HtmlSafe $StudentClass
    $eLogin = ConvertTo-HtmlSafe $StudentLogin
    $ePassword = ConvertTo-HtmlSafe $TempPassword

    $htmlBody = @"
<!DOCTYPE html>
<html>
<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
    <div style="max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
        <div style="background-color: #0078d4; color: white; padding: 20px; text-align: center;">
            <h2 style="margin: 0;">Szkolne Centrum IT - Microsoft 365</h2>
            <p style="margin: 5px 0 0 0;">Powiadomienie o realizacji zgłoszenia: $eAction</p>
        </div>
        <div style="padding: 20px;">
            <p>Dzień dobry,</p>
            <p>Zgłoszenie dotyczące ucznia <strong>$eName</strong> (Klasa: <strong>$eClass</strong>) zostało zrealizowane.</p>
            <div style="background-color: #f3f2f1; padding: 15px; border-left: 4px solid #0078d4; margin: 20px 0;">
                <p style="margin: 0 0 10px 0;"><strong>Login (UPN):</strong> <span style="font-family: monospace; color: #0078d4;">$eLogin</span></p>
                <p style="margin: 0;"><strong>Hasło startowe:</strong> <span style="font-family: monospace; color: #d83b01;">$ePassword</span></p>
            </div>
            <p style="font-size: 13px; color: #666;">Hasło jest jednorazowe. Uczeń zostanie poproszony o jego zmianę przy pierwszym logowaniu na stronie <strong>portal.office.com</strong>.</p>
        </div>
        <div style="background-color: #f8f9fa; padding: 10px; text-align: center; font-size: 11px; color: #888;">
            Wiadomość wygenerowana automatycznie przez Asystenta Administratora Microsoft 365.
        </div>
    </div>
</body>
</html>
"@

    $messageParams = @{
        message = @{
            subject      = "Dane dostępowe Microsoft 365 dla: $StudentName ($StudentClass)"
            body         = @{
                contentType = "html"
                content     = $htmlBody
            }
            toRecipients = @(
                @{ emailAddress = @{ address = $ToEmail } }
            )
        }
        saveToSentItems = $false
    }

    # Nadawcą jest zalogowany administrator (uprawnienie delegowane Mail.Send pozwala wysyłać tylko z własnej skrzynki).
    $sender = (Get-MgContext).Account
    try {
        Send-MgUserMail -UserId $sender -BodyParameter $messageParams -ErrorAction Stop
        Write-ColorLog -Message "Wysłano powiadomienie do $ToEmail ($StudentLogin)." -Level "info"
    }
    catch {
        Write-ColorLog -Message "Nie udało się wysłać powiadomienia do $ToEmail ($StudentLogin): $($_.Exception.Message)" -Level "warn"
    }
}

# ============================================================================
# OPCJA 6: GENERATOR RAPORTU HTML DO DRUKU
# ============================================================================

function Invoke-GeneratePrintableReport {
    Clear-Host
    Write-Host "--- GENEROWANIE RAPORTU HTML DO DRUKU (KARTECZKI DLA WYCHOWAWCÓW) ---" -ForegroundColor Cyan
    $csvPath = $script:ResultsCsv

    if (-not (Test-Path -Path $csvPath)) {
        Write-ColorLog -Message "Plik wyników '$csvPath' nie istnieje. Przetwórz najpierw zgłoszenia w Opcji 5." -Level "error"
        Wait-AnyKey; return
    }

    $records = @(Import-Csv -Path $csvPath -Delimiter ";" -Encoding utf8)
    if ($records.Count -eq 0) {
        Write-ColorLog -Message "Brak danych do wygenerowania raportu." -Level "warn"
        Wait-AnyKey; return
    }

    $tableRowsHtml = [System.Text.StringBuilder]::new()
    $cardsHtml = [System.Text.StringBuilder]::new()

    foreach ($r in $records) {
        # Wszystkie wartości pochodzą z formularza - kodujemy je przed wstawieniem do HTML.
        $id       = ConvertTo-HtmlSafe $r.ID_Zgloszenia
        $akcja    = ConvertTo-HtmlSafe $r.Akcja
        $klasa    = ConvertTo-HtmlSafe $r.Klasa
        $imie     = ConvertTo-HtmlSafe $r.Imie
        $nazwisko = ConvertTo-HtmlSafe $r.Nazwisko
        $login    = ConvertTo-HtmlSafe $r.Login
        $haslo    = ConvertTo-HtmlSafe $r.HasloStartowe
        $nauczyc  = ConvertTo-HtmlSafe $r.Nauczyciel
        $search   = ConvertTo-HtmlSafe "$($r.Klasa) $($r.Imie) $($r.Nazwisko) $($r.Login)"
        $badgeClass = if ($r.Akcja -match '(?i)reset') { "badge-reset" } else { "badge-new" }

        [void]$tableRowsHtml.AppendLine(@"
        <tr class="data-row" data-search="$search">
            <td>$id</td>
            <td><span class="badge $badgeClass">$akcja</span></td>
            <td><strong>$klasa</strong></td>
            <td>$imie $nazwisko</td>
            <td class="mono">$login</td>
            <td class="mono password-text">$haslo</td>
            <td>$nauczyc</td>
        </tr>
"@)

        [void]$cardsHtml.AppendLine(@"
        <div class="slip-card" data-search="$search">
            <div class="slip-header">
                <svg class="icon-svg" viewBox="0 0 24 24"><path fill="#0078d4" d="M12 3L1 9l4 2.18v6L12 21l7-3.82v-6l2-1.09V17h2V9L12 3zm6.82 6L12 12.72 5.18 9 12 5.28 18.82 9zM17 15.99l-5 2.73-5-2.73v-3.72L12 15l5-2.73v3.72z"/></svg>
                <span>DOSTĘP DO MICROSOFT 365 (EDUKACJA)</span>
            </div>
            <div class="slip-body">
                <div><strong>Uczeń:</strong> $imie $nazwisko</div>
                <div><strong>Klasa:</strong> $klasa</div>
                <div><strong>Login:</strong> <span class="mono">$login</span></div>
                <div><strong>Hasło startowe:</strong> <span class="mono">$haslo</span></div>
            </div>
            <div class="slip-footer">
                Zaloguj się na: https://portal.office.com | Zmień hasło przy 1. logowaniu.
            </div>
        </div>
"@)
    }

    $htmlContent = @"
<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Raport Dostępów Microsoft 365 - Asystent Szkolny</title>
    <style>
        :root {
            --primary: #0078d4;
            --success: #107c41;
            --warning: #d83b01;
            --bg: #f8f9fa;
            --text: #201f1e;
            --border: #edebe9;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: var(--bg);
            color: var(--text);
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
        }

        .header-panel {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 1.6px 3.6px 0 rgba(0,0,0,0.132);
            margin-bottom: 20px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 15px;
        }

        h1 {
            margin: 0;
            font-size: 22px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .icon-svg {
            width: 20px !important;
            height: 20px !important;
            max-width: 20px !important;
            max-height: 20px !important;
            vertical-align: middle;
        }

        .controls {
            display: flex;
            gap: 10px;
            align-items: center;
        }

        #searchInput {
            padding: 8px 12px;
            border: 1px solid #8a8886;
            border-radius: 4px;
            font-size: 14px;
            width: 260px;
        }

        .btn {
            background-color: var(--primary);
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 4px;
            cursor: pointer;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }

        .btn:hover {
            background-color: #106ebe;
        }

        .table-card {
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 1.6px 3.6px 0 rgba(0,0,0,0.132);
            margin-bottom: 30px;
            overflow-x: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 14px;
        }

        th, td {
            padding: 10px 12px;
            text-align: left;
            border-bottom: 1px solid var(--border);
        }

        th {
            background-color: #faf9f8;
            font-weight: 600;
        }

        .mono {
            font-family: Consolas, "Courier New", monospace;
        }

        .password-text {
            color: var(--warning);
            font-weight: bold;
        }

        .badge {
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: bold;
            text-transform: uppercase;
        }

        .badge-reset {
            background-color: #fff4ce;
            color: #797673;
        }

        .badge-new {
            background-color: #dff6dd;
            color: #107c41;
        }

        /* Widok karteczek dla wychowawców */
        .slips-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }

        .slip-card {
            border: 1.5px dashed #605e5c;
            border-radius: 6px;
            padding: 12px;
            background: white;
            box-sizing: border-box;
            break-inside: avoid;
            page-break-inside: avoid;
        }

        .slip-header {
            display: flex;
            align-items: center;
            gap: 8px;
            font-size: 11px;
            font-weight: bold;
            color: var(--primary);
            border-bottom: 1px solid #edebe9;
            padding-bottom: 5px;
            margin-bottom: 8px;
        }

        .slip-body {
            font-size: 13px;
            line-height: 1.5;
        }

        .slip-footer {
            margin-top: 8px;
            font-size: 10px;
            color: #605e5c;
            border-top: 1px dotted #ccc;
            padding-top: 4px;
            text-align: center;
        }

        /* Style wydruku */
        @media print {
            body {
                background: white;
                padding: 0;
            }

            .no-print {
                display: none !important;
            }

            .slips-grid {
                grid-template-columns: 1fr 1fr;
                gap: 10px;
            }

            .slip-card {
                border: 1px dashed #333;
                margin-bottom: 8px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header-panel no-print">
            <h1>
                <svg class="icon-svg" viewBox="0 0 24 24"><path fill="#0078d4" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
                Zrealizowane Zgłoszenia Microsoft 365
            </h1>
            <div class="controls">
                <input type="text" id="searchInput" placeholder="Filtruj (Klasa, Nazwisko, Login)..." onkeyup="filterView()">
                <button class="btn" onclick="window.print()">
                    <svg class="icon-svg" viewBox="0 0 24 24"><path fill="#ffffff" d="M19 8H5c-1.66 0-3 1.34-3 3v6h4v4h12v-4h4v-6c0-1.66-1.34-3-3-3zm-3 11H8v-5h8v5zm3-7c-.55 0-1-.45-1-1s.45-1 1-1 1 .45 1 1-.45 1-1 1zm-1-9H6v4h12V3z"/></svg>
                    Drukuj Karteczki
                </button>
            </div>
        </div>

        <div class="table-card no-print">
            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Akcja</th>
                        <th>Klasa</th>
                        <th>Uczeń</th>
                        <th>Login (UPN)</th>
                        <th>Hasło startowe</th>
                        <th>Nauczyciel</th>
                    </tr>
                </thead>
                <tbody id="tableBody">
                    $($tableRowsHtml.ToString())
                </tbody>
            </table>
        </div>

        <div class="slips-grid" id="slipsContainer">
            $($cardsHtml.ToString())
        </div>
    </div>

    <script>
        function filterView() {
            var query = document.getElementById('searchInput').value.toLowerCase();
            
            // Filtrowanie tabeli
            var rows = document.querySelectorAll('.data-row');
            rows.forEach(function(row) {
                var searchContext = row.getAttribute('data-search').toLowerCase();
                row.style.display = searchContext.includes(query) ? '' : 'none';
            });

            // Filtrowanie karteczek do wydruku
            var cards = document.querySelectorAll('.slip-card');
            cards.forEach(function(card) {
                var searchContext = card.getAttribute('data-search').toLowerCase();
                card.style.display = searchContext.includes(query) ? '' : 'none';
            });
        }
    </script>
</body>
</html>
"@

    try {
        $htmlContent | Out-File -FilePath $script:HtmlReport -Encoding utf8
        Set-PrivatePermissions -Path $script:HtmlReport
        Write-ColorLog -Message "Wygenerowano raport gotowy do wydruku: $($script:HtmlReport)" -Level "success"
        Write-ColorLog -Message "Raport zawiera hasła - po wydrukowaniu usuń go (Opcja 7)." -Level "warn"
    }
    catch {
        Write-ColorLog -Message "Błąd zapisu raportu HTML: $($_.Exception.Message)" -Level "error"
    }

    Wait-AnyKey
}

# ============================================================================
# OPCJA 7: USUWANIE PLIKÓW Z HASŁAMI
# ============================================================================

function Remove-PasswordFiles {
    Clear-Host
    Write-Host "--- USUWANIE PLIKÓW ZAWIERAJĄCYCH HASŁA STARTOWE ---" -ForegroundColor Cyan

    $files = @($script:ResultsCsv, $script:HtmlReport) | Where-Object { Test-Path -Path $_ }
    if ($files.Count -eq 0) {
        Write-ColorLog -Message "Brak plików z hasłami do usunięcia." -Level "info"
        Wait-AnyKey; return
    }

    Show-List -Items $files -Header "Pliki do usunięcia" -Format { param($x) $x }
    Write-Host "`nPlik 'przetworzone_id.txt' zostaje (nie zawiera haseł, chroni przed ponownym przetworzeniem)." -ForegroundColor DarkGray
    $confirm = Read-Host "Upewnij się, że karteczki są wydrukowane. Wpisz 'USUN', aby trwale usunąć pliki"
    if ($confirm -cne "USUN") {
        Write-ColorLog -Message "Operacja anulowana." -Level "info"
        Wait-AnyKey; return
    }

    foreach ($f in $files) {
        try {
            Remove-Item -Path $f -Force -ErrorAction Stop
            Write-ColorLog -Message "Usunięto: $f" -Level "success"
        }
        catch {
            Write-ColorLog -Message "Nie udało się usunąć $f : $($_.Exception.Message)" -Level "error"
        }
    }
    Wait-AnyKey
}

# ============================================================================
# OPCJA 8: BEZPIECZNE WYLOGOWANIE I ZAKOŃCZENIE PRACY
# ============================================================================

function Exit-AssistantSession {
    Clear-Host
    Write-Host "==================================================================" -ForegroundColor Cyan
    Write-Host "         ZAMYKANIE SESJI ASYSTENTA ADMINISTRATORA M365           " -ForegroundColor White
    Write-Host "==================================================================" -ForegroundColor Cyan

    if ($global:ExoConnected) {
        Write-ColorLog -Message "Zamykanie sesji Exchange Online..." -Level "info"
        try {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
            Write-ColorLog -Message "Rozłączono z Exchange Online." -Level "success"
        }
        catch {
            Write-ColorLog -Message "Błąd podczas rozłączania z Exchange Online: $($_.Exception.Message)" -Level "warn"
        }
    }

    if ($global:GraphConnected) {
        Write-ColorLog -Message "Zamykanie sesji Microsoft Graph..." -Level "info"
        try {
            Disconnect-MgGraph -ErrorAction Stop | Out-Null
            Write-ColorLog -Message "Rozłączono z Microsoft Graph API." -Level "success"
        }
        catch {
            Write-ColorLog -Message "Błąd podczas rozłączania z Microsoft Graph: $($_.Exception.Message)" -Level "warn"
        }
    }

    foreach ($f in @($script:ResultsCsv, $script:HtmlReport)) {
        if (Test-Path -Path $f) {
            Write-Host "UWAGA: plik z hasłami nadal istnieje: $f" -ForegroundColor Yellow
        }
    }

    Write-Host "`nBezpiecznie zakończono pracę. Do widzenia!" -ForegroundColor Green
    exit 0
}

# ============================================================================
# PĘTLA GŁÓWNA PROGRAMU
# ============================================================================

function Main {
    Initialize-WorkDir
    Write-ColorLog -Message "Start sesji$(if ($DryRun) { ' (TRYB PRÓBNY)' })." -Level "info"

    $authSuccess = Connect-SzkolnyTenant
    if (-not $authSuccess) {
        Write-ColorLog -Message "Nie udało się poprawnie nawiązać wymaganych połączeń. Skrypt zostanie przerwany." -Level "error"
        exit 1
    }

    $menuItems = @(
        "1. Generuj raport nieaktywnych skrzynek uczniów (Exchange Online)",
        "2. Zablokuj konta z raportu nieaktywnych (Microsoft Graph)",
        "3. Odbierz wszystkie licencje zablokowanym z raportu (Microsoft Graph)",
        "4. [POMOC] Instrukcja obsługi zgłoszeń z Microsoft Forms",
        "5. Obsługa zgłoszeń z Forms (Masowy reset i tworzenie kont z Excela)",
        "6. Generuj raport HTML do druku (z przetworzonych zgłoszeń Forms)",
        "7. Usuń pliki z hasłami startowymi (po wydruku)",
        "8. Wyloguj i zamknij aplikację"
    )

    while ($true) {
        $choice = Invoke-InteractiveMenu -Options $menuItems -Title "Wybierz moduł operacyjny:"

        switch ($choice) {
            0 { Invoke-InactiveMailboxReport }
            1 { Invoke-BlockInactiveAccounts }
            2 { Invoke-RevokeLicensesFromReport }
            3 { Show-FormsHelpScreen }
            4 { Invoke-ProcessFormsRequests }
            5 { Invoke-GeneratePrintableReport }
            6 { Remove-PasswordFiles }
            7 { Exit-AssistantSession }
        }
    }
}

# Uruchomienie aplikacji (pominięte przy dot-sourcingu, np. w testach: . ./szkolnyasystent365.ps1)
if ($MyInvocation.InvocationName -ne '.') {
    Main
}

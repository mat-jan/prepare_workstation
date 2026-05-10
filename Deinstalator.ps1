# ============================================================
#  SKRYPT DEINSTALACJI - MASZYNA TESTOWA
#  Uruchom jako Administrator: klik PPM -> "Uruchom jako administrator"
# ============================================================

# Kolory w konsoli
function Write-Step { param($msg) Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Skip { param($msg) Write-Host "    [--] $msg" -ForegroundColor Yellow }
function Write-Err  { param($msg) Write-Host "    [!!] $msg" -ForegroundColor Red }

# Sprawdz uprawnienia administratora
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Err "Uruchom skrypt jako Administrator!"
    pause; exit 1
}

Write-Host ""
Write-Host "============================================" -ForegroundColor Red
Write-Host "  DEINSTALATOR - MASZYNA TESTOWA" -ForegroundColor Red
Write-Host "  Usuwa: AnyDesk, ESET, Chrome, Intel DSA," -ForegroundColor Red
Write-Host "         Office oraz cofa ustawienia systemu" -ForegroundColor Red
Write-Host "============================================" -ForegroundColor Red
Write-Host ""

# ============================================================
#  FUNKCJA POMOCNICZA: znajdz i odinstaluj po nazwie z rejestru
# ============================================================
function Uninstall-ByName {
    param([string]$Pattern, [string]$DisplayLabel)

    Write-Step "Odinstalowywanie: $DisplayLabel"

    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )

    $found = $false
    foreach ($regPath in $regPaths) {
        if (-not (Test-Path $regPath)) { continue }
        $entries = Get-ChildItem $regPath
        foreach ($entry in $entries) {
            $props = $entry | Get-ItemProperty -ErrorAction SilentlyContinue
            if ($props.DisplayName -match $Pattern -and $props.UninstallString) {
                $found = $true
                $uninstall = $props.UninstallString
                Write-Host "    Znaleziono: $($props.DisplayName)" -ForegroundColor Magenta

                if ($uninstall -match "MsiExec") {
                    $guid = ($uninstall -replace "MsiExec.exe\s*/[IX]", "").Trim()
                    $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
                } else {
                    $proc = Start-Process "cmd.exe" -ArgumentList "/c $uninstall /S /silent /quiet /norestart" -Wait -PassThru
                }
                Write-OK "$($props.DisplayName) usunieto (kod: $($proc.ExitCode))"
            }
        }
    }

    if (-not $found) {
        Write-Skip "$DisplayLabel - nie znaleziono w systemie"
    }
}

# ============================================================
#  1. ANYDESK
# ============================================================
Write-Step "Odinstalowywanie: AnyDesk"
$anydesk = "C:\Program Files (x86)\AnyDesk\AnyDesk.exe"
if (Test-Path $anydesk) {
    $proc = Start-Process $anydesk -ArgumentList "--remove" -Wait -PassThru
    Write-OK "AnyDesk usunieto (kod: $($proc.ExitCode))"
} else {
    Uninstall-ByName -Pattern "AnyDesk" -DisplayLabel "AnyDesk (rejestr)"
}

# ============================================================
#  2. ESET
# ============================================================
Write-Step "Odinstalowywanie: ESET"

# Szukaj w rejestrze
$regPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)
$esetFound = $false
foreach ($regPath in $regPaths) {
    if (-not (Test-Path $regPath)) { continue }
    $entries = Get-ChildItem $regPath
    foreach ($entry in $entries) {
        $props = $entry | Get-ItemProperty -ErrorAction SilentlyContinue
        if ($props.DisplayName -match "ESET" -and $props.UninstallString) {
            $esetFound = $true
            $guid = ($props.UninstallString -replace "MsiExec.exe\s*/[IX]", "").Trim()
            Write-Host "    Znaleziono: $($props.DisplayName)" -ForegroundColor Magenta
            $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
            Write-OK "ESET usunieto (kod: $($proc.ExitCode))"
        }
    }
}
if (-not $esetFound) {
    Write-Skip "ESET - nie znaleziono w systemie"
}

# ============================================================
#  3. GOOGLE CHROME
# ============================================================
Uninstall-ByName -Pattern "Google Chrome" -DisplayLabel "Google Chrome"

# Czyszczenie danych Chrome jesli zostaly
$chromePaths = @(
    "$env:ProgramFiles\Google\Chrome",
    "${env:ProgramFiles(x86)}\Google\Chrome"
)
foreach ($p in $chromePaths) {
    if (Test-Path $p) {
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-OK "Usunieto folder: $p"
    }
}

# ============================================================
#  4. INTEL DRIVER AND SUPPORT ASSISTANT
# ============================================================
Uninstall-ByName -Pattern "Intel.*(Driver|DSA|Support Assistant)" -DisplayLabel "Intel Driver and Support Assistant"

# ============================================================
#  5. MICROSOFT OFFICE
# ============================================================
Write-Step "Odinstalowywanie: Microsoft Office"

$officeUninstallKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

$officeFound = @()
foreach ($regPath in $officeUninstallKeys) {
    if (Test-Path $regPath) {
        Get-ChildItem $regPath | ForEach-Object {
            $props = $_ | Get-ItemProperty -ErrorAction SilentlyContinue
            if ($props.DisplayName -match "Microsoft 365|Microsoft Office|Office 16|Office 15|Office 14" -and $props.UninstallString) {
                $officeFound += [PSCustomObject]@{ Name = $props.DisplayName; Uninstall = $props.UninstallString }
            }
        }
    }
}

if ($officeFound.Count -eq 0) {
    Write-Skip "Microsoft Office - nie znaleziono w systemie"
} else {
    foreach ($app in $officeFound) {
        Write-Host "    Odinstalowuje: $($app.Name)" -ForegroundColor Magenta
        $uninstall = $app.Uninstall
        if ($uninstall -match "MsiExec") {
            $guid = ($uninstall -replace "MsiExec.exe\s*/[IX]", "").Trim()
            $proc = Start-Process "msiexec.exe" -ArgumentList "/qn /x $guid /norestart" -Wait -PassThru
        } else {
            $proc = Start-Process "cmd.exe" -ArgumentList "/c $uninstall /quiet /norestart" -Wait -PassThru
        }
        Write-OK "Usunieto: $($app.Name) (kod: $($proc.ExitCode))"
    }
}

# Click-to-Run
$c2rSetup = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $c2rSetup) {
    Write-Host "    Usuwanie Office Click-to-Run..." -ForegroundColor Magenta
    $proc = Start-Process $c2rSetup -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365ProPlusRetail.16_pl-pl_x-none" -Wait -PassThru
    Write-OK "Office Click-to-Run usunieto (kod: $($proc.ExitCode))"
}

# Czyszczenie folderow Office
$officeFolders = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office",
    "C:\ProgramData\Microsoft\Office"
)
foreach ($folder in $officeFolders) {
    if (Test-Path $folder) {
        Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue
        Write-OK "Usunieto folder: $folder"
    }
}

# ============================================================
#  6. COFNIECIE USTAWIEN ZASILANIA
# ============================================================
Write-Step "Cofanie ustawien zasilania do domyslnych"

# Plan zasilania - przywroc Zrownowazony
$balancedGUID = (powercfg /list | Select-String "Zrownow|Balanced" | ForEach-Object { ($_ -split "\s+")[3] })
if ($balancedGUID) {
    powercfg /setactive $balancedGUID
    Write-OK "Plan zasilania: Zrownowazony przywrocony"
}

# Timeouty ekranu - przywroc do 10 minut
powercfg /change monitor-timeout-ac 10
powercfg /change monitor-timeout-dc 5
Write-OK "Wygaszenie ekranu: 10 min (AC), 5 min (DC)"

# Usypianie - przywroc do 30 minut
powercfg /change standby-timeout-ac 30
powercfg /change standby-timeout-dc 15
Write-OK "Usypianie: 30 min (AC), 15 min (DC)"

# Wygaszacz ekranu - przywroc
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveActive" -Value "1" -Force
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "ScreenSaveTimeOut" -Value "600" -Force
Write-OK "Wygaszacz ekranu: przywrocony (10 min)"

# Hibernacja - przywroc
powercfg /hibernate on
Write-OK "Hibernacja: przywrocona"

# ============================================================
#  7. COFNIECIE EXECUTION POLICY
# ============================================================
Write-Step "Przywracanie ExecutionPolicy do Restricted"
Set-ExecutionPolicy -ExecutionPolicy Restricted -Scope LocalMachine -Force
Write-OK "ExecutionPolicy: Restricted"

# ============================================================
#  PODSUMOWANIE
# ============================================================
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  GOTOWE! Maszyna testowa wyczyszczona." -ForegroundColor Green
Write-Host "  Mozesz uruchomic Setup-Firmowy.ps1" -ForegroundColor Green
Write-Host "  aby zaczac od nowa." -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host ""
pause